import { randomUUID } from 'node:crypto';
import { createWriteStream } from 'node:fs';
import { lstat, mkdir, readdir, rename, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { finished } from 'node:stream/promises';

export function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[,"\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function csv(rows) {
  return `${rows.map((row) => row.map(csvCell).join(',')).join('\n')}\n`;
}

function safeText(text) {
  return /^(?:[=+\-@\t\r\n]|json:)/.test(text) ? `json:${JSON.stringify(text)}` : text;
}

function untypedValue([type, value]) {
  if (type === 'blank') return '';
  if (type === 'formula' && Array.isArray(value)) {
    return { formula: value[0], result: untypedValue(value[1]) };
  }
  if (type === 'hyperlink') return { ...value, text: untypedValue(value.text) };
  return value;
}

function safeValue(value) {
  return safeText(value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value ?? ''));
}

function displayKey(key) {
  const values = key.map(untypedValue);
  return safeValue(values.length === 1 ? values[0] : values);
}

function runId() {
  return `${new Date().toISOString().replace(/\.\d{3}Z$/, 'Z').replace(/[-:]/g, '')}-${randomUUID().slice(0, 8)}`;
}

function writeChunk(stream, chunk) {
  return new Promise((resolve, reject) => {
    stream.write(chunk, (error) => error ? reject(error) : resolve());
  });
}

export async function createReportWriter(spec, options = {}) {
  const initialRunId = runId();
  const staging = join(spec.output.directory, `.${initialRunId}.tmp`);
  const fileIds = spec.files.map(({ id: fileId }) => fileId);
  const definitions = spec.mode?.type === 'multiset' ? {
    multiset: ['values', 'sheet', ...fileIds.map((fileId) => safeText(`${fileId}.count`)), 'baselineRelation']
  } : {
    changed: ['key', 'sheet', 'column', ...fileIds.flatMap((fileId) => [safeText(`${fileId}.value`), safeText(`${fileId}.type`), safeText(`${fileId}.row`)])],
    missing: ['key', 'sheet', 'presentFiles', 'missingFiles', 'baselineRelation'],
    duplicates: ['key', 'files']
  };
  const artifacts = Object.fromEntries(Object.keys(definitions).map((name) => [
    name, name === 'duplicates' ? 'duplicate-keys.csv' : `${name}.csv`
  ]));
  const streams = new Map();
  const createStream = options.createStream ?? ((path) => createWriteStream(path, { encoding: 'utf8' }));
  let firstStreamError;
  let pending = Promise.resolve();
  let state = 'active';

  const closeStreams = async (destroy = false) => {
    for (const stream of streams.values()) {
      if (destroy) stream.destroy();
      else stream.end();
      await finished(stream).catch((error) => {
        if (!destroy) throw error;
      });
    }
  };
  const enqueue = (name, row) => {
    const operation = pending.then(() => {
      const stream = streams.get(name);
      if (!stream) return;
      if (firstStreamError) throw firstStreamError;
      return writeChunk(stream, csv([row]));
    });
    pending = operation;
    return operation;
  };
  const scrubStaging = async () => {
    let firstError;
    let names = [];
    try {
      names = await readdir(staging);
    } catch (error) {
      if (error.code !== 'ENOENT') firstError = error;
    }
    for (const name of names) {
      try {
        await rm(join(staging, name), { recursive: true, force: true });
      } catch (error) {
        firstError ??= error;
      }
    }
    return firstError;
  };
  const publish = async (summary) => {
    let id = initialRunId;
    while (true) {
      const directory = join(spec.output.directory, id);
      const publishedSummary = { ...summary, runId: id };
      await writeFile(join(staging, 'summary.json'), `${JSON.stringify(publishedSummary, null, 2)}\n`, 'utf8');
      try {
        await lstat(directory);
        id = runId();
        continue;
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
      try {
        await rename(staging, directory);
        return { directory, summary: publishedSummary };
      } catch (renameError) {
        try {
          await lstat(directory);
          id = runId();
        } catch (error) {
          if (error.code === 'ENOENT') throw renameError;
          throw error;
        }
      }
    }
  };

  await mkdir(spec.output.directory, { recursive: true });
  await mkdir(staging);
  pending = (async () => {
    for (const [name, header] of Object.entries(definitions)) {
      const stream = createStream(join(staging, artifacts[name]), { encoding: 'utf8' });
      stream.on('error', (error) => { firstStreamError ??= error; });
      streams.set(name, stream);
      await writeChunk(stream, csv([header]));
    }
  })();
  pending.catch(() => {});

  return {
    onChanged: (entry) => enqueue('changed', [
      displayKey(entry.key), safeText(entry.sheetName), safeText(entry.column),
      ...fileIds.flatMap((fileId) => {
        const value = entry.files[fileId].value;
        return [safeValue(untypedValue(value)), safeText(value[0]), entry.files[fileId].rowNumber];
      })
    ]),
    onMissing: (entry) => enqueue('missing', [
      displayKey(entry.key), safeText(entry.sheetName), JSON.stringify(entry.presentFiles), JSON.stringify(entry.missingFiles), entry.baselineRelation
    ]),
    onDuplicate: (entry) => enqueue('duplicates', [displayKey(entry.key), JSON.stringify(entry.files)]),
    onMultiset: (entry) => enqueue('multiset', [
      safeValue(entry.values.map(untypedValue)), safeText(entry.sheetName),
      ...fileIds.map((fileId) => entry.counts[fileId]), entry.baselineRelation
    ]),
    async complete(summary) {
      if (state !== 'active') throw new Error('report writer is not active');
      await pending;
      if (firstStreamError) throw firstStreamError;
      await closeStreams();
      const completed = await publish({ ...summary, status: 'COMPLETED', artifacts });
      state = 'completed';
      return completed;
    },
    async abort(error) {
      if (state !== 'active') throw new Error('report writer is not active');
      state = 'aborting';
      try {
        await pending.catch(() => {});
        await closeStreams(true);
        const scrubError = await scrubStaging();
        if (scrubError) throw scrubError;
        await rm(staging, { recursive: true, force: true });
        await mkdir(staging);
        const diskFull = error?.code === 'ENOSPC' || error?.code === 'EDQUOT';
        const code = diskFull ? 'DISK_FULL' : typeof error?.code === 'string' && /^[A-Z][A-Z0-9_]*$/.test(error.code)
          ? error.code
          : 'INTERNAL_ERROR';
        const aborted = await publish({
          status: 'FAILED', code, message: 'comparison failed',
          ...(typeof error?.tempDirectory === 'string' ? { tempDirectory: error.tempDirectory } : {})
        });
        state = 'aborted';
        return aborted;
      } catch (abortError) {
        await scrubStaging();
        await rm(staging, { recursive: true, force: true }).catch(() => {});
        if (abortError && typeof abortError === 'object' && abortError.cause === undefined) {
          try { abortError.cause = error; } catch {}
        }
        throw abortError;
      }
    }
  };
}

export async function writeReport(spec, result) {
  const writer = await createReportWriter(spec);
  try {
    const changed = [...result.changed].sort((left, right) => {
      const leftKey = JSON.stringify(left.key);
      const rightKey = JSON.stringify(right.key);
      return (leftKey > rightKey) - (leftKey < rightKey)
        || (left.column > right.column) - (left.column < right.column);
    });
    const multiset = [...(result.multiset ?? [])].sort((left, right) => {
      const leftValues = JSON.stringify(left.values);
      const rightValues = JSON.stringify(right.values);
      return (leftValues > rightValues) - (leftValues < rightValues);
    });
    for (const entry of changed) await writer.onChanged(entry);
    for (const entry of result.missing) await writer.onMissing(entry);
    for (const entry of result.duplicates ?? []) await writer.onDuplicate(entry);
    for (const entry of multiset) await writer.onMultiset(entry);
    return await writer.complete(result.summary);
  } catch (error) {
    await writer.abort(error);
    throw error;
  }
}
