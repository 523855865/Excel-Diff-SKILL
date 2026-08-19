import { constants } from 'node:fs';
import { access, stat } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';

import { equalValues, encodeKey } from './normalize.js';
import { PartitionStore, readPartition, repartition, sha256 } from './partitions.js';
import { InputError, scanFileRows } from './read-xlsx.js';

export class CompareError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'CompareError';
    this.code = code;
  }
}

function ruleForColumn(spec, column) {
  const normalization = spec.normalization ?? {};
  return {
    emptyEqualsNull: false,
    caseSensitive: true,
    formulaMode: 'formula-and-cached-result',
    ...normalization,
    ...(normalization.columns?.[column] ?? {})
  };
}

function limitsFor(spec) {
  return {
    maxInputBytes: Infinity,
    maxRows: Infinity,
    maxCells: Infinity,
    maxTempBytes: Infinity,
    maxPartitionBytes: Infinity,
    maxRuntimeMs: Infinity,
    ...(spec.resources ?? {})
  };
}

function comparisonColumns(spec, columns) {
  const selected = spec.compareColumns === '*' ? columns : spec.compareColumns;
  return spec.mode.type === 'row' || spec.mode.type === 'multiset'
    ? selected
    : selected.filter((column) => !spec.mode.keyColumns.includes(column));
}

function rowKey(spec, compareColumns, row) {
  if (spec.mode.type === 'row') return [['number', row.rowNumber]];
  if (spec.mode.type === 'multiset') return compareColumns.map((column) => row.values[column]);
  return spec.mode.keyColumns.map((column) => row.values[column]);
}

function resourceError(code, resource) {
  return new CompareError(code, `comparison exceeds resources.${resource}`);
}

async function inputBytes(files) {
  const sizes = await Promise.all(files.map(async (file) => {
    try {
      await access(file.path, constants.R_OK);
      const metadata = await stat(file.path);
      if (!metadata.isFile()) throw new Error('not a file');
      return metadata.size;
    } catch {
      throw new InputError('INPUT_ERROR', `cannot read XLSX input for ${file.id}`);
    }
  }));
  return sizes.reduce((sum, size) => sum + size, 0);
}

async function processEntry(entry, spec, compareColumns, valueIndexes, sink, summary, enforceRuntime) {
  if (spec.mode.type === 'multiset') {
    const counts = Object.fromEntries(spec.files.map(({ id }) => [id, entry.rows.get(id)?.length ?? 0]));
    if (spec.files.every(({ id }) => counts[id] === counts[spec.baseline])) {
      summary.identicalKeys += 1;
      return;
    }
    summary.changedKeys += 1;
    await sink.onMultiset?.({
      values: entry.key,
      sheetName: entry.rows.values().next().value[0].sheetName,
      counts,
      baselineRelation: counts[spec.baseline] === 0 ? 'ADDED' : 'DELETED'
    });
    enforceRuntime();
    return;
  }

  const duplicateFiles = spec.files
    .filter(({ id }) => (entry.rows.get(id)?.length ?? 0) > 1)
    .map(({ id }) => id);
  if (duplicateFiles.length > 0) {
    if (spec.duplicateKeyPolicy === 'fail') {
      throw new CompareError('DUPLICATE_KEY', `duplicate business key in files ${duplicateFiles.join(', ')}`);
    }
    summary.duplicateKeys += 1;
    await sink.onDuplicate?.({ key: entry.key, files: duplicateFiles });
    enforceRuntime();
    return;
  }

  const presentFiles = spec.files.filter(({ id }) => entry.rows.has(id)).map(({ id }) => id);
  if (presentFiles.length !== spec.files.length) {
    summary.missingKeys += 1;
    await sink.onMissing?.({
      key: entry.key,
      sheetName: entry.rows.values().next().value[0].sheetName,
      presentFiles,
      missingFiles: spec.files.filter(({ id }) => !entry.rows.has(id)).map(({ id }) => id),
      baselineRelation: entry.rows.has(spec.baseline) ? 'DELETED' : 'ADDED'
    });
    enforceRuntime();
    return;
  }

  const baseline = entry.rows.get(spec.baseline)[0];
  if (spec.files.every(({ id }) => {
    const row = entry.rows.get(id)[0];
    return row.rowHash === baseline.rowHash && row.rowBytes === baseline.rowBytes;
  })) {
    enforceRuntime();
    summary.identicalKeys += 1;
    return;
  }
  let keyChanged = false;
  for (const column of compareColumns) {
    enforceRuntime();
    const index = valueIndexes.get(column);
    const rule = ruleForColumn(spec, column);
    if (spec.files.every(({ id }) => equalValues(baseline.values[index], entry.rows.get(id)[0].values[index], rule))) continue;
    keyChanged = true;
    await sink.onChanged?.({
      key: entry.key,
      sheetName: baseline.sheetName,
      column,
      files: Object.fromEntries(spec.files.map(({ id }) => {
        const row = entry.rows.get(id)[0];
        return [id, { value: row.values[index], rowNumber: row.rowNumber }];
      }))
    });
    enforceRuntime();
  }
  enforceRuntime();
  if (keyChanged) summary.changedKeys += 1;
  else summary.identicalKeys += 1;
}

export async function comparePartitioned(spec, sink = {}, options = {}) {
  const limits = limitsFor(spec);
  const now = options.now ?? (() => performance.now());
  const startedAt = now();
  const enforceRuntime = () => {
    if (now() - startedAt > limits.maxRuntimeMs) throw resourceError('RUNTIME_LIMIT_EXCEEDED', 'maxRuntimeMs');
  };
  let store;
  let failure;

  try {
    store = await PartitionStore.create({ maxTempBytes: limits.maxTempBytes });
    if (await inputBytes(spec.files) > limits.maxInputBytes) {
      throw resourceError('INPUT_LIMIT_EXCEEDED', 'maxInputBytes');
    }
    enforceRuntime();

    const baselineFile = spec.files.find(({ id }) => id === spec.baseline);
    let columns;
    let compareColumns;
    let valueColumns;
    let totalRowsScanned = 0;
    let matchedRows = 0;
    let invalidRows = 0;
    let totalCellsScanned = 0;

    const scan = async (file, standardColumns = null) => {
      const result = await scanFileRows(file, spec, standardColumns, async (row, rowColumns) => {
        compareColumns ??= comparisonColumns(spec, rowColumns);
        const selected = new Set(compareColumns);
        valueColumns ??= rowColumns.filter((column) => selected.has(column));
        const key = rowKey(spec, compareColumns, row);
        const keyEncoding = encodeKey(key);
        const keyHash = sha256(keyEncoding);
        if (spec.mode.type === 'multiset') {
          await store.append([keyHash, keyEncoding, row.fileId, row.sheetName, row.rowNumber]);
          return;
        }
        const values = valueColumns.map((column) => row.values[column]);
        const rowBytes = JSON.stringify(values);
        await store.append([
          keyHash,
          keyEncoding,
          row.fileId,
          row.sheetName,
          row.rowNumber,
          sha256(rowBytes),
          values
        ]);
      }, async ({ cellCount }) => {
        totalRowsScanned += 1;
        totalCellsScanned += cellCount;
        if (totalRowsScanned > limits.maxRows) throw resourceError('ROW_LIMIT_EXCEEDED', 'maxRows');
        if (totalCellsScanned > limits.maxCells) throw resourceError('CELL_LIMIT_EXCEEDED', 'maxCells');
        enforceRuntime();
        await options.onProgress?.({
          rowsScanned: totalRowsScanned,
          currentFile: file.id,
          bytesWritten: store.bytesWritten
        });
        enforceRuntime();
      });
      columns ??= result.columns;
      compareColumns ??= comparisonColumns(spec, columns);
      const selected = new Set(compareColumns);
      valueColumns ??= columns.filter((column) => selected.has(column));
      matchedRows += result.matchedRows;
      invalidRows += result.invalidRows;
    };

    await scan(baselineFile);
    for (const file of spec.files) {
      if (file.id !== baselineFile.id) await scan(file, columns);
    }
    await store.close();
    enforceRuntime();

    const summary = {
      files: spec.files.length,
      totalRowsScanned,
      matchedRows,
      identicalKeys: 0,
      changedKeys: 0,
      missingKeys: 0,
      duplicateKeys: 0,
      invalidRows
    };
    const valueIndexes = new Map(valueColumns.map((column, index) => [column, index]));
    for (const path of store.partitionPaths()) {
      enforceRuntime();
      const boundedPaths = await repartition(path, 1, {
        maxPartitionBytes: limits.maxPartitionBytes,
        maxTempBytes: limits.maxTempBytes,
        maxOpenFiles: 32,
        check: enforceRuntime
      });
      enforceRuntime();
      for (const boundedPath of boundedPaths) {
        const entries = new Map();
        for await (const tuple of readPartition(boundedPath)) {
          enforceRuntime();
          const [keyHash, keyEncoding, fileId, sheetName, rowNumber, rowHash, values] = tuple;
          const key = JSON.parse(keyEncoding);
          const row = { keyHash, keyEncoding, key, fileId, sheetName, rowNumber, rowHash, values, rowBytes: JSON.stringify(values) };
          let entry = entries.get(row.keyEncoding);
          if (!entry) {
            entry = { key: row.key, rows: new Map() };
            entries.set(row.keyEncoding, entry);
          }
          if (!entry.rows.has(row.fileId)) entry.rows.set(row.fileId, []);
          entry.rows.get(row.fileId).push(row);
        }
        for (const [, entry] of [...entries].sort(([left], [right]) => (left > right) - (left < right))) {
          enforceRuntime();
          await processEntry(entry, spec, compareColumns, valueIndexes, sink, summary, enforceRuntime);
        }
      }
    }

    enforceRuntime();
    return options.keepTemp ? { summary, tempDirectory: store.directory } : { summary };
  } catch (error) {
    failure = error;
    if (options.keepTemp && store && error && (typeof error === 'object' || typeof error === 'function')) {
      try {
        error.tempDirectory = store.directory;
      } catch {
        // Preserve non-extensible callback errors unchanged.
      }
    }
    throw error;
  } finally {
    if (store) {
      try {
        if (options.keepTemp) await store.close();
        else await store.cleanup();
      } catch (error) {
        if (!failure) throw error;
      }
    }
  }
}

export async function compare(spec) {
  const changed = [];
  const missing = [];
  const duplicates = [];
  const multiset = [];
  const { summary } = await comparePartitioned(spec, {
    onChanged: async (item) => changed.push(item),
    onMissing: async (item) => missing.push(item),
    onDuplicate: async (item) => duplicates.push(item),
    onMultiset: async (item) => multiset.push(item)
  });
  const byKey = (left, right) => {
    const leftKey = encodeKey(left.key);
    const rightKey = encodeKey(right.key);
    return (leftKey > rightKey) - (leftKey < rightKey);
  };
  changed.sort(byKey);
  missing.sort(byKey);
  duplicates.sort(byKey);
  const result = { summary, changed, missing, duplicates };
  if (spec.mode.type === 'multiset') {
    multiset.sort((left, right) => byKey({ key: left.values }, { key: right.values }));
    result.multiset = multiset;
  }
  return result;
}
