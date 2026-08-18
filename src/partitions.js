import { createHash } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir, mkdtemp, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { createInterface } from 'node:readline';
import { finished } from 'node:stream/promises';

export const sha256 = (value) => createHash('sha256').update(value).digest('hex');

export class PartitionError extends Error {
  constructor(code, message, options) {
    super(message, options);
    this.name = 'PartitionError';
    this.code = code;
  }
}

export class PartitionStore {
  static async create(options = {}) {
    const root = options.root ?? await mkdtemp(join(tmpdir(), 'excel-diff-'));
    if (options.root) await mkdir(root, { recursive: true });
    return new PartitionStore(root, options);
  }

  constructor(root, options) {
    this.root = root;
    this.maxOpenFiles = options.maxOpenFiles ?? 32;
    this.maxTempBytes = options.maxTempBytes ?? Infinity;
    this.filePrefix = options.filePrefix ?? 'partition-';
    this.bytesWritten = 0;
    this.streams = new Map();
    this.paths = new Set();
  }

  async append(record, depth = 0) {
    const bucket = record?.keyHash?.slice(depth * 2, depth * 2 + 2);
    if (!/^[0-9a-f]{2}$/.test(bucket)) {
      throw new PartitionError('PARTITION_INVALID', 'record has an invalid keyHash');
    }
    const line = `${JSON.stringify(record)}\n`;
    const bytes = Buffer.byteLength(line);
    if (this.bytesWritten + bytes > this.maxTempBytes) {
      throw new PartitionError('TEMP_LIMIT_EXCEEDED', 'temporary data exceeds resources.maxTempBytes');
    }
    const stream = await this.#streamFor(depth, bucket);
    try {
      await new Promise((resolve, reject) => stream.write(line, (error) => error ? reject(error) : resolve()));
    } catch (error) {
      this.streams.delete(stream.path);
      stream.destroy();
      await finished(stream).catch(() => {});
      throw error;
    }
    this.bytesWritten += bytes;
  }

  async close() {
    let firstError;
    for (const [path, stream] of [...this.streams]) {
      try {
        await this.#closeStream(path, stream);
      } catch (error) {
        firstError ??= error;
      }
    }
    if (firstError) throw firstError;
  }

  partitionPaths() {
    return [...this.paths].sort();
  }

  async cleanup() {
    let closeError;
    try {
      await this.close();
    } catch (error) {
      closeError = error;
    }
    await rm(this.root, { recursive: true, force: true });
    if (closeError) throw closeError;
  }

  async #streamFor(depth, bucket) {
    const path = join(this.root, `${this.filePrefix}${depth}-${bucket}.jsonl`);
    const existing = this.streams.get(path);
    if (existing) {
      this.streams.delete(path);
      this.streams.set(path, existing);
      return existing;
    }
    if (this.streams.size >= this.maxOpenFiles) {
      const [oldestPath, oldestStream] = this.streams.entries().next().value;
      await this.#closeStream(oldestPath, oldestStream);
    }
    const stream = createWriteStream(path, { flags: 'a', encoding: 'utf8' });
    this.streams.set(path, stream);
    this.paths.add(path);
    return stream;
  }

  async #closeStream(path, stream) {
    this.streams.delete(path);
    stream.end();
    await finished(stream);
  }
}

export async function* readPartition(path) {
  const input = createReadStream(path, { encoding: 'utf8' });
  const lines = createInterface({ input, crlfDelay: Infinity });
  let lineNumber = 0;
  try {
    for await (const line of lines) {
      lineNumber += 1;
      if (line === '') throw new PartitionError('PARTITION_INVALID', `blank line in partition at line ${lineNumber}`);
      try {
        yield JSON.parse(line);
      } catch (error) {
        if (error instanceof PartitionError) throw error;
        throw new PartitionError('PARTITION_INVALID', `invalid partition JSON at line ${lineNumber}`, { cause: error });
      }
    }
  } finally {
    lines.close();
    input.destroy();
  }
}

export async function repartition(path, depth, options = {}) {
  const maxPartitionBytes = options.maxPartitionBytes ?? Infinity;
  if ((await stat(path)).size <= maxPartitionBytes) return [path];

  const children = await PartitionStore.create({
    ...options,
    root: dirname(path),
    filePrefix: `${basename(path, '.jsonl')}-`
  });
  let firstHash;
  let mixedHashes = false;
  let operationError;
  try {
    for await (const record of readPartition(path)) {
      firstHash ??= record.keyHash;
      if (record.keyHash !== firstHash) mixedHashes = true;
      await children.append(record, depth);
    }
  } catch (error) {
    operationError = error;
  }
  try {
    await children.close();
  } catch (error) {
    operationError ??= error;
  }
  if (operationError) throw operationError;

  const childPaths = children.partitionPaths();
  if (!mixedHashes) {
    await Promise.all(childPaths.map((childPath) => rm(childPath, { force: true })));
    throw new PartitionError('HOT_KEY_TOO_LARGE', 'single key exceeds resources.maxPartitionBytes');
  }

  const boundedPaths = [];
  for (const childPath of childPaths) {
    if ((await stat(childPath)).size > maxPartitionBytes) {
      boundedPaths.push(...await repartition(childPath, depth + 1, options));
    } else {
      boundedPaths.push(childPath);
    }
  }
  await rm(path);
  return boundedPaths.sort();
}
