import { createHash } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir, mkdtemp, readdir, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve, sep } from 'node:path';
import { createInterface } from 'node:readline';
import { finished } from 'node:stream/promises';

const budgetOption = Symbol('budget');
const trackerOption = Symbol('tracker');
const pathTrackers = new Map();

const keyHashOf = (record) => Array.isArray(record) ? record[0] : record?.keyHash;

export const sha256 = (value) => createHash('sha256').update(value).digest('hex');

export class PartitionError extends Error {
  constructor(code, message, options) {
    super(message, options);
    this.name = 'PartitionError';
    this.code = code;
  }
}

class LiveByteBudget {
  constructor(limit, liveBytes = 0) {
    this.limit = limit;
    this.liveBytes = liveBytes;
  }

  reserve(bytes) {
    if (this.liveBytes + bytes > this.limit) {
      throw new PartitionError('TEMP_LIMIT_EXCEEDED', 'temporary data exceeds resources.maxTempBytes');
    }
    this.liveBytes += bytes;
  }

  release(bytes) {
    this.liveBytes -= bytes;
  }
}

function forgetPath(path, tracker) {
  tracker.files.delete(path);
  if (pathTrackers.get(path) === tracker) pathTrackers.delete(path);
}

function forgetTree(root, tracker) {
  const prefix = `${root}${sep}`;
  for (const path of [...tracker.files]) {
    if (path.startsWith(prefix)) forgetPath(path, tracker);
  }
  for (const path of [...tracker.directories]) {
    if (path === root || path.startsWith(prefix)) tracker.directories.delete(path);
  }
}

async function directoryBytes(root) {
  let bytes = 0;
  for (const entry of await readdir(root, { withFileTypes: true })) {
    const path = join(root, entry.name);
    bytes += entry.isDirectory() ? await directoryBytes(path) : (await stat(path)).size;
  }
  return bytes;
}

async function trackedBytes(tracker) {
  let bytes = 0;
  for (const path of tracker.files) {
    try {
      bytes += (await stat(path)).size;
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  return bytes;
}

export class PartitionStore {
  #budget;
  #filePrefix;
  #maxOpenFiles;
  #paths = new Set();
  #root;
  #streams = new Map();
  #tracker;

  static async create(options = {}) {
    const parent = resolve(options.root ?? tmpdir());
    if (options.root) await mkdir(parent, { recursive: true });
    return new PartitionStore(await mkdtemp(join(parent, 'excel-diff-')), options);
  }

  constructor(root, options) {
    this.#root = root;
    this.#maxOpenFiles = options.maxOpenFiles ?? 32;
    this.#filePrefix = options.filePrefix ?? 'partition-';
    this.#budget = options[budgetOption] ?? new LiveByteBudget(options.maxTempBytes ?? Infinity);
    this.#tracker = options[trackerOption] ?? { files: new Set(), directories: new Set() };
    this.bytesWritten = 0;
  }

  get openFileCount() {
    return this.#streams.size;
  }

  get directory() {
    return resolve(this.#root);
  }

  openPartitionPaths() {
    return [...this.#streams.keys()].sort();
  }

  async append(record, depth = 0) {
    const bucket = keyHashOf(record)?.slice(depth * 2, depth * 2 + 2);
    if (!/^[0-9a-f]{2}$/.test(bucket)) {
      throw new PartitionError('PARTITION_INVALID', 'record has an invalid keyHash');
    }
    const line = `${JSON.stringify(record)}\n`;
    const bytes = Buffer.byteLength(line);
    this.#budget.reserve(bytes);
    let stream;
    try {
      stream = await this.#streamFor(depth, bucket);
      await new Promise((resolve, reject) => stream.write(line, (error) => error ? reject(error) : resolve()));
    } catch (error) {
      this.#budget.release(bytes);
      if (stream) {
        this.#streams.delete(stream.path);
        stream.destroy();
        await finished(stream).catch(() => {});
      }
      throw error;
    }
    this.bytesWritten += bytes;
  }

  async close() {
    let firstError;
    for (const [path, stream] of [...this.#streams]) {
      try {
        await this.#closeStream(path, stream);
      } catch (error) {
        firstError ??= error;
      }
    }
    if (firstError) throw firstError;
  }

  partitionPaths() {
    return [...this.#paths].sort();
  }

  async cleanup() {
    let closeError;
    try {
      await this.close();
    } catch (error) {
      closeError = error;
    }
    let cleanupError;
    try {
      await rm(this.#root, { recursive: true, force: true });
    } catch (error) {
      cleanupError = error;
    }
    if (!cleanupError) forgetTree(this.#root, this.#tracker);
    if (closeError) throw closeError;
    if (cleanupError) throw cleanupError;
  }

  async #streamFor(depth, bucket) {
    const path = join(this.#root, `${this.#filePrefix}${depth}-${bucket}.jsonl`);
    const existing = this.#streams.get(path);
    if (existing) {
      this.#streams.delete(path);
      this.#streams.set(path, existing);
      return existing;
    }
    if (this.#streams.size >= this.#maxOpenFiles) {
      const [oldestPath, oldestStream] = this.#streams.entries().next().value;
      await this.#closeStream(oldestPath, oldestStream);
    }
    const stream = createWriteStream(path, { flags: 'a', encoding: 'utf8' });
    this.#streams.set(path, stream);
    this.#paths.add(path);
    this.#tracker.files.add(path);
    pathTrackers.set(path, this.#tracker);
    return stream;
  }

  async #closeStream(path, stream) {
    this.#streams.delete(path);
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
        throw new PartitionError('PARTITION_INVALID', `invalid partition JSON at line ${lineNumber}`, { cause: error });
      }
    }
  } finally {
    lines.close();
    input.destroy();
  }
}

async function splitPartition(path, depth, context) {
  context.check();
  const parentSize = (await stat(path)).size;
  if (parentSize <= context.maxPartitionBytes) return [path];
  if (depth >= 32) {
    throw new PartitionError('HOT_KEY_TOO_LARGE', 'single key exceeds resources.maxPartitionBytes');
  }

  const staging = await mkdtemp(join(context.stagingRoot, 'split-'));
  context.tracker.directories.add(staging);
  let children;
  try {
    children = await PartitionStore.create({
      root: staging,
      maxOpenFiles: context.maxOpenFiles,
      [budgetOption]: context.budget,
      [trackerOption]: context.tracker
    });
    let firstHash;
    let mixedHashes = false;
    for await (const record of readPartition(path)) {
      context.check();
      const keyHash = keyHashOf(record);
      firstHash ??= keyHash;
      if (keyHash !== firstHash) mixedHashes = true;
      await children.append(record, depth);
      context.check();
    }
    await children.close();
    context.check();
    if (!mixedHashes) {
      throw new PartitionError('HOT_KEY_TOO_LARGE', 'single key exceeds resources.maxPartitionBytes');
    }

    const boundedPaths = [];
    for (const childPath of children.partitionPaths()) {
      context.check();
      boundedPaths.push(...await splitPartition(childPath, depth + 1, context));
    }
    await rm(path);
    context.budget.release(parentSize);
    forgetPath(path, context.tracker);
    return boundedPaths;
  } catch (error) {
    await children?.close().catch(() => {});
    throw error;
  }
}

export async function repartition(path, depth, options = {}) {
  const check = options.check ?? (() => {});
  check();
  const parentSize = (await stat(path)).size;
  const maxPartitionBytes = options.maxPartitionBytes ?? Infinity;
  if (parentSize <= maxPartitionBytes) return [path];
  if (depth >= 32) {
    throw new PartitionError('HOT_KEY_TOO_LARGE', 'single key exceeds resources.maxPartitionBytes');
  }
  const tracker = pathTrackers.get(path);
  const runRoot = dirname(path);
  const liveBytes = tracker ? await trackedBytes(tracker) : await directoryBytes(runRoot);
  const stagingRoot = await mkdtemp(join(runRoot, '.excel-diff-repartition-'));
  const activeTracker = tracker ?? { files: new Set([path]), directories: new Set() };
  activeTracker.directories.add(stagingRoot);
  const context = {
    budget: new LiveByteBudget(options.maxTempBytes ?? Infinity, liveBytes),
    maxOpenFiles: options.maxOpenFiles ?? 32,
    maxPartitionBytes,
    stagingRoot,
    tracker: activeTracker,
    check
  };
  try {
    return await splitPartition(path, depth, context);
  } catch (error) {
    const stagingBytes = await directoryBytes(stagingRoot).catch(() => 0);
    await rm(stagingRoot, { recursive: true, force: true });
    context.budget.release(stagingBytes);
    forgetTree(stagingRoot, activeTracker);
    throw error;
  }
}
