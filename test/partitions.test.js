import assert from 'node:assert/strict';
import { access, chmod, mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import test from 'node:test';

import {
  PartitionError,
  PartitionStore,
  readPartition,
  repartition,
  sha256
} from '../src/partitions.js';

const record = (key, extra = {}) => ({ key, keyHash: sha256(JSON.stringify(key)), ...extra });

async function collect(path) {
  const records = [];
  for await (const item of readPartition(path)) records.push(item);
  return records;
}

async function directoryBytes(path) {
  let bytes = 0;
  for (const entry of await readdir(path, { withFileTypes: true })) {
    const entryPath = join(path, entry.name);
    bytes += entry.isDirectory() ? await directoryBytes(entryPath) : (await stat(entryPath)).size;
  }
  return bytes;
}

test('sha256 returns the known lowercase 64-character digest', () => {
  assert.equal(sha256('abc'), 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
  assert.match(sha256('anything'), /^[0-9a-f]{64}$/);
});

test('writes records to sorted default-depth partitions and preserves them', async (t) => {
  const store = await PartitionStore.create();
  t.after(() => store.cleanup());

  await store.append(record('alpha'));
  await store.append(record('beta'));
  await store.append(record('alpha'));
  await store.close();

  const paths = store.partitionPaths();
  assert.deepEqual(paths, [...paths].sort());
  assert.equal(store.bytesWritten > 0, true);
  const records = (await Promise.all(paths.map(collect))).flat();
  assert.deepEqual(records.map(({ key }) => key).sort(), ['alpha', 'alpha', 'beta']);
  assert.equal(records.every(({ keyHash }) => keyHash.length === 64), true);
});

test('reopens least-recently-used buckets when maxOpenFiles is two', async (t) => {
  const keysByBucket = new Map();
  for (let index = 0; keysByBucket.size < 3; index += 1) {
    const item = record(`key-${index}`);
    if (!keysByBucket.has(item.keyHash.slice(0, 2))) keysByBucket.set(item.keyHash.slice(0, 2), item);
  }
  const store = await PartitionStore.create({ maxOpenFiles: 2 });
  t.after(() => store.cleanup());
  const [a, b, c] = keysByBucket.values();

  for (const item of [a, b, a, c]) await store.append(item);
  const openPaths = store.openPartitionPaths();
  assert.equal(store.openFileCount, 2);
  assert.equal(store.streams, undefined);
  assert.equal(openPaths.some((path) => path.endsWith(`0-${a.keyHash.slice(0, 2)}.jsonl`)), true);
  assert.equal(openPaths.some((path) => path.endsWith(`0-${c.keyHash.slice(0, 2)}.jsonl`)), true);
  assert.equal(openPaths.some((path) => path.endsWith(`0-${b.keyHash.slice(0, 2)}.jsonl`)), false);

  await store.append(b);
  await store.close();

  assert.equal(store.partitionPaths().length, 3);
  assert.deepEqual(
    (await Promise.all(store.partitionPaths().map(collect))).flat().map(({ key }) => key).sort(),
    [a.key, b.key, a.key, c.key, b.key].sort()
  );
});

test('cleanup preserves a caller-owned root and unrelated files', async (t) => {
  const root = await mkdtemp(join(tmpdir(), 'excel-diff-owned-test-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const sentinel = join(root, 'keep.txt');
  await writeFile(sentinel, 'keep');
  const store = await PartitionStore.create({ root });
  await store.append(record('owned-root'));
  const paths = store.partitionPaths();
  await store.cleanup();

  assert.equal(await readFile(sentinel, 'utf8'), 'keep');
  await access(root);
  for (const path of paths) await assert.rejects(() => access(path));
});

test('caller-root filename collisions remain untouched', async (t) => {
  const root = await mkdtemp(join(tmpdir(), 'excel-diff-collision-test-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const item = record('collision');
  const collision = join(root, `partition-0-${item.keyHash.slice(0, 2)}.jsonl`);
  await writeFile(collision, 'caller-owned\n');
  const store = await PartitionStore.create({ root });

  await store.append(item);
  await store.close();
  assert.notEqual(store.partitionPaths()[0], collision);
  await store.cleanup();

  assert.equal(await readFile(collision, 'utf8'), 'caller-owned\n');
  await access(root);
});

test('cleanup can be retried after run-directory deletion fails', async (t) => {
  const parent = await mkdtemp(join(tmpdir(), 'excel-diff-cleanup-retry-test-'));
  t.after(async () => {
    await chmod(parent, 0o700).catch(() => {});
    await rm(parent, { recursive: true, force: true });
  });
  const store = await PartitionStore.create({ root: parent });
  await store.append(record('cleanup-retry'));
  await store.close();
  const runRoot = dirname(store.partitionPaths()[0]);

  await chmod(parent, 0o500);
  try {
    await assert.rejects(() => store.cleanup());
  } finally {
    await chmod(parent, 0o700);
  }
  await store.cleanup();

  await assert.rejects(() => access(runRoot));
  await access(parent);
});

test('rejects before exceeding maxTempBytes and cleanup removes the run root', async () => {
  const first = record('first');
  const lineBytes = Buffer.byteLength(`${JSON.stringify(first)}\n`);
  const store = await PartitionStore.create({ maxTempBytes: lineBytes });

  await store.append(first);
  await assert.rejects(
    () => store.append(record('second')),
    (error) => error instanceof PartitionError && error.code === 'TEMP_LIMIT_EXCEEDED'
  );
  await store.close();
  const [path] = store.partitionPaths();
  assert.deepEqual(await collect(path), [first]);
  const root = dirname(path);
  await store.cleanup();
  await assert.rejects(() => access(root));
});

test('readPartition rejects malformed JSON and unexpected blank lines', async (t) => {
  const directory = join(tmpdir(), `excel-diff-invalid-${process.pid}-${Date.now()}`);
  await mkdir(directory);
  t.after(() => rm(directory, { recursive: true, force: true }));
  const malformed = join(directory, 'malformed.jsonl');
  const blank = join(directory, 'blank.jsonl');
  await writeFile(malformed, '{"ok":true}\nnot-json\n');
  await writeFile(blank, '{"ok":true}\n\n');

  for (const path of [malformed, blank]) {
    await assert.rejects(
      async () => { for await (const _ of readPartition(path)) void _; },
      (error) => error instanceof PartitionError && error.code === 'PARTITION_INVALID'
    );
  }
});

test('recursively splits an oversized mixed-key parent into bounded children', async (t) => {
  const store = await PartitionStore.create({ maxOpenFiles: 2 });
  t.after(() => store.cleanup());
  const records = Array.from({ length: 12 }, (_, index) => ({
    key: `key-${index}`,
    keyHash: `aa${(index % 4).toString(16).padStart(2, '0')}${String(index).padStart(60, '0')}`,
    payload: 'x'.repeat(80)
  }));
  for (const item of records) await store.append(item);
  await store.close();
  const [parent] = store.partitionPaths();

  const children = await repartition(parent, 1, {
    maxOpenFiles: 2,
    maxPartitionBytes: 600,
    maxTempBytes: 1024 * 1024
  });

  assert.equal(children.length > 1, true);
  for (const path of children) assert.equal((await stat(path)).size <= 600, true);
  assert.deepEqual((await Promise.all(children.map(collect))).flat().map(({ key }) => key).sort(), records.map(({ key }) => key).sort());
  await assert.rejects(() => access(parent));
});

test('failed repartition removes staged children and can be retried without duplicates', async (t) => {
  const store = await PartitionStore.create();
  t.after(() => store.cleanup());
  const records = Array.from({ length: 8 }, (_, index) => ({
    key: `retry-${index}`,
    keyHash: `aa${(index % 4).toString(16).padStart(2, '0')}${String(index).padStart(60, '0')}`,
    payload: 'x'.repeat(80)
  }));
  for (const item of records) await store.append(item);
  await store.close();
  const [parent] = store.partitionPaths();
  const parentSize = (await stat(parent)).size;
  const oneLine = Buffer.byteLength(`${JSON.stringify(records[0])}\n`);

  await assert.rejects(
    () => repartition(parent, 1, { maxPartitionBytes: 600, maxTempBytes: parentSize + oneLine }),
    (error) => error instanceof PartitionError && error.code === 'TEMP_LIMIT_EXCEEDED'
  );
  await access(parent);
  assert.equal((await readdir(dirname(parent))).some((name) => name.startsWith('.excel-diff-repartition-')), false);

  const children = await repartition(parent, 1, { maxPartitionBytes: 600, maxTempBytes: parentSize * 3 });
  assert.deepEqual(
    (await Promise.all(children.map(collect))).flat().map(({ key }) => key).sort(),
    records.map(({ key }) => key).sort()
  );
});

test('charges parent, siblings, and nested staging to one live temp budget', async (t) => {
  const store = await PartitionStore.create();
  t.after(() => store.cleanup());
  const parentRecords = Array.from({ length: 6 }, (_, index) => ({
    key: `nested-${index}`,
    keyHash: `aabb${(index % 3).toString(16).padStart(2, '0')}${String(index).padStart(58, '0')}`,
    payload: 'x'.repeat(80)
  }));
  for (const item of parentRecords) await store.append(item);
  await store.append({ key: 'sibling', keyHash: `cc${'00'.repeat(31)}`, payload: 'x'.repeat(80) });
  await store.close();
  const parent = store.partitionPaths().find((path) => path.endsWith('0-aa.jsonl'));
  const root = dirname(parent);
  const initialBytes = await directoryBytes(root);
  const maxTempBytes = initialBytes + (await stat(parent)).size;

  await assert.rejects(
    () => repartition(parent, 1, { maxPartitionBytes: 500, maxTempBytes }),
    (error) => error instanceof PartitionError && error.code === 'TEMP_LIMIT_EXCEEDED'
  );
  assert.equal(await directoryBytes(root) <= maxTempBytes, true);
  assert.equal((await readdir(root)).some((name) => name.startsWith('.excel-diff-repartition-')), false);
  await access(parent);
  await access(store.partitionPaths().find((path) => path.endsWith('0-cc.jsonl')));
});

test('reports HOT_KEY_TOO_LARGE when all 32 hash bytes are exhausted', async (t) => {
  const store = await PartitionStore.create();
  t.after(() => store.cleanup());
  const records = [0, 1].map((suffix) => ({
    key: `last-byte-${suffix}`,
    keyHash: `${'ab'.repeat(31)}${suffix.toString(16).padStart(2, '0')}`,
    payload: 'x'.repeat(40)
  }));
  for (const item of records) await store.append(item);
  await store.close();
  const [parent] = store.partitionPaths();
  const maxPartitionBytes = Buffer.byteLength(`${JSON.stringify(records[0])}\n`) - 1;

  await assert.rejects(
    () => repartition(parent, 1, { maxPartitionBytes }),
    (error) => error instanceof PartitionError && error.code === 'HOT_KEY_TOO_LARGE'
  );
  await access(parent);
  assert.equal((await readdir(dirname(parent))).some((name) => name.startsWith('.excel-diff-repartition-')), false);
});

test('retains an oversized hot-key parent and reports HOT_KEY_TOO_LARGE', async (t) => {
  const store = await PartitionStore.create();
  t.after(() => store.cleanup());
  const item = record('hot', { payload: 'x'.repeat(80) });
  for (let index = 0; index < 4; index += 1) await store.append(item);
  await store.close();
  const [parent] = store.partitionPaths();

  await assert.rejects(
    () => repartition(parent, 1, { maxPartitionBytes: 100 }),
    (error) => error instanceof PartitionError && error.code === 'HOT_KEY_TOO_LARGE'
  );
  await access(parent);
});
