import assert from 'node:assert/strict';
import { access, mkdir, rm, stat, writeFile } from 'node:fs/promises';
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
  const openPaths = [...store.streams.keys()];
  assert.equal(store.streams.size, 2);
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
