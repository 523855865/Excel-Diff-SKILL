import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { access, mkdir, readFile, readdir, rm } from 'node:fs/promises';
import { join } from 'node:path';
import test from 'node:test';

import { compare, comparePartitioned, CompareError } from '../src/compare.js';
import { encodeKey } from '../src/normalize.js';
import { sha256 } from '../src/partitions.js';
import { makeTempDir, writeWorkbook } from './helpers.js';

function spec(directory, files, overrides = {}) {
  return {
    baseline: 'A',
    files,
    sheet: { name: '人员', headerRow: 1 },
    mode: { keyColumns: ['编号'] },
    compareColumns: '*',
    columnAliases: { 姓名: ['Name'] },
    filters: [{ column: '状态', operator: 'eq', value: '在职' }],
    normalization: {
      emptyEqualsNull: false,
      caseSensitive: true,
      formulaMode: 'formula-and-cached-result',
      columns: {}
    },
    duplicateKeyPolicy: 'report',
    output: { directory },
    ...overrides
  };
}

async function filesFor(t, rowsById, order = ['C', 'A', 'B']) {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const paths = {};
  for (const [id, rows] of Object.entries(rowsById)) {
    paths[id] = join(directory, `${id}.xlsx`);
    await writeWorkbook(paths[id], '人员', rows);
  }
  return { directory, files: order.map((id) => ({ id, path: paths[id] })) };
}

test('aggregates typed keys, classifies rows, and preserves file and field order', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [
      ['编号', '部门', '姓名', '工资', '未比较', '状态'],
      ['1', '研发', 'Same', 10, 'x', '在职'],
      ['2', '研发', 'Alice', 100, 'x', '在职'],
      ['3', '研发', 'Deleted', 20, 'x', '在职'],
      ['5', '研发', 'Duplicate', 50, 'x', '在职'],
      ['6', '研发', 'Before', 60, 'x', '在职'],
      [null, '研发', 'Invalid', 0, 'x', '在职'],
      ['skip', '研发', 'Filtered', 0, 'x', '离职']
    ],
    B: [
      ['编号', '部门', 'Name', '工资', '未比较', '状态'],
      ['6', '研发', 'After', 61, 'x', '在职'],
      ['5', '研发', 'Duplicate', 50, 'x', '在职'],
      ['2', '研发', 'Alice', 101, 'x', '在职'],
      ['4', '研发', 'Added', 40, 'x', '在职'],
      ['5', '研发', 'Duplicate again', 50, 'x', '在职'],
      ['1', '研发', 'Same', 10, 'x', '在职']
    ],
    C: [
      ['编号', '部门', 'Ｎａｍｅ', '工资', '未比较', '状态'],
      ['4', '研发', 'Added', 40, 'x', '在职'],
      ['2', '研发', 'Alice', 100, 'x', '在职'],
      ['5', '研发', 'Duplicate', 50, 'x', '在职'],
      ['6', '研发', 'Before', 60, 'x', '在职'],
      ['1', '研发', 'Same', 10, 'x', '在职']
    ]
  });

  const result = await compare(spec(directory, files));

  assert.deepEqual(result.summary, {
    files: 3,
    totalRowsScanned: 18,
    matchedRows: 16,
    identicalKeys: 1,
    changedKeys: 2,
    missingKeys: 2,
    duplicateKeys: 1,
    invalidRows: 1
  });
  assert.deepEqual(result.duplicates, [{ key: [['string', '5']], files: ['B'] }]);
  assert.deepEqual(result.missing, [
    { key: [['string', '3']], sheetName: '人员', presentFiles: ['A'], missingFiles: ['C', 'B'], baselineRelation: 'DELETED' },
    { key: [['string', '4']], sheetName: '人员', presentFiles: ['C', 'B'], missingFiles: ['A'], baselineRelation: 'ADDED' }
  ]);
  assert.deepEqual(result.changed.map(({ key, column }) => [key, column]), [
    [[['string', '2']], '工资'],
    [[['string', '6']], '姓名'],
    [[['string', '6']], '工资']
  ]);
  assert.deepEqual(result.changed[0], {
    key: [['string', '2']],
    sheetName: '人员',
    column: '工资',
    files: {
      C: { value: ['number', 100], rowNumber: 3 },
      A: { value: ['number', 100], rowNumber: 3 },
      B: { value: ['number', 101], rowNumber: 4 }
    }
  });
  assert.deepEqual(Object.keys(result.changed[1].files), ['C', 'A', 'B']);
});

test('streams details to awaited sinks without retaining them in the result', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], ['1', 'before'], ['2', 'deleted'], ['4', 'duplicate']],
    B: [['编号', '姓名'], ['1', 'after'], ['3', 'added'], ['4', 'duplicate'], ['4', 'again']]
  }, ['A', 'B']);
  const details = { changed: [], missing: [], duplicates: [] };

  const result = await comparePartitioned(spec(directory, files, {
    filters: [],
    columnAliases: {}
  }), {
    onChanged: async (item) => details.changed.push(item),
    onMissing: async (item) => details.missing.push(item),
    onDuplicate: async (item) => details.duplicates.push(item)
  });

  assert.deepEqual(Object.keys(result), ['summary']);
  assert.deepEqual(result.summary, {
    files: 2,
    totalRowsScanned: 7,
    matchedRows: 7,
    identicalKeys: 0,
    changedKeys: 1,
    missingKeys: 2,
    duplicateKeys: 1,
    invalidRows: 0
  });
  assert.deepEqual(details.changed.map(({ key, column }) => [key, column]), [[[['string', '1']], '姓名']]);
  assert.deepEqual(details.missing.map(({ key }) => key), [[['string', '2']], [['string', '3']]]);
  assert.deepEqual(details.duplicates, [{ key: [['string', '4']], files: ['B'] }]);
});

test('enforces all five resource limits without leaking cell values', async (t) => {
  const secret = 'RESOURCE-SECRET-001';
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], [secret, 'before'], ['2', 'same']],
    B: [['编号', '姓名'], [secret, 'after'], ['2', 'same']]
  }, ['A', 'B']);
  const rules = spec(directory, files, { filters: [], columnAliases: {} });
  const cases = [
    ['ROW_LIMIT_EXCEEDED', { maxRows: 1 }],
    ['CELL_LIMIT_EXCEEDED', { maxCells: 1 }],
    ['INPUT_LIMIT_EXCEEDED', { maxInputBytes: 1 }],
    ['TEMP_LIMIT_EXCEEDED', { maxTempBytes: 1 }]
  ];
  for (const [code, resources] of cases) {
    await assert.rejects(
      () => comparePartitioned({ ...rules, resources }),
      (error) => error.code === code && !error.message.includes(secret)
    );
  }
  const ticks = [0, 2];
  await assert.rejects(
    () => comparePartitioned({ ...rules, resources: { maxRuntimeMs: 1 } }, {}, { now: () => ticks.shift() ?? 2 }),
    (error) => error.code === 'RUNTIME_LIMIT_EXCEEDED' && !error.message.includes(secret)
  );
});

test('keeps an accessible temp directory on success and failure only when requested', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号'], ['1']],
    B: [['编号'], ['1']]
  }, ['A', 'B']);
  const rules = spec(directory, files, { filters: [], columnAliases: {} });

  const result = await comparePartitioned(rules, {}, { keepTemp: true });
  await access(result.tempDirectory);
  const [partitionName] = await readdir(result.tempDirectory);
  const rawRecords = (await readFile(join(result.tempDirectory, partitionName), 'utf8')).trim().split('\n');
  assert.equal(rawRecords.every((line) => Array.isArray(JSON.parse(line)) && JSON.parse(line).length === 8), true);
  assert.equal(/编号|rowBytes|keyHash|fileId/.test(rawRecords.join('\n')), false);
  await rm(result.tempDirectory, { recursive: true, force: true });

  let failure;
  try {
    await comparePartitioned({ ...rules, resources: { maxRows: 1 } }, {}, { keepTemp: true });
  } catch (error) {
    failure = error;
  }
  assert.equal(failure.code, 'ROW_LIMIT_EXCEEDED');
  await access(failure.tempDirectory);
  await rm(failure.tempDirectory, { recursive: true, force: true });
});

test('enforces maxRuntimeMs after awaited sink processing', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号'], ['1']],
    B: [['编号']]
  }, ['A', 'B']);
  let expired = false;

  await assert.rejects(
    () => comparePartitioned(
      spec(directory, files, {
        filters: [],
        columnAliases: {},
        resources: { maxRuntimeMs: 1 }
      }),
      { onMissing: async () => { expired = true; } },
      { now: () => expired ? 2 : 0 }
    ),
    (error) => error.code === 'RUNTIME_LIMIT_EXCEEDED'
  );
});

test('enforces maxRuntimeMs after the final identical fast-path comparison', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], ['1', 'same']],
    B: [['编号', '姓名'], ['1', 'same']]
  }, ['A', 'B']);
  let clockReads = 0;

  await assert.rejects(
    () => comparePartitioned(
      spec(directory, files, {
        filters: [],
        columnAliases: {},
        resources: { maxRuntimeMs: 1 }
      }),
      {},
      { now: () => ++clockReads >= 13 ? 2 : 0 }
    ),
    (error) => error.code === 'RUNTIME_LIMIT_EXCEEDED'
  );
});

test('enforces maxRuntimeMs before a successful empty return', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号']],
    B: [['编号']]
  }, ['A', 'B']);
  let clockReads = 0;

  await assert.rejects(
    () => comparePartitioned(
      spec(directory, files, {
        filters: [],
        columnAliases: {},
        resources: { maxRuntimeMs: 1 }
      }),
      {},
      { now: () => ++clockReads >= 4 ? 2 : 0 }
    ),
    (error) => error.code === 'RUNTIME_LIMIT_EXCEEDED'
  );
});

test('propagates sink failures unchanged', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], ['1', 'before']],
    B: [['编号', '姓名'], ['1', 'after']]
  }, ['A', 'B']);
  const sentinel = new Error('sink stopped');

  await assert.rejects(
    () => comparePartitioned(spec(directory, files, { filters: [], columnAliases: {} }), {
      onChanged: async () => { throw sentinel; }
    }),
    (error) => error === sentinel
  );
});

test('cleans temporary partitions after success and unchanged sink failures', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], ['1', 'before']],
    B: [['编号', '姓名'], ['1', 'after']]
  }, ['A', 'B']);
  const tempRoot = join(directory, 'partitions');
  await mkdir(tempRoot);
  const previousTmp = process.env.TMPDIR;
  process.env.TMPDIR = tempRoot;
  try {
    const rules = spec(directory, files, { filters: [], columnAliases: {} });
    await comparePartitioned(rules);
    assert.deepEqual(await readdir(tempRoot), []);

    const sentinel = new Error('sink cleanup sentinel');
    await assert.rejects(
      () => comparePartitioned(rules, { onChanged: async () => { throw sentinel; } }),
      (error) => error === sentinel
    );
    assert.deepEqual(await readdir(tempRoot), []);
  } finally {
    if (previousTmp === undefined) delete process.env.TMPDIR;
    else process.env.TMPDIR = previousTmp;
  }
});

test('awaits non-sensitive progress in baseline-first scan order', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号'], ['1']],
    B: [['编号'], ['1']]
  }, ['B', 'A']);
  const progress = [];

  await comparePartitioned(
    spec(directory, files, { filters: [], columnAliases: {} }),
    {},
    { onProgress: async (item) => { await Promise.resolve(); progress.push(item); } }
  );

  assert.deepEqual(progress.map(({ rowsScanned, currentFile }) => [rowsScanned, currentFile]), [[1, 'A'], [2, 'B']]);
  assert.equal(progress.every((item) => Object.keys(item).join(',') === 'rowsScanned,currentFile,bytesWritten'), true);
});

test('repartitions colliding buckets while preserving duplicate report and fail policies', async (t) => {
  const buckets = new Map();
  let keys;
  for (let index = 0; keys === undefined; index += 1) {
    const key = String(index);
    const hash = sha256(encodeKey([['string', key]]));
    const previous = buckets.get(hash.slice(0, 2));
    if (previous && previous.hash.slice(2, 4) !== hash.slice(2, 4)) keys = [previous.key, key];
    else buckets.set(hash.slice(0, 2), { hash, key });
  }
  const [duplicateKey, identicalKey] = keys;
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], [duplicateKey, 'duplicate'], [identicalKey, 'same']],
    B: [['编号', '姓名'], [duplicateKey, 'duplicate'], [duplicateKey, 'again'], [identicalKey, 'same']]
  }, ['A', 'B']);
  const rules = spec(directory, files, {
    filters: [],
    columnAliases: {},
    resources: { maxPartitionBytes: 1_200, maxTempBytes: 1_000_000 }
  });

  const reported = await compare(rules);
  assert.equal(reported.summary.identicalKeys, 1);
  assert.equal(reported.summary.duplicateKeys, 1);
  assert.deepEqual(reported.duplicates, [{ key: [['string', duplicateKey]], files: ['B'] }]);

  await assert.rejects(
    () => compare({ ...rules, duplicateKeyPolicy: 'fail' }),
    (error) => error instanceof CompareError && error.code === 'DUPLICATE_KEY'
  );
});

test('emits recursively repartitioned sink records in stable logical hash order', async (t) => {
  const byFirst = new Map();
  let selected;
  for (let index = 0; selected === undefined; index += 1) {
    const key = String(index);
    const hash = sha256(encodeKey([['string', key]]));
    const first = hash.slice(0, 2);
    const second = hash.slice(2, 4);
    const group = byFirst.get(first) ?? new Map();
    const peers = group.get(second) ?? [];
    peers.push({ key, hash });
    group.set(second, peers);
    byFirst.set(first, group);
    const nested = [...group.values()].find((items) => items.length >= 2 && items[0].hash.slice(4, 6) !== items[1].hash.slice(4, 6));
    const sibling = [...group.entries()].find(([bucket]) => bucket !== second)?.[1]?.[0];
    if (nested && sibling) selected = [nested[0], nested[1], sibling];
  }
  const expected = [...selected].sort((left, right) => left.hash.localeCompare(right.hash)).map(({ key }) => key);
  const { directory, files } = await filesFor(t, {
    A: [['编号', '姓名'], ...selected.map(({ key }) => [key, 'x'.repeat(80)])],
    B: [['编号', '姓名']]
  }, ['A', 'B']);
  const rules = spec(directory, files, {
    filters: [],
    columnAliases: {},
    resources: { maxPartitionBytes: 550, maxTempBytes: 1_000_000 }
  });

  const runs = [];
  for (let run = 0; run < 5; run += 1) {
    const keys = [];
    await comparePartitioned(rules, { onMissing: async ({ key }) => keys.push(key[0][1]) });
    runs.push(keys);
  }

  assert.deepEqual(runs, Array.from({ length: 5 }, () => expected));
});

test('preserves canonical numeric-like header order in star comparisons', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['id', '10', '2'], ['001', 'before-ten', 'before-two']],
    B: [['id', '10', '2'], ['001', 'after-ten', 'after-two']]
  }, ['A', 'B']);

  const result = await compare(spec(directory, files, {
    mode: { type: 'key', keyColumns: ['id'] },
    filters: [],
    columnAliases: {}
  }));

  assert.deepEqual(result.changed.map(({ column }) => column), ['10', '2']);
});

test('redacts sensitive duplicate keys and classifies duplicates before missing rows', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '备注'], ['SECRET-EMP-001', 'TOP-SECRET-ROW']],
    B: [['编号', '备注'], ['SECRET-EMP-001', 'TOP-SECRET-ROW'], ['SECRET-EMP-001', 'TOP-SECRET-ROW']],
    C: [['编号', '备注']]
  });
  const rules = spec(directory, files, { filters: [] });
  const reported = await compare(rules);

  assert.equal(reported.summary.duplicateKeys, 1);
  assert.equal(reported.summary.missingKeys, 0);
  assert.deepEqual(reported.duplicates, [{ key: [['string', 'SECRET-EMP-001']], files: ['B'] }]);

  await assert.rejects(
    () => compare({ ...rules, duplicateKeyPolicy: 'fail' }),
    (error) => error instanceof CompareError
      && error.code === 'DUPLICATE_KEY'
      && error.message === 'duplicate business key in files B'
      && !/SECRET-EMP-001|TOP-SECRET-ROW|rowNumber|keyHash|16afa22382763eb45d4fec1255adf423335f2c27eda4ae209aec53237b1ccbf1/.test(error.message)
  );
});

test('honors selected columns, star columns, numeric tolerance, and typed composite keys', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const pathA = join(directory, 'A.xlsx');
  const pathB = join(directory, 'B.xlsx');
  await writeWorkbook(pathA, '人员', [
    ['编号', '分组', '姓名', '金额', '忽略'],
    ['a|b', 'c', 'Same', 10, 'old'],
    ['a', 'b|c', 'Same', 10, 'old'],
    ['3', 'x', 'Same', 10, 'old'],
    ['4', 'x', 'Same', 10, 'old']
  ]);
  await writeWorkbook(pathB, '人员', [
    ['编号', '分组', '姓名', '金额', '忽略'],
    ['a', 'b|c', 'Same', 10, 'new'],
    ['a|b', 'c', 'Same', 10, 'new'],
    ['3', 'x', 'Same', 10.05, 'new'],
    ['4', 'x', 'Same', '10', 'new']
  ]);
  const files = [{ id: 'B', path: pathB }, { id: 'A', path: pathA }];
  const base = spec(directory, files, {
    files,
    mode: { keyColumns: ['编号', '分组'] },
    filters: [],
    columnAliases: {},
    normalization: {
      emptyEqualsNull: false,
      caseSensitive: true,
      formulaMode: 'formula-and-cached-result',
      columns: { 金额: { numericTolerance: 0.1 } }
    }
  });

  const selected = await compare({ ...base, compareColumns: ['姓名', '金额'] });
  assert.equal(selected.summary.identicalKeys, 3);
  assert.equal(selected.summary.changedKeys, 1);
  assert.deepEqual(selected.changed.map(({ key, column }) => [key, column]), [
    [[['string', '4'], ['string', 'x']], '金额']
  ]);

  const all = await compare({ ...base, compareColumns: '*' });
  assert.equal(all.summary.changedKeys, 4);
  assert.deepEqual(all.changed.map(({ column }) => column), ['忽略', '金额', '忽略', '忽略', '忽略']);
});

test('treats shared formula followers as their equivalent expanded formulas', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const sharedPath = join(directory, 'shared.xlsx');
  const plainPath = join(directory, 'plain.xlsx');

  const shared = new ExcelJS.Workbook();
  const sharedSheet = shared.addWorksheet('人员');
  sharedSheet.addRow(['编号', '基础', '计算']);
  sharedSheet.getCell('A2').value = '001';
  sharedSheet.getCell('B2').value = { formula: 'A2+1', result: 2 };
  sharedSheet.getCell('C2').value = { sharedFormula: 'B2', result: 3 };
  await shared.xlsx.writeFile(sharedPath);

  const plain = new ExcelJS.Workbook();
  const plainSheet = plain.addWorksheet('人员');
  plainSheet.addRow(['编号', '基础', '计算']);
  plainSheet.getCell('A2').value = '001';
  plainSheet.getCell('B2').value = { formula: 'A2+1', result: 2 };
  plainSheet.getCell('C2').value = { formula: 'B2+1', result: 3 };
  await plain.xlsx.writeFile(plainPath);

  const result = await compare(spec(directory, [
    { id: 'A', path: sharedPath },
    { id: 'B', path: plainPath }
  ], { filters: [], columnAliases: {} }));

  assert.equal(result.summary.identicalKeys, 1);
  assert.deepEqual(result.changed, []);
});

test('reports changed hyperlink text and targets from real XLSX cells', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const paths = Object.fromEntries(['baseline', 'text', 'target'].map((id) => [id, join(directory, `${id}.xlsx`)]));
  for (const [id, text, hyperlink] of [
    ['baseline', 'Open', 'https://example.test/one'],
    ['text', 'Closed', 'https://example.test/one'],
    ['target', 'Open', 'https://example.test/two']
  ]) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('人员');
    sheet.addRow(['编号', '链接']);
    sheet.getCell('A2').value = '001';
    sheet.getCell('B2').value = { text, hyperlink };
    await workbook.xlsx.writeFile(paths[id]);
  }

  for (const id of ['text', 'target']) {
    const result = await compare(spec(directory, [
      { id: 'A', path: paths.baseline },
      { id: 'B', path: paths[id] }
    ], { filters: [], columnAliases: {} }));
    assert.equal(result.summary.changedKeys, 1);
    assert.deepEqual(result.changed.map(({ column }) => column), ['链接']);
  }
});
