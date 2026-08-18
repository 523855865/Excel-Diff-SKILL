import assert from 'node:assert/strict';
import { rm } from 'node:fs/promises';
import { join } from 'node:path';
import test from 'node:test';

import { compare, CompareError } from '../src/compare.js';
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

test('fails without exposing rows when duplicate-key policy is fail', async (t) => {
  const { directory, files } = await filesFor(t, {
    A: [['编号', '状态'], ['1', '在职']],
    B: [['编号', '状态'], ['1', '在职'], ['1', '在职']],
    C: [['编号', '状态'], ['1', '在职']]
  });

  await assert.rejects(
    () => compare(spec(directory, files, { duplicateKeyPolicy: 'fail' })),
    (error) => error instanceof CompareError
      && error.code === 'DUPLICATE_KEY'
      && /B/.test(error.message)
      && !/rowNumber|在职/.test(error.message)
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
