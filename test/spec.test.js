import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import test from 'node:test';

import { SpecError, loadSpec } from '../src/spec.js';

function validSpec(overrides = {}) {
  return {
    version: '1.0',
    baseline: 'before',
    files: [
      { id: 'before', path: 'before.xlsx' },
      { id: 'after', path: 'after.xlsx' }
    ],
    sheet: { name: 'Sheet1', headerRow: 1 },
    mode: { type: 'key', keyColumns: ['员工编号'] },
    compareColumns: '*',
    duplicateKeyPolicy: 'report',
    output: { directory: 'result' },
    ...overrides
  };
}

async function writeSpec(t, spec) {
  const directory = await mkdtemp(resolve(tmpdir(), 'excel-diff-spec-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const path = resolve(directory, 'rules', 'compare.json');
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, JSON.stringify(spec), { encoding: 'utf8', flag: 'w' });
  return path;
}

async function rejectsSpec(t, spec, expected) {
  const path = await writeSpec(t, spec);
  await assert.rejects(
    () => loadSpec(path),
    (error) => error instanceof SpecError && error.code === 'SPEC_INVALID' && expected.test(error.message)
  );
}

test('resolves input and output paths from the rules file directory', async (t) => {
  const path = await writeSpec(t, validSpec({
    files: [
      { id: 'before', path: '../input/before.XLSX' },
      { id: 'after', path: '../input/after.xlsx' }
    ],
    output: { directory: '../output' }
  }));

  const spec = await loadSpec(path);
  const rulesDirectory = dirname(path);
  assert.equal(spec.files[0].path, resolve(rulesDirectory, '../input/before.XLSX'));
  assert.equal(spec.files[1].path, resolve(rulesDirectory, '../input/after.xlsx'));
  assert.equal(spec.output.directory, resolve(rulesDirectory, '../output'));
});

test('rejects unknown root properties', async (t) => {
  await rejectsSpec(t, validSpec({ unexpected: true }), /must NOT have additional properties/);
});

test('rejects duplicate file IDs', async (t) => {
  await rejectsSpec(t, validSpec({
    files: [
      { id: 'same', path: 'before.xlsx' },
      { id: 'same', path: 'after.xlsx' }
    ]
  }), /duplicate file ID/);
});

test('rejects duplicate resolved input paths', async (t) => {
  await rejectsSpec(t, validSpec({
    files: [
      { id: 'before', path: './same.xlsx' },
      { id: 'after', path: 'sub/../same.xlsx' }
    ]
  }), /duplicate input path/);
});

test('rejects a baseline that is not a file ID', async (t) => {
  await rejectsSpec(t, validSpec({ baseline: 'missing' }), /baseline/);
});

test('rejects non-XLSX input files', async (t) => {
  await rejectsSpec(t, validSpec({
    files: [
      { id: 'before', path: 'before.csv' },
      { id: 'after', path: 'after.xlsx' }
    ]
  }), /\.xlsx/);
});

test('rejects an output directory equal to an input file', async (t) => {
  await rejectsSpec(t, validSpec({ output: { directory: 'before.xlsx' } }), /output directory/);
});

test('rejects filters with missing or invalid value counts', async (t) => {
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'eq' }] }), /requires value/);
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'in', values: [] }] }), /at least one value/);
  await rejectsSpec(t, validSpec({ filters: [{ column: '日期', operator: 'between', values: ['2026-01-01'] }] }), /exactly two values/);
});

test('accepts arrays in value for membership and range filters', async (t) => {
  const spec = await loadSpec(await writeSpec(t, validSpec({
    filters: [
      { column: '状态', operator: 'in', value: ['在职'] },
      { column: '状态', operator: 'notIn', value: ['离职'] },
      { column: '日期', operator: 'between', value: ['2026-01-01', '2026-12-31'] }
    ]
  })));

  assert.equal(spec.filters.length, 3);
});

test('rejects values on null filters', async (t) => {
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'isNull', value: null }] }), /must not include value or values/);
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'isNotNull', values: [] }] }), /must not include value or values/);
});

test('rejects conflicting filter value fields', async (t) => {
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'eq', value: '在职', values: ['在职'] }] }), /must not include values/);
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'eq', values: ['在职'] }] }), /must not include values/);
  await rejectsSpec(t, validSpec({ filters: [{ column: '状态', operator: 'in', value: ['在职'], values: ['离职'] }] }), /exactly one of value or values/);
});

test('rejects unknown properties in nested rule objects', async (t) => {
  const cases = [
    ['files item', (spec) => { spec.files[0].unexpected = true; }],
    ['sheet', (spec) => { spec.sheet.unexpected = true; }],
    ['mode', (spec) => { spec.mode.unexpected = true; }],
    ['filter', (spec) => { spec.filters = [{ column: '状态', operator: 'eq', value: '在职', unexpected: true }]; }],
    ['normalization column rule', (spec) => { spec.normalization = { columns: { 状态: { unexpected: true } } }; }],
    ['output', (spec) => { spec.output.unexpected = true; }]
  ];

  for (const [name, mutate] of cases) {
    const spec = validSpec();
    mutate(spec);
    await rejectsSpec(t, spec, /must NOT have additional properties/);
  }
});

test('defaults sampleSize and optional collections', async (t) => {
  const spec = await loadSpec(await writeSpec(t, validSpec()));

  assert.equal(spec.output.sampleSize, 20);
  assert.deepEqual(spec.columnAliases, {});
  assert.deepEqual(spec.filters, []);
  assert.deepEqual(spec.normalization, {
    emptyEqualsNull: false,
    caseSensitive: true,
    formulaMode: 'formula-and-cached-result',
    columns: {}
  });
});
