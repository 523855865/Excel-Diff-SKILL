import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import test from 'node:test';

import { makeTempDir, writeWorkbook } from './helpers.js';
import { InputError, readFileRows } from '../src/read-xlsx.js';

function spec(overrides = {}) {
  return {
    sheet: { name: '人员', headerRow: 1 },
    mode: { keyColumns: ['编号'] },
    compareColumns: '*',
    columnAliases: {},
    filters: [],
    normalization: {
      emptyEqualsNull: false,
      caseSensitive: true,
      formulaMode: 'formula-and-cached-result',
      columns: {}
    },
    ...overrides
  };
}

async function fixture(t, rows, sheetName = '人员') {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const path = join(directory, 'input.xlsx');
  await writeWorkbook(path, sheetName, rows);
  return { id: 'before', path };
}

async function rejects(t, file, rules, expected, columns = null) {
  await assert.rejects(
    () => readFileRows(file, rules, columns),
    (error) => error instanceof InputError && error.code === expected.code && expected.message.test(error.message)
  );
}

test('reads typed values with Excel provenance and physical row counts', async (t) => {
  const file = await fixture(t, [
    ['编号', '姓名', '工资'],
    ['001', 'Alice', 12],
    ['002', 'Bob', 18]
  ]);

  const result = await readFileRows(file, spec());

  assert.deepEqual(result.columns, ['编号', '姓名', '工资']);
  assert.equal(result.invalidRows, 0);
  assert.equal(result.totalRowsScanned, 2);
  assert.deepEqual(result.rows, [
    { fileId: 'before', sheetName: '人员', rowNumber: 2, values: { 编号: ['string', '001'], 姓名: ['string', 'Alice'], 工资: ['number', 12] } },
    { fileId: 'before', sheetName: '人员', rowNumber: 3, values: { 编号: ['string', '002'], 姓名: ['string', 'Bob'], 工资: ['number', 18] } }
  ]);
});

test('filters normalized strings, numbers, and dates while retaining scan counts', async (t) => {
  const file = await fixture(t, [
    ['编号', '状态', '金额', '日期'],
    ['001', '  ACTIVE ', 10, new Date('2026-01-02T00:00:00.000Z')],
    ['002', 'inactive', 20, new Date('2026-01-03T00:00:00.000Z')]
  ]);
  const rules = spec({
    filters: [
      { column: '状态', operator: 'eq', value: 'ACTIVE' },
      { column: '金额', operator: 'gte', value: 10 },
      { column: '日期', operator: 'eq', value: '2026-01-02' }
    ],
    normalization: {
      emptyEqualsNull: false,
      caseSensitive: true,
      formulaMode: 'formula-and-cached-result',
      columns: { 状态: { trim: true, caseSensitive: false } }
    }
  });

  const result = await readFileRows(file, rules);

  assert.equal(result.totalRowsScanned, 2);
  assert.equal(result.rows.length, 1);
  assert.deepEqual(result.rows[0].values.状态, ['string', 'active']);
  assert.deepEqual(result.rows[0].values.日期, ['date', '2026-01-02T00:00:00.000Z']);
});

test('wraps invalid date filters without leaking filter or cell values', async (t) => {
  const file = await fixture(t, [
    ['编号', '日期'],
    ['001', new Date('2026-01-02T00:00:00.000Z')]
  ]);
  const rules = spec({ filters: [{ column: '日期', operator: 'eq', value: '2026-02-30T00:00:00Z' }] });

  await assert.rejects(
    () => readFileRows(file, rules),
    (error) => error instanceof InputError
      && error.code === 'FILTER_INVALID'
      && error.message === 'invalid filter for column 日期'
      && !/2026-02-30|2026-01-02|001/.test(error.message)
  );
});

test('counts blank typed keys as invalid and leaves source order unchanged', async (t) => {
  const file = await fixture(t, [
    ['编号', '姓名'],
    ['002', 'second'],
    [null, 'invalid'],
    ['001', 'first']
  ]);

  const result = await readFileRows(file, spec());

  assert.equal(result.invalidRows, 1);
  assert.deepEqual(result.rows.map((row) => row.rowNumber), [2, 4]);
  assert.deepEqual(result.rows.map((row) => row.values.编号), [['string', '002'], ['string', '001']]);
});

test('filters rows before counting blank keys as invalid', async (t) => {
  const file = await fixture(t, [
    ['编号', '状态'],
    [null, '离职'],
    ['001', '在职']
  ]);

  const result = await readFileRows(file, spec({ filters: [{ column: '状态', operator: 'eq', value: '在职' }] }));

  assert.equal(result.invalidRows, 0);
  assert.deepEqual(result.rows.map((row) => row.values.编号), [['string', '001']]);
});

test('reports absent sheets with available sheet names', async (t) => {
  const file = await fixture(t, [['编号'], ['001']], '实际表');

  await rejects(t, file, spec(), { code: 'SHEET_NOT_FOUND', message: /实际表/ });
});

test('rejects exact and NFKC-equivalent duplicate headers with positions', async (t) => {
  const exact = await fixture(t, [['编号', '编号'], ['001', 'x']]);
  await rejects(t, exact, spec(), { code: 'HEADER_DUPLICATED', message: /编号.*1.*2/ });

  const nfkc = await fixture(t, [['Ａ', 'A'], ['001', 'x']]);
  await rejects(t, nfkc, spec({ mode: { keyColumns: ['Ａ'] } }), { code: 'HEADER_DUPLICATED', message: /Ａ.*1.*2/ });
});

test('requires key, explicit compare, and filter columns', async (t) => {
  const file = await fixture(t, [['编号'], ['001']]);
  const rules = spec({
    compareColumns: ['姓名'],
    filters: [{ column: '状态', operator: 'eq', value: '在职' }]
  });

  await rejects(t, file, rules, { code: 'COLUMN_MISSING', message: /before.*姓名|姓名.*before/ });
  await rejects(t, file, spec({ filters: [{ column: '状态', operator: 'eq', value: '在职' }] }), { code: 'COLUMN_MISSING', message: /状态/ });
});

test('requires key columns to exist in the workbook', async (t) => {
  const file = await fixture(t, [['姓名'], ['Alice']]);

  await rejects(
    t,
    file,
    spec({ mode: { keyColumns: ['工号'] } }),
    { code: 'COLUMN_MISSING', message: /before.*工号|工号.*before/ }
  );
});

test('maps later NFKC and alias headers to baseline standard columns', async (t) => {
  const baseline = await fixture(t, [['编号', '姓名', 'ABC'], ['001', 'Alice', 'x']]);
  const rules = spec({ columnAliases: { 姓名: ['Name'] } });
  const first = await readFileRows(baseline, rules);
  const later = await fixture(t, [['ＡＢＣ', 'Name', '编号'], ['x', 'Alice', '001']]);
  const laterRules = spec({
    mode: { keyColumns: ['编号'] },
    compareColumns: '*',
    columnAliases: { 姓名: ['Name'], ABC: ['ＡＢＣ'] }
  });

  const result = await readFileRows(later, laterRules, first.columns);

  assert.deepEqual(first.columns, ['编号', '姓名', 'ABC']);
  assert.deepEqual(result.columns, ['编号', '姓名', 'ABC']);
  assert.deepEqual(result.rows[0].values, {
    编号: ['string', '001'],
    姓名: ['string', 'Alice'],
    ABC: ['string', 'x']
  });
});

test('resolves exact columns before shared aliases', async (t) => {
  const file = await fixture(t, [['编号', 'ID'], ['001', 'X']]);

  const result = await readFileRows(
    file,
    spec({ columnAliases: { 编号: ['ID'], 代码: ['ID'] } }),
    ['编号', '代码']
  );

  assert.deepEqual(result.rows[0].values, { 编号: ['string', '001'], 代码: ['string', 'X'] });
});

test('uses canonical aliases when building baseline star columns', async (t) => {
  const file = await fixture(t, [['ID', '姓名'], ['001', 'Alice']]);

  const result = await readFileRows(file, spec({ columnAliases: { 编号: ['ID'] } }));

  assert.deepEqual(result.columns, ['编号', '姓名']);
  assert.deepEqual(result.rows[0].values, { 编号: ['string', '001'], 姓名: ['string', 'Alice'] });
});

test('canonicalizes baseline alias columns before mapping later star workbooks', async (t) => {
  const rules = spec({ columnAliases: { 姓名: ['Name'] } });
  const baseline = await fixture(t, [['编号', 'Name'], ['001', 'Alice']]);
  const first = await readFileRows(baseline, rules);
  const later = await fixture(t, [['编号', '姓名'], ['001', 'Alice']]);

  const result = await readFileRows(later, rules, first.columns);

  assert.deepEqual(first.columns, ['编号', '姓名']);
  assert.deepEqual(result.rows[0].values, { 编号: ['string', '001'], 姓名: ['string', 'Alice'] });
});

test('reads formula cache values, shared formulas, and rich text from XLSX cells', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'values.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '普通公式', '共享公式', '富文本']);
  sheet.getCell('A2').value = '001';
  sheet.getCell('B2').value = { formula: 'A2+1', result: 2 };
  sheet.getCell('C2').value = { sharedFormula: 'B2', result: 2 };
  sheet.getCell('D2').value = { richText: [{ text: 'Hello' }, { text: ' world' }] };
  await workbook.xlsx.writeFile(file.path);

  const result = await readFileRows(file, spec());

  assert.deepEqual(result.rows[0].values, {
    编号: ['string', '001'],
    普通公式: ['formula', ['A2+1', ['number', 2]]],
    共享公式: ['formula', ['B2+1', ['number', 2]]],
    富文本: ['string', 'Hello world']
  });
});

test('ignores unselected standard columns in explicit comparisons', async (t) => {
  const file = await fixture(t, [['编号', '姓名'], ['001', 'Alice']]);

  const result = await readFileRows(file, spec({ compareColumns: ['姓名'] }), ['编号', '姓名', '工资']);

  assert.deepEqual(result.columns, ['编号', '姓名']);
  assert.deepEqual(result.rows[0].values, { 编号: ['string', '001'], 姓名: ['string', 'Alice'] });
});

test('rejects ambiguous source-to-standard and standard-to-source mappings', async (t) => {
  const preferred = await fixture(t, [['编号', 'ID'], ['001', 'x']]);
  const preferredResult = await readFileRows(preferred, spec({ columnAliases: { 编号: ['ID'] } }), ['编号']);
  assert.deepEqual(preferredResult.rows[0].values.编号, ['string', '001']);

  const sourcePriority = await fixture(t, [['编号'], ['001']]);
  await rejects(
    t,
    sourcePriority,
    spec({ columnAliases: { 代码: ['编号'] } }),
    { code: 'COLUMN_MISSING', message: /代码/ },
    ['编号', '代码']
  );

  const oneStandard = await fixture(t, [['ID', 'Legacy ID'], ['001', 'x']]);
  await rejects(
    t,
    oneStandard,
    spec({ columnAliases: { 编号: ['ID', 'Legacy ID'] } }),
    { code: 'COLUMN_MAPPING_AMBIGUOUS', message: /编号/ },
    ['编号']
  );

  const oneSource = await fixture(t, [['ID'], ['001']]);
  await rejects(
    t,
    oneSource,
    spec({ mode: { keyColumns: ['编号'] }, columnAliases: { 编号: ['ID'], 代码: ['ID'] } }),
    { code: 'COLUMN_MAPPING_AMBIGUOUS', message: /ID/ },
    ['编号', '代码']
  );
});

test('requires every baseline column in star comparisons', async (t) => {
  const file = await fixture(t, [['编号'], ['001']]);

  await rejects(t, file, spec(), { code: 'COLUMN_MISSING', message: /姓名/ }, ['编号', '姓名']);
});

test('wraps unreadable workbooks as INPUT_ERROR', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const broken = { id: 'broken', path: join(directory, 'broken.xlsx') };
  await writeFile(broken.path, 'not an xlsx');

  await rejects(t, broken, spec(), { code: 'INPUT_ERROR', message: /broken/ });
  await rejects(t, { id: 'missing', path: join(directory, 'missing.xlsx') }, spec(), { code: 'INPUT_ERROR', message: /missing/ });
});

test('wraps unsupported cell values without leaking their contents', async (t) => {
  const file = await fixture(t, [['编号', '姓名'], ['001', 'Alice']]);
  const Xlsx = new ExcelJS.Workbook().xlsx.constructor;
  const originalReadFile = Xlsx.prototype.readFile;
  Xlsx.prototype.readFile = async function patchedReadFile(path) {
    await originalReadFile.call(this, path);
    this.workbook.getWorksheet('人员').getCell('B2').value = { secret: 'do-not-leak' };
  };
  t.after(() => { Xlsx.prototype.readFile = originalReadFile; });

  await assert.rejects(
    () => readFileRows(file, spec()),
    (error) => error instanceof InputError
      && error.code === 'CELL_VALUE_UNSUPPORTED'
      && error.message === 'unsupported cell value in before row 2 column 姓名'
      && !/do-not-leak|secret/.test(error.message)
  );
});
