import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import test from 'node:test';
import unzipper from 'unzipper';

import { makeTempDir, writeWorkbook } from './helpers.js';
import { InputError, readFileRows, scanFileRows } from '../src/read-xlsx.js';

const require = createRequire(import.meta.url);
const tmp = require('tmp');

function spec(overrides = {}) {
  return {
    sheet: { name: '人员', headerRow: 1 },
    mode: { type: 'key', keyColumns: ['编号'] },
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

test('loads the ExcelJS 4.4 streaming compatibility boundary', async () => {
  const compatibility = await import('../src/exceljs-stream-compat.js');
  assert.equal(typeof compatibility.openStreamingWorkbook, 'function');
});

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

test('accepts blank values in row and multiset modes', async (t) => {
  const file = await fixture(t, [['编号'], [null]]);

  for (const type of ['row', 'multiset']) {
    const result = await readFileRows(file, spec({ mode: { type } }));
    assert.equal(result.invalidRows, 0);
    assert.equal(result.rows.length, 1);
    assert.deepEqual(result.rows[0].values.编号, ['blank', null]);
  }
});

test('streams 2,000 real XLSX rows through awaited callbacks without returning rows', async (t) => {
  const file = await fixture(t, [
    ['编号', '姓名'],
    ...Array.from({ length: 2_000 }, (_, index) => [`${index + 1}`.padStart(4, '0'), `员工${index + 1}`])
  ]);
  const records = [];

  const result = await scanFileRows(file, spec(), null, async (record) => {
    await Promise.resolve();
    if (record.rowNumber === 2 || record.rowNumber === 2_001) records.push(record);
  });

  assert.deepEqual(result, {
    columns: ['编号', '姓名'],
    invalidRows: 0,
    totalRowsScanned: 2_000,
    matchedRows: 2_000
  });
  assert.equal(Object.hasOwn(result, 'rows'), false);
  assert.deepEqual(records.map(({ rowNumber }) => rowNumber), [2, 2_001]);
  assert.deepEqual(records[0].values.编号, ['string', '0001']);
  assert.deepEqual(records[1].values.姓名, ['string', '员工2000']);
});

test('propagates callback errors exactly and stops before later rows', async (t) => {
  const file = await fixture(t, [
    ['编号'],
    ['001'],
    ['002'],
    ['003']
  ]);
  const sentinel = new Error('stop scanning');
  const seen = [];

  await assert.rejects(
    () => scanFileRows(file, spec(), null, async (record) => {
      seen.push(record.rowNumber);
      if (record.rowNumber === 3) throw sentinel;
    }),
    (error) => error === sentinel
  );
  assert.deepEqual(seen, [2, 3]);
});

test('releases formula metadata for rows before the header', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'pre-header-formulas.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.getCell('B1').value = { formula: '1=1', result: true };
  sheet.getCell('C1').value = { formula: '1/0', result: { error: '#DIV/0!' } };
  sheet.addRow(['编号', '姓名']);
  sheet.addRow(['001', 'Alice']);
  await workbook.xlsx.writeFile(file.path);

  let released = false;
  const deleteMapEntry = Map.prototype.delete;
  Map.prototype.delete = function deleteWithProbe(key) {
    const pending = this.get(key);
    if (key === 1 && pending instanceof Map && pending.has('B1') && pending.has('C1')) released = true;
    return deleteMapEntry.call(this, key);
  };
  try {
    await scanFileRows(file, spec({ sheet: { name: '人员', headerRow: 2 } }));
  } finally {
    Map.prototype.delete = deleteMapEntry;
  }
  assert.equal(released, true);
});

test('streams ExcelJS workbooks whose worksheet entry precedes workbook metadata', async (t) => {
  const file = await fixture(t, [
    ['编号', '日期'],
    ['001', new Date('2026-01-02T00:00:00.000Z')]
  ]);

  const result = await readFileRows(file, spec());

  assert.deepEqual(result.rows[0].values.日期, ['date', '2026-01-02T00:00:00.000Z']);
});

test('streams inline strings without temporary worksheet spooling', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'inline-strings.xlsx') };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file.path, useSharedStrings: false });
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '姓名']).commit();
  sheet.addRow(['001', 'Alice']).commit();
  sheet.commit();
  await workbook.commit();

  const archive = await unzipper.Open.file(file.path);
  const paths = archive.files.map(({ path }) => path);
  assert.equal(paths.includes('xl/sharedStrings.xml'), false);
  assert.ok(paths.indexOf('xl/worksheets/sheet1.xml') < paths.indexOf('xl/workbook.xml'));

  const originalTmpFile = tmp.file;
  tmp.file = () => { throw new Error('unexpected worksheet temp spool'); };
  try {
    const result = await scanFileRows(file, spec());
    assert.equal(result.matchedRows, 1);
    assert.deepEqual(result.columns, ['编号', '姓名']);
  } finally {
    tmp.file = originalTmpFile;
  }
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

test('expands interleaved shared formula groups by their shared IDs', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'interleaved-formulas.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '加一', '乘十']);
  sheet.addRow(['001']);
  sheet.addRow(['002']);
  sheet.fillFormula('B2:B3', 'A2+1', [2, 3]);
  sheet.fillFormula('C2:C3', 'B2*10', [20, 30]);
  await workbook.xlsx.writeFile(file.path);

  const result = await readFileRows(file, spec());

  assert.deepEqual(result.rows[1].values.加一, ['formula', ['A3+1', ['number', 3]]]);
  assert.deepEqual(result.rows[1].values.乘十, ['formula', ['B3*10', ['number', 30]]]);
});

test('normalizes cached formula dates in every formula mode', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'formula-date.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '日期公式']);
  sheet.getCell('A2').value = '001';
  sheet.getCell('B2').value = { formula: 'DATE(2026,1,2)', result: new Date('2026-01-02T00:00:00.000Z') };
  sheet.getCell('B2').numFmt = 'yyyy-mm-dd';
  await workbook.xlsx.writeFile(file.path);

  for (const [formulaMode, expected] of [
    ['formula', ['formula', 'DATE(2026,1,2)']],
    ['cached-result', ['date', '2026-01-02T00:00:00.000Z']],
    ['formula-and-cached-result', ['formula', ['DATE(2026,1,2)', ['date', '2026-01-02T00:00:00.000Z']]]]
  ]) {
    const rules = spec({ normalization: { columns: {}, formulaMode } });
    const result = await readFileRows(file, rules);
    assert.deepEqual(result.rows[0].values.日期公式, expected);
  }
});

test('normalizes cached formula booleans and errors in every formula mode', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'formula-types.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '布尔公式', '错误公式', '共享布尔公式']);
  sheet.addRow(['001']);
  sheet.addRow(['002']);
  sheet.getCell('B2').value = { formula: '1=1', result: true };
  sheet.getCell('B3').value = { formula: '1=0', result: false };
  sheet.getCell('C2').value = { formula: '1/0', result: { error: '#DIV/0!' } };
  sheet.getCell('C3').value = { formula: 'NA()', result: { error: '#N/A' } };
  sheet.fillFormula('D2:D3', 'A2<>""', [true, false]);
  await workbook.xlsx.writeFile(file.path);

  for (const [formulaMode, expected] of [
    ['formula', {
      布尔公式: ['formula', '1=0'],
      错误公式: ['formula', 'NA()'],
      共享布尔公式: ['formula', 'A3<>""']
    }],
    ['cached-result', {
      布尔公式: ['boolean', false],
      错误公式: ['error', '#N/A'],
      共享布尔公式: ['boolean', false]
    }],
    ['formula-and-cached-result', {
      布尔公式: ['formula', ['1=0', ['boolean', false]]],
      错误公式: ['formula', ['NA()', ['error', '#N/A']]],
      共享布尔公式: ['formula', ['A3<>""', ['boolean', false]]]
    }]
  ]) {
    const rules = spec({ normalization: { columns: {}, formulaMode } });
    const result = await readFileRows(file, rules);
    const values = result.rows[1].values;
    assert.deepEqual({
      布尔公式: values.布尔公式,
      错误公式: values.错误公式,
      共享布尔公式: values.共享布尔公式
    }, expected);
  }
});

test('preserves hyperlink tooltips from real XLSX cells', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = { id: 'before', path: join(directory, 'hyperlink-tooltip.xlsx') };
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('人员');
  sheet.addRow(['编号', '链接']);
  sheet.getCell('A2').value = '001';
  sheet.getCell('B2').value = { text: 'Open', hyperlink: 'https://example.test', tooltip: 'details' };
  await workbook.xlsx.writeFile(file.path);

  const result = await readFileRows(file, spec());

  assert.deepEqual(result.rows[0].values.链接, ['hyperlink', {
    text: ['string', 'Open'],
    target: 'https://example.test',
    tooltip: 'details'
  }]);
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
  const Reader = ExcelJS.stream.xlsx.WorkbookReader;
  const originalIterator = Reader.prototype[Symbol.asyncIterator];
  Reader.prototype[Symbol.asyncIterator] = async function* patchedIterator() {
    for await (const sheet of originalIterator.call(this)) {
      const iterateRows = sheet[Symbol.asyncIterator].bind(sheet);
      sheet[Symbol.asyncIterator] = async function* patchedRows() {
        for await (const row of iterateRows()) {
          if (row.number === 2) row.getCell(2).value = { secret: 'do-not-leak' };
          yield row;
        }
      };
      yield sheet;
    }
  };
  t.after(() => { Reader.prototype[Symbol.asyncIterator] = originalIterator; });

  await assert.rejects(
    () => readFileRows(file, spec()),
    (error) => error instanceof InputError
      && error.code === 'CELL_VALUE_UNSUPPORTED'
      && error.message === 'unsupported cell value in before row 2 column 姓名'
      && !/do-not-leak|secret/.test(error.message)
  );
});
