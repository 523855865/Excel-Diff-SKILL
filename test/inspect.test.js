import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { copyFile, realpath, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import test from 'node:test';

import { makeTempDir } from './helpers.js';
import { inspectFiles } from '../src/inspect.js';

async function workbookFixture(t, name, build) {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const path = join(directory, name);
  const workbook = new ExcelJS.Workbook();
  await build(workbook);
  await workbook.xlsx.writeFile(path);
  return path;
}

function risk(file, code) {
  return file.risks.find((item) => item.code === code);
}

async function filePair(path) {
  const copy = `${path}.copy.xlsx`;
  await copyFile(path, copy);
  return [path, copy];
}

test('inspects sheet metadata, headers, sampled columns, keys, formulas, and merges without exposing values', async (t) => {
  const path = await workbookFixture(t, 'inspect.xlsx', async (workbook) => {
    const target = workbook.addWorksheet('人员');
    target.addRow(['ＩＤ', '姓名', '分数', '可空', '计算']);
    target.addRow(['secret-001', 'secret-name', 10, 'present', { formula: '1+1', result: 2 }]);
    target.addRow(['secret-002', 'secret-name', 10, null, { formula: '1+1', result: 2 }]);
    target.mergeCells('D2:D3');
    workbook.addWorksheet('归档', { state: 'hidden' });
    workbook.addWorksheet('系统', { state: 'veryHidden' });
  });

  const result = await inspectFiles(await filePair(path), { sheet: '人员' });
  const file = result.files[0];

  assert.deepEqual({ status: result.status, sampleRows: result.sampleRows, fullTypes: result.fullTypes }, {
    status: 'INSPECTED',
    sampleRows: 10_000,
    fullTypes: false
  });
  assert.equal(file.file, await realpath(path));
  assert.deepEqual(file.sheets, [
    { name: '人员', visibility: 'visible' },
    { name: '归档', visibility: 'hidden' },
    { name: '系统', visibility: 'veryHidden' }
  ]);
  assert.deepEqual(file.sheet, { name: '人员', visibility: 'visible', headerRow: 1 });
  assert.equal(file.sampledRows, 2);
  assert.deepEqual(file.headers, [
    { index: 1, raw: 'ＩＤ', normalized: 'ID', types: { string: 2 }, emptyRatio: 0 },
    { index: 2, raw: '姓名', normalized: '姓名', types: { string: 2 }, emptyRatio: 0 },
    { index: 3, raw: '分数', normalized: '分数', types: { number: 2 }, emptyRatio: 0 },
    { index: 4, raw: '可空', normalized: '可空', types: { string: 1, blank: 1 }, emptyRatio: 0.5 },
    { index: 5, raw: '计算', normalized: '计算', types: { formula: 2 }, emptyRatio: 0 }
  ]);
  assert.deepEqual(file.keyCandidates, [{ columns: ['ID'], duplicateRate: 0, sampledRows: 2 }]);
  assert.deepEqual(risk(file, 'FORMULA_CELLS'), { code: 'FORMULA_CELLS', count: 2 });
  assert.deepEqual(risk(file, 'MERGED_CELLS'), { code: 'MERGED_CELLS', count: 1 });
  assert.doesNotMatch(JSON.stringify(result), /secret-001|secret-002|secret-name|present/);
});

test('bounds sampled types and candidate memory while fullTypes scans every data row', async (t) => {
  const path = await workbookFixture(t, 'sampling.xlsx', async (workbook) => {
    const sheet = workbook.addWorksheet('数据');
    sheet.addRows([
      ['编号', '混合'],
      ['A', 1],
      ['B', 'two'],
      ['A', true],
      ['C', null]
    ]);
  });

  const paths = await filePair(path);
  const sampled = await inspectFiles(paths, { sheet: '数据', sampleRows: 2 });
  const full = await inspectFiles(paths, { sheet: '数据', sampleRows: 2, fullTypes: true });

  assert.deepEqual(sampled.files[0].headers[1], {
    index: 2,
    raw: '混合',
    normalized: '混合',
    types: { number: 1, string: 1 },
    emptyRatio: 0
  });
  assert.deepEqual(full.files[0].headers[1], {
    index: 2,
    raw: '混合',
    normalized: '混合',
    types: { number: 1, string: 1, boolean: 1, blank: 1 },
    emptyRatio: 0.25
  });
  for (const result of [sampled, full]) {
    assert.equal(result.files[0].sampledRows, 2);
    assert.deepEqual(result.files[0].keyCandidates.find(({ columns }) => columns[0] === '编号'), {
      columns: ['编号'],
      duplicateRate: 0,
      sampledRows: 2
    });
  }
});

test('records NFKC duplicate headers as risks and excludes them from key candidates', async (t) => {
  const path = await workbookFixture(t, 'duplicates.xlsx', async (workbook) => {
    workbook.addWorksheet('数据').addRows([
      ['Ａ', 'A', '唯一'],
      ['left-1', 'right-1', 'key-1'],
      ['left-2', 'right-2', 'key-2']
    ]);
  });

  const { files: [file] } = await inspectFiles(await filePair(path), { sheet: '数据' });

  assert.deepEqual(file.headers.map(({ raw, normalized }) => ({ raw, normalized })), [
    { raw: 'Ａ', normalized: 'A' },
    { raw: 'A', normalized: 'A' },
    { raw: '唯一', normalized: '唯一' }
  ]);
  assert.deepEqual(risk(file, 'HEADER_DUPLICATED'), {
    code: 'HEADER_DUPLICATED',
    normalized: 'A',
    columns: [
      { index: 1, raw: 'Ａ' },
      { index: 2, raw: 'A' }
    ]
  });
  assert.deepEqual(file.keyCandidates, [{ columns: ['唯一'], duplicateRate: 0, sampledRows: 2 }]);
});

test('requires at least two XLSX inputs', async (t) => {
  const path = await workbookFixture(t, 'single.xlsx', async (workbook) => {
    workbook.addWorksheet('数据').addRow(['编号']);
  });

  for (const paths of [[path], []]) {
    await assert.rejects(
      () => inspectFiles(paths, { sheet: '数据' }),
      (error) => error?.code === 'INPUT_ERROR' && error.message === 'at least two XLSX inputs are required'
    );
  }
});

test('resolves real XLSX inputs and rejects duplicate, unreadable, non-XLSX, and missing-sheet inputs', async (t) => {
  const path = await workbookFixture(t, 'input.xlsx', async (workbook) => {
    workbook.addWorksheet('实际表').addRow(['编号']);
  });
  const paths = await filePair(path);
  const text = join(await realpath(join(path, '..')), 'input.txt');
  await writeFile(text, 'not xlsx');

  await assert.rejects(
    () => inspectFiles([path, path], { sheet: '实际表' }),
    (error) => error?.code === 'INPUT_ERROR' && /duplicate input path/.test(error.message)
  );
  await assert.rejects(
    () => inspectFiles([text, paths[0]], { sheet: '实际表' }),
    (error) => error?.code === 'INPUT_ERROR' && /\.xlsx/.test(error.message)
  );
  await assert.rejects(
    () => inspectFiles([join(path, '..', 'missing.xlsx'), paths[0]], { sheet: '实际表' }),
    (error) => error?.code === 'INPUT_ERROR'
  );
  await assert.rejects(
    () => inspectFiles(paths, { sheet: '不存在' }),
    (error) => error?.code === 'SHEET_NOT_FOUND' && error.message === 'sheet 不存在 not found; available: 实际表'
  );
});
