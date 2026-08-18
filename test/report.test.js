import assert from 'node:assert/strict';
import { readFile, rm, stat } from 'node:fs/promises';
import { join } from 'node:path';
import test from 'node:test';

import { makeTempDir } from './helpers.js';
import { csvCell, writeReport } from '../src/report.js';

const typed = (type, value) => [type, value];

function reportSpec(directory, files = ['B', 'A', 'C']) {
  return {
    baseline: 'A',
    files: files.map((id) => ({ id, path: `${id}.xlsx` })),
    output: { directory }
  };
}

function reportResult(overrides = {}) {
  return {
    summary: { files: 3, changedKeys: 2 },
    changed: [],
    missing: [],
    duplicates: [],
    ...overrides
  };
}

function parseCsv(text) {
  const rows = [[]];
  let cell = '';
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted && character === '"' && text[index + 1] === '"') {
      cell += '"';
      index += 1;
    } else if (character === '"') quoted = !quoted;
    else if (!quoted && character === ',') {
      rows.at(-1).push(cell);
      cell = '';
    } else if (!quoted && character === '\n') {
      rows.at(-1).push(cell);
      cell = '';
      if (index + 1 < text.length) rows.push([]);
    } else cell += character;
  }
  return rows;
}

test('csvCell escapes only CSV-special values', () => {
  assert.equal(csvCell(null), '');
  assert.equal(csvCell(undefined), '');
  assert.equal(csvCell('plain'), 'plain');
  assert.equal(csvCell('a,b'), '"a,b"');
  assert.equal(csvCell('a"b'), '"a""b"');
  assert.equal(csvCell('a\rb'), '"a\rb"');
  assert.equal(csvCell('a\nb'), '"a\nb"');
  assert.equal(csvCell('a,"b\r\nc'), '"a,""b\r\nc"');
});

test('writeReport writes deterministic CSV artifacts and a protected summary', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const changed = [
    {
      key: [typed('string', 'z')],
      sheetName: 'Data',
      column: 'beta',
      files: {
        A: { value: typed('formula', ['=SUM(A1:A2)', typed('number', 2)]), rowNumber: 4 },
        B: { value: typed('string', 'a,"b\nc'), rowNumber: 3 },
        C: { value: typed('blank', null), rowNumber: 5 }
      }
    },
    {
      key: [typed('string', 'a')],
      sheetName: 'Data',
      column: 'zeta',
      files: {
        A: { value: typed('number', 1), rowNumber: 2 },
        B: { value: typed('number', 2), rowNumber: 2 },
        C: { value: typed('number', 1), rowNumber: 2 }
      }
    },
    {
      key: [typed('string', 'a')],
      sheetName: 'Data',
      column: 'alpha',
      files: {
        A: { value: typed('number', 1), rowNumber: 2 },
        B: { value: typed('number', 2), rowNumber: 2 },
        C: { value: typed('number', 1), rowNumber: 2 }
      }
    }
  ];
  const result = reportResult({
    summary: { status: 'BAD', runId: 'bad', artifacts: { changed: 'bad.csv' }, files: 3 },
    changed,
    missing: [{
      key: [typed('string', 'has|pipe')],
      sheetName: 'Data',
      presentFiles: ['B|east', 'A'],
      missingFiles: ['C|west'],
      baselineRelation: 'ADDED'
    }]
  });
  const originalResult = structuredClone(result);

  const report = await writeReport(reportSpec(directory), result);
  const [changedCsv, missingCsv, summaryText] = await Promise.all([
    readFile(join(report.directory, 'changed.csv'), 'utf8'),
    readFile(join(report.directory, 'missing.csv'), 'utf8'),
    readFile(join(report.directory, 'summary.json'), 'utf8')
  ]);
  const summary = JSON.parse(summaryText);

  assert.match(report.summary.runId, /^\d{8}T\d{6}Z-[0-9a-f]{8}$/);
  assert.equal(report.directory, join(directory, report.summary.runId));
  assert.equal((await stat(report.directory)).isDirectory(), true);
  assert.deepEqual(parseCsv(changedCsv), [
    ['key', 'sheet', 'column', 'B.value', 'B.type', 'B.row', 'A.value', 'A.type', 'A.row', 'C.value', 'C.type', 'C.row'],
    ['a', 'Data', 'alpha', '2', 'number', '2', '1', 'number', '2', '1', 'number', '2'],
    ['a', 'Data', 'zeta', '2', 'number', '2', '1', 'number', '2', '1', 'number', '2'],
    ['z', 'Data', 'beta', 'a,"b\nc', 'string', '3', JSON.stringify({ formula: '=SUM(A1:A2)', result: 2 }), 'formula', '4', '', 'blank', '5']
  ]);
  assert.equal(missingCsv, 'key,sheet,presentFiles,missingFiles,baselineRelation\nhas|pipe,Data,"[""B|east"",""A""]","[""C|west""]",ADDED\n');
  assert.equal(summary.status, 'COMPLETED');
  assert.equal(summary.runId, report.summary.runId);
  assert.deepEqual(summary.artifacts, { changed: 'changed.csv', missing: 'missing.csv' });
  assert.equal(summary.files, 3);
  assert.equal(summaryText.endsWith('\n'), true);
  assert.deepEqual(result, originalResult);
});

test('writeReport preserves every typed value without spreadsheet formula activation', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const values = [
    typed('blank', null), typed('string', ''), typed('number', 1),
    typed('string', '1'), typed('boolean', true), typed('string', 'true')
  ];
  const files = ['A', 'B', 'C', 'D', 'E', 'F'];
  const report = await writeReport(reportSpec(directory, files), reportResult({
    changed: [{
      key: [typed('string', 'k')], sheetName: 'Data', column: 'value',
      files: Object.fromEntries(files.map((id, index) => [id, { value: values[index], rowNumber: index + 2 }]))
    }]
  }));
  const [, row] = parseCsv(await readFile(join(report.directory, 'changed.csv'), 'utf8'));

  assert.deepEqual(
    values.map((_, index) => [row[3 + index * 3], row[4 + index * 3]]),
    [['', 'blank'], ['', 'string'], ['1', 'number'], ['1', 'string'], ['true', 'boolean'], ['true', 'string']]
  );

  const formula = typed('formula', '=HYPERLINK("https://example.test","open")');
  const formulaReport = await writeReport(reportSpec(directory), reportResult({
    changed: [{
      key: [typed('string', 'formula')], sheetName: 'Data', column: 'value',
      files: { A: { value: formula, rowNumber: 2 }, B: { value: formula, rowNumber: 2 }, C: { value: formula, rowNumber: 2 } }
    }]
  }));
  const formulaCell = parseCsv(await readFile(join(formulaReport.directory, 'changed.csv'), 'utf8'))[1][3];
  assert.equal(formulaCell, `json:${JSON.stringify('=HYPERLINK("https://example.test","open")')}`);
  assert.equal(/^[=+\-@]/.test(formulaCell), false);
});

test('writeReport writes untyped composite keys and hyperlink payloads', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const result = reportResult({
    changed: [{
      key: [typed('string', '001'), typed('number', 2)], sheetName: 'Data', column: 'link',
      files: {
        A: { value: typed('hyperlink', { text: typed('string', 'Open'), target: 'https://a.test', tooltip: null }), rowNumber: 2 },
        B: { value: typed('hyperlink', { text: typed('string', 'Open'), target: 'https://b.test', tooltip: 'new' }), rowNumber: 2 }
      }
    }],
    missing: [{
      key: [typed('string', '001'), typed('number', 2)], sheetName: 'Data',
      presentFiles: ['A'], missingFiles: ['B'], baselineRelation: 'DELETED'
    }]
  });
  const report = await writeReport(reportSpec(directory, ['A', 'B']), result);
  const [changed, missing] = await Promise.all([
    readFile(join(report.directory, 'changed.csv'), 'utf8'),
    readFile(join(report.directory, 'missing.csv'), 'utf8')
  ]);

  assert.deepEqual(parseCsv(changed)[1], [
    JSON.stringify(['001', 2]), 'Data', 'link',
    JSON.stringify({ text: 'Open', target: 'https://a.test', tooltip: null }), 'hyperlink', '2',
    JSON.stringify({ text: 'Open', target: 'https://b.test', tooltip: 'new' }), 'hyperlink', '2'
  ]);
  assert.equal(parseCsv(missing)[1][0], JSON.stringify(['001', 2]));
});

test('writeReport protects direct keys and values from spreadsheet activation', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const report = await writeReport(reportSpec(directory, ['A', 'B']), reportResult({
    changed: [{
      key: [typed('string', '=changed-key')], sheetName: 'Data', column: 'value',
      files: {
        A: { value: typed('string', '\t=changed-value'), rowNumber: 2 },
        B: { value: typed('string', 'safe'), rowNumber: 2 }
      }
    }],
    missing: [{
      key: [typed('string', '\nmissing-key')], sheetName: 'Data',
      presentFiles: ['A'], missingFiles: ['B'], baselineRelation: 'DELETED'
    }]
  }));
  const [changedHeader, changedRow] = parseCsv(await readFile(join(report.directory, 'changed.csv'), 'utf8'));
  const [, missingRow] = parseCsv(await readFile(join(report.directory, 'missing.csv'), 'utf8'));

  assert.equal(changedRow[0], `json:${JSON.stringify('=changed-key')}`);
  assert.equal(changedRow[3], `json:${JSON.stringify('\t=changed-value')}`);
  assert.equal(missingRow[0], `json:${JSON.stringify('\nmissing-key')}`);
  assert.equal(changedHeader.concat(changedRow, missingRow).some((value) => /^[=+\-@\t\r\n]/.test(value)), false);
});

test('writeReport safely and reversibly encodes controlled labels', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const files = ['=file', '+file', '-file', '@file', 'json:file'];
  const report = await writeReport(reportSpec(directory, files), reportResult({
    changed: [{
      key: [typed('string', 'k')], sheetName: '=sheet', column: '+column',
      files: Object.fromEntries(files.map((id) => [id, { value: typed('number', 1), rowNumber: 2 }]))
    }]
  }));
  const [header, row] = parseCsv(await readFile(join(report.directory, 'changed.csv'), 'utf8'));

  assert.deepEqual(header.slice(3), files.flatMap((id) => [
    `json:${JSON.stringify(`${id}.value`)}`,
    `json:${JSON.stringify(`${id}.type`)}`,
    `json:${JSON.stringify(`${id}.row`)}`
  ]));
  assert.equal(row[1], `json:${JSON.stringify('=sheet')}`);
  assert.equal(row[2], `json:${JSON.stringify('+column')}`);
  assert.equal(header.concat(row).some((value) => /^[=+\-@]/.test(value)), false);
});

test('writeReport covers tab and line-break CSV injection prefixes in every label', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const files = ['\n@file', 'A'];
  const report = await writeReport(reportSpec(directory, files), reportResult({
    changed: [{
      key: [typed('string', 'changed')], sheetName: '\t=changed', column: '\r=column',
      files: Object.fromEntries(files.map((id) => [id, { value: typed('number', 1), rowNumber: 2 }]))
    }],
    missing: [{
      key: [typed('string', 'missing')], sheetName: '=missing',
      presentFiles: ['A'], missingFiles: ['\n@file'], baselineRelation: 'DELETED'
    }]
  }));
  const [changedHeader, changedRow] = parseCsv(await readFile(join(report.directory, 'changed.csv'), 'utf8'));
  const [, missingRow] = parseCsv(await readFile(join(report.directory, 'missing.csv'), 'utf8'));
  const decodeSafe = (value) => JSON.parse(value.slice('json:'.length));

  assert.equal(changedHeader.concat(changedRow, missingRow).some((value) => /^[=+\-@\t\r\n]/.test(value)), false);
  assert.equal(decodeSafe(changedHeader[3]), '\n@file.value');
  assert.equal(decodeSafe(changedRow[1]), '\t=changed');
  assert.equal(decodeSafe(changedRow[2]), '\r=column');
  assert.equal(decodeSafe(missingRow[1]), '=missing');
});

test('writeReport creates unique runs and writes header-only CSVs for empty details', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = reportSpec(directory);

  const first = await writeReport(spec, reportResult());
  const second = await writeReport(spec, reportResult());

  assert.notEqual(first.summary.runId, second.summary.runId);
  assert.match(second.summary.runId, /^\d{8}T\d{6}Z-[0-9a-f]{8}$/);
  assert.equal(await readFile(join(first.directory, 'changed.csv'), 'utf8'), 'key,sheet,column,B.value,B.type,B.row,A.value,A.type,A.row,C.value,C.type,C.row\n');
  assert.equal(await readFile(join(first.directory, 'missing.csv'), 'utf8'), 'key,sheet,presentFiles,missingFiles,baselineRelation\n');
});
