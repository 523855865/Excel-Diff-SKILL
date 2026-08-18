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
  const originalOrder = result.changed.map((item) => [structuredClone(item.key), item.column]);

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
    ['key', 'sheet', 'column', 'B.value', 'B.row', 'A.value', 'A.row', 'C.value', 'C.row'],
    [JSON.stringify([typed('string', 'a')]), 'Data', 'alpha', JSON.stringify(typed('number', 2)), '2', JSON.stringify(typed('number', 1)), '2', JSON.stringify(typed('number', 1)), '2'],
    [JSON.stringify([typed('string', 'a')]), 'Data', 'zeta', JSON.stringify(typed('number', 2)), '2', JSON.stringify(typed('number', 1)), '2', JSON.stringify(typed('number', 1)), '2'],
    [JSON.stringify([typed('string', 'z')]), 'Data', 'beta', JSON.stringify(typed('string', 'a,"b\nc')), '3', JSON.stringify(typed('formula', ['=SUM(A1:A2)', typed('number', 2)])), '4', JSON.stringify(typed('blank', null)), '5']
  ]);
  assert.equal(missingCsv, 'key,sheet,presentFiles,missingFiles,baselineRelation\n"[[""string"",""has|pipe""]]",Data,"[""B|east"",""A""]","[""C|west""]",ADDED\n');
  assert.equal(summary.status, 'COMPLETED');
  assert.equal(summary.runId, report.summary.runId);
  assert.deepEqual(summary.artifacts, { changed: 'changed.csv', missing: 'missing.csv' });
  assert.equal(summary.files, 3);
  assert.equal(summaryText.endsWith('\n'), true);
  assert.deepEqual(result.changed.map((item) => [item.key, item.column]), originalOrder);
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

  assert.deepEqual(row.filter((_, index) => index >= 3 && index % 2 === 1), values.map(JSON.stringify));

  const formula = typed('string', '=HYPERLINK("https://example.test","open")');
  const formulaReport = await writeReport(reportSpec(directory), reportResult({
    changed: [{
      key: [typed('string', 'formula')], sheetName: 'Data', column: 'value',
      files: { A: { value: formula, rowNumber: 2 }, B: { value: formula, rowNumber: 2 }, C: { value: formula, rowNumber: 2 } }
    }]
  }));
  const formulaCell = parseCsv(await readFile(join(formulaReport.directory, 'changed.csv'), 'utf8'))[1][3];
  assert.equal(formulaCell, JSON.stringify(formula));
  assert.equal(/^[=+\-@]/.test(formulaCell), false);
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

  assert.deepEqual(header.slice(3), files.flatMap((id) => [`json:${JSON.stringify(`${id}.value`)}`, `json:${JSON.stringify(`${id}.row`)}`]));
  assert.equal(row[1], `json:${JSON.stringify('=sheet')}`);
  assert.equal(row[2], `json:${JSON.stringify('+column')}`);
  assert.equal(header.concat(row).some((value) => /^[=+\-@]/.test(value)), false);
});

test('writeReport creates unique runs and writes header-only CSVs for empty details', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = reportSpec(directory);

  const first = await writeReport(spec, reportResult());
  const second = await writeReport(spec, reportResult());

  assert.notEqual(first.summary.runId, second.summary.runId);
  assert.match(second.summary.runId, /^\d{8}T\d{6}Z-[0-9a-f]{8}$/);
  assert.equal(await readFile(join(first.directory, 'changed.csv'), 'utf8'), 'key,sheet,column,B.value,B.row,A.value,A.row,C.value,C.row\n');
  assert.equal(await readFile(join(first.directory, 'missing.csv'), 'utf8'), 'key,sheet,presentFiles,missingFiles,baselineRelation\n');
});
