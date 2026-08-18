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
  assert.equal(changedCsv, [
    'key,sheet,column,B.value,B.row,A.value,A.row,C.value,C.row',
    '"[[""string"",""a""]]",Data,alpha,2,2,1,2,1,2',
    '"[[""string"",""a""]]",Data,zeta,2,2,1,2,1,2',
    '"[[""string"",""z""]]",Data,beta,"a,""b',
    'c",3,"[""=SUM(A1:A2)"",[""number"",2]]",4,,5',
    ''
  ].join('\n'));
  assert.equal(missingCsv, 'key,sheet,presentFiles,missingFiles,baselineRelation\n"[[""string"",""has|pipe""]]",Data,"[""B|east"",""A""]","[""C|west""]",ADDED\n');
  assert.equal(summary.status, 'COMPLETED');
  assert.equal(summary.runId, report.summary.runId);
  assert.deepEqual(summary.artifacts, { changed: 'changed.csv', missing: 'missing.csv' });
  assert.equal(summary.files, 3);
  assert.equal(summaryText.endsWith('\n'), true);
  assert.deepEqual(result.changed, changed);
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
