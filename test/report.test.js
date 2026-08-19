import assert from 'node:assert/strict';
import { chmodSync, mkdirSync } from 'node:fs';
import { chmod, mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { Writable } from 'node:stream';
import test from 'node:test';

import { makeTempDir } from './helpers.js';
import { createReportWriter, csvCell, writeReport } from '../src/report.js';

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

test('createReportWriter atomically publishes streamed key details', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const writer = await createReportWriter(reportSpec(directory, ['A', 'B']));

  const staged = await readdir(directory);
  assert.equal(staged.length, 1);
  assert.match(staged[0], /^\..+\.tmp$/);
  await writer.onChanged({
    key: [typed('string', '1')], sheetName: '人员', column: '姓名',
    files: {
      A: { value: typed('string', 'before'), rowNumber: 2 },
      B: { value: typed('string', 'after'), rowNumber: 2 }
    }
  });
  await writer.onMissing({
    key: [typed('string', '2')], sheetName: '人员',
    presentFiles: ['A'], missingFiles: ['B'], baselineRelation: 'DELETED'
  });

  const completed = await writer.complete({ files: 2, changedKeys: 1, missingKeys: 1 });
  assert.equal(completed.summary.status, 'COMPLETED');
  assert.equal((await stat(completed.directory)).isDirectory(), true);
  assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
  assert.deepEqual((await readdir(completed.directory)).sort(), ['changed.csv', 'missing.csv', 'summary.json']);
  assert.equal(parseCsv(await readFile(join(completed.directory, 'changed.csv'), 'utf8')).length, 2);
  assert.equal(parseCsv(await readFile(join(completed.directory, 'missing.csv'), 'utf8')).length, 2);
});

test('createReportWriter preserves empty and non-empty final path collisions', async (t) => {
  const root = await makeTempDir();
  t.after(() => rm(root, { recursive: true, force: true }));
  for (const kind of ['empty', 'non-empty']) {
    const directory = join(root, kind);
    await mkdir(directory);
    const writer = await createReportWriter(reportSpec(directory, ['A', 'B']));
    const [stagingName] = (await readdir(directory)).filter((name) => name.endsWith('.tmp'));
    const initialRunId = stagingName.slice(1, -4);
    const collision = join(directory, initialRunId);
    await mkdir(collision);
    if (kind === 'non-empty') await writeFile(join(collision, 'sentinel.txt'), 'preserve me', 'utf8');

    const completed = await writer.complete({ files: 2 });

    assert.notEqual(completed.summary.runId, initialRunId);
    assert.equal(completed.directory, join(directory, completed.summary.runId));
    assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
    assert.equal((await stat(collision)).isDirectory(), true);
    if (kind === 'non-empty') assert.equal(await readFile(join(collision, 'sentinel.txt'), 'utf8'), 'preserve me');
  }
});

test('createReportWriter rechecks its final path after writing a large summary', { timeout: 10_000 }, async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const writer = await createReportWriter(reportSpec(directory, ['A', 'B']));
  const [stagingName] = (await readdir(directory)).filter((name) => name.endsWith('.tmp'));
  const initialRunId = stagingName.slice(1, -4);
  const staging = join(directory, stagingName);
  const collision = join(directory, initialRunId);
  const createCollision = async () => {
    while (true) {
      try {
        await stat(join(staging, 'summary.json'));
        mkdirSync(collision);
        return;
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
        await new Promise((resolve) => setImmediate(resolve));
      }
    }
  };

  const [completed] = await Promise.all([
    writer.complete({ files: 2, padding: 'x'.repeat(20 * 1024 * 1024) }),
    createCollision()
  ]);

  assert.notEqual(completed.summary.runId, initialRunId);
  assert.deepEqual(await readdir(collision), []);
  assert.equal(completed.directory, join(directory, completed.summary.runId));
  assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
});

test('abort reselects a run ID when its initial failure path is occupied', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const writer = await createReportWriter(reportSpec(directory, ['A', 'B']));
  const [stagingName] = (await readdir(directory)).filter((name) => name.endsWith('.tmp'));
  const initialRunId = stagingName.slice(1, -4);
  const collision = join(directory, initialRunId);
  await mkdir(collision);
  await writeFile(join(collision, 'sentinel.txt'), 'preserve me', 'utf8');

  const aborted = await writer.abort(Object.assign(new Error('failed'), { code: 'INPUT_ERROR' }));

  assert.notEqual(aborted.summary.runId, initialRunId);
  assert.equal(aborted.directory, join(directory, aborted.summary.runId));
  assert.deepEqual(await readdir(aborted.directory), ['summary.json']);
  assert.equal(await readFile(join(collision, 'sentinel.txt'), 'utf8'), 'preserve me');
  assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
});

test('header initialization failures remain abortable and publish only FAILED summaries', async (t) => {
  const root = await makeTempDir();
  t.after(() => rm(root, { recursive: true, force: true }));
  for (const kind of ['sync', 'async']) {
    const directory = join(root, kind);
    const headerError = Object.assign(new Error(`SECRET-${kind}-HEADER`), { code: kind === 'sync' ? 'STREAM_INIT' : 'EIO' });
    const writer = await createReportWriter(reportSpec(directory, ['A', 'B']), {
      createStream: kind === 'sync'
        ? () => { throw headerError; }
        : () => new Writable({ write(_chunk, _encoding, callback) { setImmediate(callback, headerError); } })
    });

    await assert.rejects(() => writer.complete({ files: 2 }), (error) => error === headerError);
    const aborted = await writer.abort(headerError);
    const summaryText = await readFile(join(aborted.directory, 'summary.json'), 'utf8');
    assert.deepEqual(await readdir(aborted.directory), ['summary.json']);
    assert.equal(aborted.summary.code, headerError.code);
    assert.doesNotMatch(summaryText, /SECRET-/);
    assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
  }
});

test('createReportWriter aborts failed writes into a redacted summary-only report', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const writeError = Object.assign(new Error('SECRET-WRITE-VALUE'), { code: 'WRITE_FAILED' });
  let writes = 0;
  const writer = await createReportWriter(reportSpec(directory, ['A', 'B']), {
    createStream: () => new Writable({
      write(_chunk, _encoding, callback) {
        writes += 1;
        callback(writes === 3 ? writeError : null);
      }
    })
  });

  await assert.rejects(() => writer.onChanged({
    key: [typed('string', 'SECRET-KEY')], sheetName: 'Data', column: 'value',
    files: {
      A: { value: typed('string', 'before'), rowNumber: 2 },
      B: { value: typed('string', 'after'), rowNumber: 2 }
    }
  }), (error) => error === writeError);
  const aborted = await writer.abort(writeError);
  const summaryText = await readFile(join(aborted.directory, 'summary.json'), 'utf8');

  assert.deepEqual(await readdir(aborted.directory), ['summary.json']);
  assert.equal((await readdir(directory)).some((name) => name.endsWith('.tmp')), false);
  assert.equal(aborted.summary.status, 'FAILED');
  assert.equal(aborted.summary.code, 'WRITE_FAILED');
  assert.doesNotMatch(summaryText, /SECRET-WRITE-VALUE|SECRET-KEY|before|after/);
  assert.equal(summaryText.endsWith('\n'), true);
});

test('abort scrubs staged details and surfaces an unwritable-parent failure', { skip: process.platform === 'win32' }, async (t) => {
  const directory = await makeTempDir();
  t.after(async () => {
    await chmod(directory, 0o755).catch(() => {});
    await rm(directory, { recursive: true, force: true });
  });
  const writer = await createReportWriter(reportSpec(directory, ['A', 'B']));
  await writer.onChanged({
    key: [typed('string', 'SECRET-KEY')], sheetName: 'Data', column: 'value',
    files: {
      A: { value: typed('string', 'SECRET-BEFORE'), rowNumber: 2 },
      B: { value: typed('string', 'SECRET-AFTER'), rowNumber: 2 }
    }
  });
  await chmod(directory, 0o555);

  await assert.rejects(() => writer.abort(new Error('comparison failed')), (error) => error.code === 'EACCES');
  await chmod(directory, 0o755);
  const staged = (await readdir(directory)).filter((name) => name.endsWith('.tmp'));
  for (const name of staged) {
    const files = await readdir(join(directory, name));
    for (const file of files) {
      assert.doesNotMatch(await readFile(join(directory, name, file), 'utf8'), /SECRET-/);
    }
  }
});

test('writeReport surfaces abort publication failures instead of the original detail error', { skip: process.platform === 'win32' }, async (t) => {
  const directory = await makeTempDir();
  t.after(async () => {
    await chmod(directory, 0o755).catch(() => {});
    await rm(directory, { recursive: true, force: true });
  });
  const detailError = new Error('detail iteration failed');
  const changed = {
    [Symbol.iterator]() {
      chmodSync(directory, 0o555);
      throw detailError;
    }
  };

  await assert.rejects(
    () => writeReport(reportSpec(directory, ['A', 'B']), reportResult({ changed })),
    (error) => error.code === 'EACCES' && error.cause === detailError
  );
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

test('writeReport writes protected deterministic multiset output only for multiset mode', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = {
    ...reportSpec(directory, ['B', 'A', 'C']),
    mode: { type: 'multiset' }
  };
  const result = reportResult({
    multiset: [
      {
        values: [typed('string', 'ä')],
        sheetName: 'Data',
        counts: { A: 0, B: 1, C: 0 },
        baselineRelation: 'ADDED'
      },
      {
        values: [typed('string', 'z')],
        sheetName: 'Data',
        counts: { A: 2, B: 1, C: 0 },
        baselineRelation: 'DELETED'
      }
    ]
  });

  const report = await writeReport(spec, result);
  const multiset = parseCsv(await readFile(join(report.directory, 'multiset.csv'), 'utf8'));

  assert.deepEqual(multiset, [
    ['values', 'sheet', 'B.count', 'A.count', 'C.count', 'baselineRelation'],
    [JSON.stringify(['z']), 'Data', '1', '2', '0', 'DELETED'],
    [JSON.stringify(['ä']), 'Data', '1', '0', '0', 'ADDED']
  ]);
  assert.deepEqual(report.summary.artifacts, {
    multiset: 'multiset.csv'
  });
  await assert.rejects(() => readFile(join(report.directory, 'changed.csv'), 'utf8'));
  await assert.rejects(() => readFile(join(report.directory, 'missing.csv'), 'utf8'));

  const keyReport = await writeReport(reportSpec(directory), reportResult());
  await assert.rejects(() => readFile(join(keyReport.directory, 'multiset.csv'), 'utf8'));
});
