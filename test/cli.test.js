import assert from 'node:assert/strict';
import { access, chmod, mkdir, readFile, realpath, readdir, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { dirname, isAbsolute, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';

import { makeTempDir, writeWorkbook } from './helpers.js';
import { failure as classifyFailure, main as runMain } from '../src/cli.js';
import { createReportWriter } from '../src/report.js';

const cli = fileURLToPath(new URL('../src/cli.js', import.meta.url));

function run(args, environment = {}) {
  return spawnSync(process.execPath, [cli, ...args], {
    encoding: 'utf8',
    env: { ...process.env, ...environment }
  });
}

async function writeSpec(directory, overrides = {}) {
  const specPath = join(directory, 'rules', 'compare.json');
  const spec = {
    version: '1.0',
    baseline: 'before',
    files: [
      { id: 'before', path: '../before.xlsx' },
      { id: 'after', path: '../after.xlsx' }
    ],
    sheet: { name: '人员', headerRow: 1 },
    mode: { type: 'key', keyColumns: ['编号'] },
    compareColumns: '*',
    duplicateKeyPolicy: 'report',
    output: { directory: '../output' },
    ...overrides
  };
  await mkdir(dirname(specPath), { recursive: true });
  await writeFile(specPath, JSON.stringify(spec), 'utf8');
  return specPath;
}

async function fixtures(t, rows = [
  ['编号', '姓名'],
  ['1', 'Ada']
]) {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  await writeWorkbook(join(directory, 'before.xlsx'), '人员', rows);
  await writeWorkbook(join(directory, 'after.xlsx'), '人员', rows);
  return directory;
}

function jsonLine(text) {
  assert.match(text, /^[^\r\n]+\n$/);
  const value = JSON.parse(text.slice(0, -1));
  assert.equal(typeof value, 'object');
  assert.notEqual(value, null);
  assert.equal(Array.isArray(value), false);
  return value;
}

function response(result) {
  assert.equal(result.stderr, '');
  return jsonLine(result.stdout);
}

function failure(result, code) {
  assert.equal(result.stdout, '');
  const value = jsonLine(result.stderr);
  assert.deepEqual(Object.keys(value).sort(), ['code', 'message', 'status']);
  assert.equal(value.status, 'FAILED');
  assert.equal(value.code, code);
  return value;
}

test('classifies raw disk exhaustion errors as redacted exit 5 failures', () => {
  for (const code of ['ENOSPC', 'EDQUOT']) {
    const error = Object.assign(new Error(`SECRET-${code}`), { code });
    assert.deepEqual(classifyFailure(error), {
      exitCode: 5,
      output: { status: 'FAILED', code: 'DISK_FULL', message: 'output storage is full' }
    });
  }
});

test('imports without an argv file and still executes through a bin symlink', async (t) => {
  const imported = spawnSync(process.execPath, ['--input-type=module', '-'], {
    encoding: 'utf8',
    input: `await import(${JSON.stringify(pathToFileURL(cli).href)});\n`
  });
  assert.equal(imported.status, 0);
  assert.equal(imported.stderr, '');

  const directory = await fixtures(t);
  const spec = await writeSpec(directory);
  const link = join(directory, 'excel-diff-bin.js');
  await symlink(cli, link);
  const linked = spawnSync(process.execPath, [link, 'compare', '--spec', spec], { encoding: 'utf8' });
  assert.equal(linked.status, 0);
  response(linked);
});

test('compare emits one completed JSON summary and report artifacts', async (t) => {
  const directory = await fixtures(t);
  const spec = await writeSpec(directory);

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 0);
  const output = response(result);
  assert.equal(output.status, 'COMPLETED');
  assert.equal(isAbsolute(output.directory), true);
  assert.deepEqual(
    Object.fromEntries(Object.keys(output).filter((key) => key !== 'directory').map((key) => [key, output[key]])),
    {
      files: 2,
      totalRowsScanned: 2,
      matchedRows: 2,
      identicalKeys: 1,
      changedKeys: 0,
      missingKeys: 0,
      duplicateKeys: 0,
      invalidRows: 0,
      status: 'COMPLETED',
      runId: output.runId,
      artifacts: { changed: 'changed.csv', missing: 'missing.csv', duplicates: 'duplicate-keys.csv' }
    }
  );
  await Promise.all(['summary.json', 'changed.csv', 'missing.csv', 'duplicate-keys.csv'].map((file) => access(join(output.directory, file), constants.F_OK)));
  assert.equal((await stat(output.directory)).isDirectory(), true);
  assert.deepEqual(
    Object.fromEntries(Object.keys(output).filter((key) => key !== 'directory').map((key) => [key, output[key]])),
    JSON.parse(await readFile(join(output.directory, 'summary.json'), 'utf8'))
  );
});

test('inspect emits one structured JSON result for two workbooks', async (t) => {
  const directory = await fixtures(t);
  const before = join(directory, 'before.xlsx');
  const after = join(directory, 'after.xlsx');

  const result = run(['inspect', '--files', before, after, '--sheet', '人员']);

  assert.equal(result.status, 0);
  const output = response(result);
  assert.equal(output.status, 'INSPECTED');
  assert.equal(output.fullTypes, false);
  assert.deepEqual(output.files.map(({ file }) => file), [await realpath(before), await realpath(after)]);
  assert.deepEqual(output.files[0].headers.map(({ raw }) => raw), ['编号', '姓名']);
});

test('inspect maps --full-types to a complete type scan', async (t) => {
  const directory = await fixtures(t);

  const result = run([
    'inspect', '--full-types', '--sheet', '人员', '--files',
    join(directory, 'before.xlsx'), join(directory, 'after.xlsx')
  ]);

  assert.equal(result.status, 0);
  assert.equal(response(result).fullTypes, true);
});

test('inspect returns a value-free duplicate-header prompt on stdout with exit 3', async (t) => {
  const directory = await fixtures(t, [
    ['Ａ', 'A', '唯一'],
    ['SECRET-LEFT', 'SECRET-RIGHT', 'SECRET-KEY']
  ]);
  const before = join(directory, 'before.xlsx');
  const after = join(directory, 'after.xlsx');

  const result = run(['inspect', '--files', before, after, '--sheet', '人员']);

  assert.equal(result.status, 3);
  const output = response(result);
  const files = await Promise.all([before, after].map((file) => realpath(file)));
  assert.deepEqual(output, {
    status: 'NEEDS_INPUT',
    code: 'HEADER_DUPLICATED',
    files: files.map((file) => ({
      file,
      sheet: '人员',
      duplicates: [{
        normalized: 'A',
        columns: [
          { index: 1, raw: 'Ａ' },
          { index: 2, raw: 'A' }
        ]
      }]
    }))
  });
  assert.doesNotMatch(result.stdout, /SECRET-/);
});

test('inspect reports a missing sheet on stderr with exit 4', async (t) => {
  const directory = await fixtures(t);

  const result = run([
    'inspect', '--files', join(directory, 'before.xlsx'), join(directory, 'after.xlsx'),
    '--sheet', '不存在'
  ]);

  assert.equal(result.status, 4);
  failure(result, 'SHEET_NOT_FOUND');
});

test('inspect rejects invalid grammar and fewer than two files as usage errors', async (t) => {
  const directory = await fixtures(t);
  const before = join(directory, 'before.xlsx');
  const after = join(directory, 'after.xlsx');
  const cases = [
    ['inspect', '--files', before, '--sheet', '人员'],
    ['inspect', '--files', before, after],
    ['inspect', '--sheet', '人员', '--files'],
    ['inspect', '--files', before, after, '--sheet', '人员', '--sheet', '人员'],
    ['inspect', '--files', before, after, '--sheet', '人员', '--unknown'],
    ['inspect', '--files', before, after, '--sheet', '人员', '--full-types', '--full-types']
  ];

  for (const args of cases) {
    const result = run(args);
    assert.equal(result.status, 2, args.join(' '));
    failure(result, 'USAGE');
  }
});

test('progress is throttled and keep-temp works with optional flags in any order', async (t) => {
  const rows = [
    ['编号', '姓名'],
    ...Array.from({ length: 1001 }, (_, index) => [index === 0 ? 'SECRET-PROGRESS-KEY' : String(index), `name-${index}`])
  ];
  const directory = await fixtures(t, rows);
  const spec = await writeSpec(directory);

  const result = run(['compare', '--keep-temp', '--progress', '--spec', spec]);

  assert.equal(result.status, 0);
  const output = jsonLine(result.stdout);
  assert.equal(output.status, 'COMPLETED');
  assert.equal(isAbsolute(output.tempDirectory), true);
  await access(output.tempDirectory);
  t.after(() => rm(output.tempDirectory, { recursive: true, force: true }));
  const events = result.stderr.trim().split('\n').map((line) => JSON.parse(line));
  assert.deepEqual(events.map(({ rowsScanned }) => rowsScanned), [1000, 2000, 2002]);
  for (const event of events) {
    assert.deepEqual(Object.keys(event), ['bytesWritten', 'currentFile', 'rowsScanned', 'type']);
    assert.equal(event.type, 'PROGRESS');
  }
  assert.equal(events.every((event, index) => index === 0
    || (event.rowsScanned > events[index - 1].rowsScanned && event.bytesWritten >= events[index - 1].bytesWritten)), true);
  assert.doesNotMatch(result.stderr, /SECRET-PROGRESS-KEY|name-/);
});

test('comparison failures publish only a FAILED summary after spec loading', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = await writeSpec(directory);

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 4);
  failure(result, 'INPUT_ERROR');
  const outputDirectory = join(directory, 'output');
  const runs = await readdir(outputDirectory);
  assert.equal(runs.some((name) => name.endsWith('.tmp')), false);
  assert.equal(runs.length, 1);
  assert.deepEqual(await readdir(join(outputDirectory, runs[0])), ['summary.json']);
  const summary = JSON.parse(await readFile(join(outputDirectory, runs[0], 'summary.json'), 'utf8'));
  assert.equal(summary.status, 'FAILED');
  assert.equal(summary.code, 'INPUT_ERROR');
  assert.equal(Object.hasOwn(summary, 'artifacts'), false);
});

test('--keep-temp exposes the retained partition directory on comparison failure', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = await writeSpec(directory);

  const result = run(['compare', '--keep-temp', '--spec', spec]);

  assert.equal(result.status, 4);
  assert.equal(result.stdout, '');
  const output = jsonLine(result.stderr);
  assert.deepEqual(Object.keys(output).sort(), ['code', 'message', 'status', 'tempDirectory']);
  assert.equal(output.code, 'INPUT_ERROR');
  assert.equal(isAbsolute(output.tempDirectory), true);
  await access(output.tempDirectory);
  t.after(() => rm(output.tempDirectory, { recursive: true, force: true }));
  const runs = await readdir(join(directory, 'output'));
  const summary = JSON.parse(await readFile(join(directory, 'output', runs[0], 'summary.json'), 'utf8'));
  assert.equal(summary.tempDirectory, output.tempDirectory);
});

test('--keep-temp survives a report complete failure in stderr and FAILED summary', async (t) => {
  const directory = await fixtures(t);
  const specPath = await writeSpec(directory);
  const completeError = Object.assign(new Error('SECRET-COMPLETE-FAILURE'), { code: 'EIO' });

  await assert.rejects(
    () => runMain(['compare', '--keep-temp', '--spec', specPath], {
      async createReportWriter(spec) {
        const writer = await createReportWriter(spec);
        return { ...writer, complete: async () => { throw completeError; } };
      }
    }),
    (error) => error === completeError
  );
  assert.equal(isAbsolute(completeError.tempDirectory), true);
  await access(completeError.tempDirectory);
  t.after(() => rm(completeError.tempDirectory, { recursive: true, force: true }));
  const runs = await readdir(join(directory, 'output'));
  const summary = JSON.parse(await readFile(join(directory, 'output', runs[0], 'summary.json'), 'utf8'));
  assert.equal(summary.status, 'FAILED');
  assert.equal(summary.tempDirectory, completeError.tempDirectory);
  assert.deepEqual(classifyFailure(completeError).output, {
    status: 'FAILED', code: 'INTERNAL_ERROR', message: 'unexpected error',
    tempDirectory: completeError.tempDirectory
  });
});

test('CLI surfaces an abort publication failure as a redacted internal error', { skip: process.platform === 'win32' }, async (t) => {
  const rows = [
    ['编号', '姓名'],
    ...Array.from({ length: 1001 }, (_, index) => [`SECRET-CLI-KEY-${index}`, `SECRET-CLI-VALUE-${index}`])
  ];
  const directory = await fixtures(t, rows);
  const spec = await writeSpec(directory);
  const outputDirectory = join(directory, 'output');
  t.after(() => chmod(outputDirectory, 0o755).catch(() => {}));

  const result = await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [cli, 'compare', '--keep-temp', '--progress', '--spec', spec], {
      env: process.env
    });
    let stdout = '';
    let stderr = '';
    let stopped = false;
    child.stdout.setEncoding('utf8').on('data', (chunk) => { stdout += chunk; });
    child.stderr.setEncoding('utf8').on('data', async (chunk) => {
      stderr += chunk;
      if (!stopped && stderr.includes('\n')) {
        stopped = true;
        child.kill('SIGSTOP');
        try {
          await chmod(outputDirectory, 0o555);
          child.kill('SIGCONT');
        } catch (error) {
          child.kill('SIGCONT');
          reject(error);
        }
      }
    });
    child.on('error', reject);
    child.on('close', (status) => resolve({ status, stdout, stderr }));
  });
  await chmod(outputDirectory, 0o755);

  assert.equal(result.status, 6);
  assert.equal(result.stdout, '');
  const lines = result.stderr.trim().split('\n').map((line) => JSON.parse(line));
  assert.equal(lines.at(-1).status, 'FAILED');
  assert.equal(lines.at(-1).code, 'INTERNAL_ERROR');
  assert.equal(isAbsolute(lines.at(-1).tempDirectory), true);
  await access(lines.at(-1).tempDirectory);
  t.after(() => rm(lines.at(-1).tempDirectory, { recursive: true, force: true }));
  assert.doesNotMatch(result.stderr, /SECRET-CLI-KEY|SECRET-CLI-VALUE/);
  const staged = (await readdir(outputDirectory)).filter((name) => name.endsWith('.tmp'));
  for (const name of staged) assert.deepEqual(await readdir(join(outputDirectory, name)), []);
});

test('returns SPEC_INVALID for an invalid spec without a default stack', async (t) => {
  const directory = await fixtures(t);
  const spec = await writeSpec(directory, { unexpected: true });

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 2);
  const output = failure(result, 'SPEC_INVALID');
  assert.equal(Object.hasOwn(output, 'stack'), false);
});

test('returns input errors for missing workbooks and invalid date filters', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const spec = await writeSpec(directory);

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 4);
  failure(result, 'INPUT_ERROR');

  const dateDirectory = await fixtures(t, [
    ['编号', '日期'],
    ['001', new Date('2026-01-02T00:00:00.000Z')]
  ]);
  const dateSpec = await writeSpec(dateDirectory, {
    filters: [{ column: '日期', operator: 'eq', value: '2026-02-30T00:00:00Z' }]
  });
  const dateResult = run(['compare', '--spec', dateSpec]);

  assert.equal(dateResult.status, 4);
  const error = failure(dateResult, 'FILTER_INVALID');
  assert.doesNotMatch(error.message, /2026-02-30|2026-01-02|001/);
});

test('returns a redacted duplicate-key failure', async (t) => {
  const directory = await fixtures(t, [
    ['编号', '姓名'],
    ['SECRET-EMP-001', 'Ada'],
    ['SECRET-EMP-001', 'Ada again']
  ]);
  const spec = await writeSpec(directory, { duplicateKeyPolicy: 'fail' });

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 4);
  failure(result, 'DUPLICATE_KEY');
  assert.doesNotMatch(result.stderr, /SECRET-EMP-001/);
});

test('duplicate report policy publishes a protected duplicate-key locator', async (t) => {
  const directory = await fixtures(t, [
    ['编号', '姓名'],
    ['=DUPLICATE-KEY', 'Ada'],
    ['=DUPLICATE-KEY', 'Ada again']
  ]);
  const spec = await writeSpec(directory);

  const result = run(['compare', '--spec', spec]);

  assert.equal(result.status, 0);
  const output = response(result);
  assert.equal(output.artifacts.duplicates, 'duplicate-keys.csv');
  const duplicateCsv = await readFile(join(output.directory, 'duplicate-keys.csv'), 'utf8');
  assert.match(duplicateCsv, /^key,files\n/);
  assert.match(duplicateCsv, /DUPLICATE-KEY/);
  assert.doesNotMatch(duplicateCsv, /Ada/);
  assert.doesNotMatch(duplicateCsv.split('\n')[1], /^[=+\-@]/);
  assert.doesNotMatch(result.stdout + result.stderr, /=DUPLICATE-KEY/);
});

test('returns exit 5 for stable resource errors instead of INTERNAL_ERROR', async (t) => {
  const directory = await fixtures(t);
  for (const [code, resources] of [
    ['ROW_LIMIT_EXCEEDED', { maxRows: 1 }],
    ['TEMP_LIMIT_EXCEEDED', { maxTempBytes: 1 }],
    ['HOT_KEY_TOO_LARGE', { maxPartitionBytes: 1 }]
  ]) {
    const spec = await writeSpec(directory, { resources });
    const result = run(['compare', '--spec', spec]);

    assert.equal(result.status, 5);
    const output = failure(result, code);
    assert.notEqual(output.code, 'INTERNAL_ERROR');
  }
});

test('rejects unknown commands and missing --spec as usage errors', () => {
  const unknown = run(['unknown', '--spec', 'ignored.json']);
  const missing = run(['compare']);

  assert.equal(unknown.status, 2);
  failure(unknown, 'USAGE');
  assert.equal(missing.status, 2);
  failure(missing, 'USAGE');
});

test('includes a stack only with EXCEL_DIFF_DEBUG=1', async (t) => {
  const directory = await fixtures(t);
  const spec = await writeSpec(directory, { unexpected: true });

  const result = run(['compare', '--spec', spec], { EXCEL_DIFF_DEBUG: '1' });

  assert.equal(result.status, 2);
  assert.equal(result.stdout, '');
  const output = jsonLine(result.stderr);
  assert.match(output.stack, /SpecError/);
});
