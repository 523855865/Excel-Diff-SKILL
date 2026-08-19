import assert from 'node:assert/strict';
import { access, chmod, mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { dirname, isAbsolute, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';

import { makeTempDir, writeWorkbook } from './helpers.js';

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
      artifacts: { changed: 'changed.csv', missing: 'missing.csv' }
    }
  );
  await Promise.all(['summary.json', 'changed.csv', 'missing.csv'].map((file) => access(join(output.directory, file), constants.F_OK)));
  assert.equal((await stat(output.directory)).isDirectory(), true);
  assert.deepEqual(
    Object.fromEntries(Object.keys(output).filter((key) => key !== 'directory').map((key) => [key, output[key]])),
    JSON.parse(await readFile(join(output.directory, 'summary.json'), 'utf8'))
  );
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

test('CLI surfaces an abort publication failure as a redacted internal error', { skip: process.platform === 'win32' }, async (t) => {
  const rows = [
    ['编号', '姓名'],
    ...Array.from({ length: 1001 }, () => ['SECRET-CLI-KEY', 'SECRET-CLI-VALUE'])
  ];
  const directory = await fixtures(t, rows);
  const spec = await writeSpec(directory, { duplicateKeyPolicy: 'fail' });
  const outputDirectory = join(directory, 'output');
  t.after(() => chmod(outputDirectory, 0o755).catch(() => {}));

  const result = await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [cli, 'compare', '--progress', '--spec', spec], {
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

test('returns stable partition resource errors instead of INTERNAL_ERROR', async (t) => {
  const directory = await fixtures(t);
  for (const [code, resources] of [
    ['TEMP_LIMIT_EXCEEDED', { maxTempBytes: 1 }],
    ['HOT_KEY_TOO_LARGE', { maxPartitionBytes: 1 }]
  ]) {
    const spec = await writeSpec(directory, { resources });
    const result = run(['compare', '--spec', spec]);

    assert.equal(result.status, 4);
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
