import assert from 'node:assert/strict';
import { access, mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { dirname, isAbsolute, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
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
