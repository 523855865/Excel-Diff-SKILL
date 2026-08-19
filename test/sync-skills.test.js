import assert from 'node:assert/strict';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

import { makeTempDir } from './helpers.js';

const script = fileURLToPath(new URL('../scripts/sync-skills.mjs', import.meta.url));
const targets = [
  '.agents/skills/excel-diff/SKILL.md',
  'plugins/excel-diff/skills/excel-diff/SKILL.md'
];

function run(...args) {
  return spawnSync(process.execPath, [script, ...args], { encoding: 'utf8' });
}

test('check fails until sync creates byte-identical copies and detects drift', async (t) => {
  const root = await makeTempDir();
  t.after(() => rm(root, { recursive: true, force: true }));
  const source = join(root, 'skill', 'SKILL.md');
  const bytes = Buffer.from('---\nname: excel-diff\n---\n\u0000utf8: \u4eba\n');
  await mkdir(dirname(source), { recursive: true });
  await writeFile(source, bytes);

  const missing = run('--check', '--root', root);
  assert.notEqual(missing.status, 0);
  assert.match(missing.stderr, /out of sync/i);

  const synced = run('--root', root);
  assert.equal(synced.status, 0, synced.stderr);
  assert.equal(synced.stderr, '');
  for (const target of targets) {
    assert.deepEqual(await readFile(join(root, target)), bytes);
  }

  const checked = run('--check', '--root', root);
  assert.equal(checked.status, 0, checked.stderr);
  assert.equal(checked.stderr, '');

  await writeFile(join(root, targets[0]), 'drift', 'utf8');
  const drifted = run('--check', '--root', root);
  assert.notEqual(drifted.status, 0);
  assert.match(drifted.stderr, /out of sync/i);
});

test('rejects unknown, duplicate, and missing-value arguments', () => {
  for (const args of [
    ['--wat'],
    ['--check', '--check'],
    ['--root'],
    ['--root', 'a', '--root', 'b']
  ]) {
    const result = run(...args);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /^usage: /i);
  }
});
