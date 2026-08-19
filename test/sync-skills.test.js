import assert from 'node:assert/strict';
import { access, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

import { makeTempDir } from './helpers.js';

const script = fileURLToPath(new URL('../scripts/sync-skills.mjs', import.meta.url));
const skillTargets = [
  '.agents/skills/excel-diff/SKILL.md',
  'plugins/excel-diff/skills/excel-diff/SKILL.md'
];
const schemaTargets = [
  'skill/references/compare-spec.schema.json',
  '.agents/skills/excel-diff/references/compare-spec.schema.json',
  'plugins/excel-diff/skills/excel-diff/references/compare-spec.schema.json'
];

function run(...args) {
  return spawnSync(process.execPath, [script, ...args], { encoding: 'utf8' });
}

async function writeSources(root, { skill = true, schema = true } = {}) {
  const skillBytes = Buffer.from('---\nname: excel-diff\n---\n\u0000utf8: \u4eba\n');
  const schemaBytes = Buffer.from('{"type":"object"}\n');
  if (skill) {
    await mkdir(join(root, 'skill'), { recursive: true });
    await writeFile(join(root, 'skill', 'SKILL.md'), skillBytes);
  }
  if (schema) {
    await mkdir(join(root, 'schemas'), { recursive: true });
    await writeFile(join(root, 'schemas', 'compare-spec.schema.json'), schemaBytes);
  }
  return { skillBytes, schemaBytes };
}

test('check fails until sync creates byte-identical copies and detects drift', async (t) => {
  const root = await makeTempDir();
  t.after(() => rm(root, { recursive: true, force: true }));
  const { skillBytes, schemaBytes } = await writeSources(root);

  const missing = run('--check', '--root', root);
  assert.notEqual(missing.status, 0);
  assert.match(missing.stderr, /out of sync/i);

  const synced = run('--root', root);
  assert.equal(synced.status, 0, synced.stderr);
  assert.equal(synced.stderr, '');
  for (const target of skillTargets) assert.deepEqual(await readFile(join(root, target)), skillBytes);
  for (const target of schemaTargets) assert.deepEqual(await readFile(join(root, target)), schemaBytes);

  const checked = run('--check', '--root', root);
  assert.equal(checked.status, 0, checked.stderr);
  assert.equal(checked.stderr, '');

  for (const target of schemaTargets) {
    await writeFile(join(root, target), 'drift', 'utf8');
    const drifted = run('--check', '--root', root);
    assert.notEqual(drifted.status, 0);
    assert.match(drifted.stderr, /out of sync/i);
    await writeFile(join(root, target), schemaBytes);
  }
});

test('write mode creates no target directories when either source is missing', async (t) => {
  for (const sources of [{ skill: false }, { schema: false }]) {
    const root = await makeTempDir();
    t.after(() => rm(root, { recursive: true, force: true }));
    await writeSources(root, sources);

    const result = run('--root', root);

    assert.notEqual(result.status, 0);
    for (const directory of ['skill/references', '.agents', 'plugins']) {
      await assert.rejects(access(join(root, directory)), { code: 'ENOENT' });
    }
  }
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
