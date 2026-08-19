import assert from 'node:assert/strict';
import { lstat, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const root = fileURLToPath(new URL('..', import.meta.url));
const skillPaths = [
  'skill/SKILL.md',
  '.agents/skills/excel-diff/SKILL.md',
  'plugins/excel-diff/skills/excel-diff/SKILL.md'
];

async function json(path) {
  return JSON.parse(await readFile(join(root, path), 'utf8'));
}

test('plugin and marketplace metadata match the package and synced skill', async () => {
  const [pkg, plugin, marketplace] = await Promise.all([
    json('package.json'),
    json('plugins/excel-diff/.claude-plugin/plugin.json'),
    json('.claude-plugin/marketplace.json')
  ]);
  const entry = marketplace.plugins[0];

  assert.equal(plugin.name, 'excel-diff');
  assert.equal(plugin.version, pkg.version);
  assert.equal(marketplace.name, 'excel-diff-tools');
  assert.equal(marketplace.plugins.length, 1);
  assert.equal(entry.name, plugin.name);
  assert.equal(entry.version, pkg.version);
  assert.equal(entry.source, './plugins/excel-diff');
  assert.equal(entry.description, plugin.description);
  assert.equal((await lstat(join(root, entry.source))).isDirectory(), true);

  const skills = await Promise.all(skillPaths.map((path) => readFile(join(root, path))));
  assert.deepEqual(skills[1], skills[0]);
  assert.deepEqual(skills[2], skills[0]);
  for (const path of skillPaths) {
    assert.equal((await lstat(join(root, path))).isSymbolicLink(), false);
  }
});
