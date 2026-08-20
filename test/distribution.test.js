import assert from 'node:assert/strict';
import { cp, lstat, readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import { makeTempDir } from './helpers.js';

const root = fileURLToPath(new URL('..', import.meta.url));
const skillPaths = [
  'skill/SKILL.md',
  '.agents/skills/excel-diff/SKILL.md',
  'plugins/excel-diff/skills/excel-diff/SKILL.md'
];
const schemaPaths = [
  'schemas/compare-spec.schema.json',
  'skill/references/compare-spec.schema.json',
  '.agents/skills/excel-diff/references/compare-spec.schema.json',
  'plugins/excel-diff/skills/excel-diff/references/compare-spec.schema.json'
];

async function json(path) {
  return JSON.parse(await readFile(join(root, path), 'utf8'));
}

test('package metadata and README require Node.js 22 or newer', async () => {
  const [pkg, lock, readme] = await Promise.all([
    json('package.json'),
    json('package-lock.json'),
    readFile(join(root, 'README.md'), 'utf8')
  ]);

  assert.equal(pkg.engines.node, '>=22');
  assert.equal(lock.packages[''].engines.node, '>=22');
  assert.match(readme, /需要 Node\.js 22 或更高版本/);
  assert.doesNotMatch(readme, /Node\.js 24|node@24|install 24|use 24|major < 24/);
});

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

test('a cache-only plugin keeps its referenced CompareSpec schema', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const cache = join(directory, 'excel-diff');
  await cp(join(root, 'plugins/excel-diff'), cache, { recursive: true });

  const cachedSkill = await readFile(join(cache, 'skills/excel-diff/SKILL.md'), 'utf8');
  const cachedSchema = await readFile(join(cache, 'skills/excel-diff/references/compare-spec.schema.json'));
  assert.match(cachedSkill, /\]\(references\/compare-spec\.schema\.json\)/);
  assert.doesNotThrow(() => JSON.parse(cachedSchema.toString('utf8')));
  assert.deepEqual(cachedSchema, await readFile(join(root, schemaPaths[0])));

  for (const path of schemaPaths) {
    const metadata = await lstat(join(root, path));
    assert.equal(metadata.isFile(), true);
    assert.equal(metadata.isSymbolicLink(), false);
  }
});
