#!/usr/bin/env node
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const usage = 'usage: node scripts/sync-skills.mjs [--check] [--root <path>]';
const defaultRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const targets = [
  '.agents/skills/excel-diff/SKILL.md',
  'plugins/excel-diff/skills/excel-diff/SKILL.md'
];

function parseArgs(args) {
  let check = false;
  let root;
  for (let index = 0; index < args.length; index += 1) {
    if (args[index] === '--check' && !check) check = true;
    else if (args[index] === '--root' && root === undefined) {
      root = args[index + 1];
      if (!root || root.startsWith('--')) throw new Error(usage);
      index += 1;
    } else throw new Error(usage);
  }
  return { check, root: root === undefined ? defaultRoot : resolve(root) };
}

async function main() {
  const { check, root } = parseArgs(process.argv.slice(2));
  const source = resolve(root, 'skill/SKILL.md');
  const expected = await readFile(source);
  if (check) {
    for (const target of targets) {
      let actual;
      try { actual = await readFile(resolve(root, target)); } catch {}
      if (!actual?.equals(expected)) throw new Error(`out of sync: ${target}`);
    }
    return;
  }
  for (const target of targets) {
    const destination = resolve(root, target);
    await mkdir(dirname(destination), { recursive: true });
    await writeFile(destination, expected);
  }
}

main().catch((error) => {
  process.stderr.write(`${error.message === usage ? usage : error.message}\n`);
  process.exitCode = error.message === usage ? 2 : 1;
});
