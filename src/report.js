import { randomUUID } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

import { printValue } from './normalize.js';

export function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[,"\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function csv(rows) {
  return `${rows.map((row) => row.map(csvCell).join(',')).join('\n')}\n`;
}

function runId() {
  return `${new Date().toISOString().replace(/\.\d{3}Z$/, 'Z').replace(/[-:]/g, '')}-${randomUUID().slice(0, 8)}`;
}

export async function writeReport(spec, result) {
  const id = runId();
  const directory = join(spec.output.directory, id);
  const fileIds = spec.files.map(({ id: fileId }) => fileId);
  const changed = [...result.changed].sort((left, right) => {
    const leftKey = JSON.stringify(left.key);
    const rightKey = JSON.stringify(right.key);
    return (leftKey > rightKey) - (leftKey < rightKey)
      || (left.column > right.column) - (left.column < right.column);
  });
  const changedRows = [
    ['key', 'sheet', 'column', ...fileIds.flatMap((fileId) => [`${fileId}.value`, `${fileId}.row`])],
    ...changed.map((entry) => [
      JSON.stringify(entry.key), entry.sheetName, entry.column,
      ...fileIds.flatMap((fileId) => [printValue(entry.files[fileId].value), entry.files[fileId].rowNumber])
    ])
  ];
  const missingRows = [
    ['key', 'sheet', 'presentFiles', 'missingFiles', 'baselineRelation'],
    ...result.missing.map((entry) => [
      JSON.stringify(entry.key), entry.sheetName, JSON.stringify(entry.presentFiles), JSON.stringify(entry.missingFiles), entry.baselineRelation
    ])
  ];
  const summary = {
    ...result.summary,
    status: 'COMPLETED',
    runId: id,
    artifacts: { changed: 'changed.csv', missing: 'missing.csv' }
  };

  await mkdir(directory, { recursive: true });
  await writeFile(join(directory, 'changed.csv'), csv(changedRows), 'utf8');
  await writeFile(join(directory, 'missing.csv'), csv(missingRows), 'utf8');
  await writeFile(join(directory, 'summary.json'), `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  return { directory, summary };
}
