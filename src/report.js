import { randomUUID } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

export function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[,"\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function csv(rows) {
  return `${rows.map((row) => row.map(csvCell).join(',')).join('\n')}\n`;
}

function safeText(text) {
  return /^(?:[=+\-@\t\r\n]|json:)/.test(text) ? `json:${JSON.stringify(text)}` : text;
}

function untypedValue([type, value]) {
  if (type === 'blank') return '';
  if (type === 'formula' && Array.isArray(value)) {
    return { formula: value[0], result: untypedValue(value[1]) };
  }
  if (type === 'hyperlink') return { ...value, text: untypedValue(value.text) };
  return value;
}

function safeValue(value) {
  return safeText(value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value ?? ''));
}

function displayKey(key) {
  const values = key.map(untypedValue);
  return safeValue(values.length === 1 ? values[0] : values);
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
    ['key', 'sheet', 'column', ...fileIds.flatMap((fileId) => [safeText(`${fileId}.value`), safeText(`${fileId}.type`), safeText(`${fileId}.row`)])],
    ...changed.map((entry) => [
      displayKey(entry.key), safeText(entry.sheetName), safeText(entry.column),
      ...fileIds.flatMap((fileId) => {
        const value = entry.files[fileId].value;
        return [safeValue(untypedValue(value)), safeText(value[0]), entry.files[fileId].rowNumber];
      })
    ])
  ];
  const missingRows = [
    ['key', 'sheet', 'presentFiles', 'missingFiles', 'baselineRelation'],
    ...result.missing.map((entry) => [
      displayKey(entry.key), safeText(entry.sheetName), JSON.stringify(entry.presentFiles), JSON.stringify(entry.missingFiles), entry.baselineRelation
    ])
  ];
  const multisetRows = [
    ['values', 'sheet', ...fileIds.map((fileId) => safeText(`${fileId}.count`)), 'baselineRelation'],
    ...[...(result.multiset ?? [])]
      .sort((left, right) => {
        const leftValues = JSON.stringify(left.values);
        const rightValues = JSON.stringify(right.values);
        return (leftValues > rightValues) - (leftValues < rightValues);
      })
      .map((entry) => [
        safeValue(entry.values.map(untypedValue)),
        safeText(entry.sheetName),
        ...fileIds.map((fileId) => entry.counts[fileId]),
        entry.baselineRelation
      ])
  ];
  const artifacts = { changed: 'changed.csv', missing: 'missing.csv' };
  if (spec.mode?.type === 'multiset') artifacts.multiset = 'multiset.csv';
  const summary = {
    ...result.summary,
    status: 'COMPLETED',
    runId: id,
    artifacts
  };

  await mkdir(directory, { recursive: true });
  await writeFile(join(directory, 'changed.csv'), csv(changedRows), 'utf8');
  await writeFile(join(directory, 'missing.csv'), csv(missingRows), 'utf8');
  if (spec.mode?.type === 'multiset') await writeFile(join(directory, 'multiset.csv'), csv(multisetRows), 'utf8');
  await writeFile(join(directory, 'summary.json'), `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  return { directory, summary };
}
