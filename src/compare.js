import { equalValues, encodeKey } from './normalize.js';
import { readFileRows } from './read-xlsx.js';

export class CompareError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'CompareError';
    this.code = code;
  }
}

function ruleForColumn(spec, column) {
  const normalization = spec.normalization ?? {};
  return {
    emptyEqualsNull: false,
    caseSensitive: true,
    formulaMode: 'formula-and-cached-result',
    ...normalization,
    ...(normalization.columns?.[column] ?? {})
  };
}

export async function compare(spec) {
  const baselineFile = spec.files.find(({ id }) => id === spec.baseline);
  const results = new Map();
  results.set(baselineFile.id, await readFileRows(baselineFile, spec));
  const columns = results.get(baselineFile.id).columns;
  for (const file of spec.files) {
    if (file.id !== baselineFile.id) results.set(file.id, await readFileRows(file, spec, columns));
  }

  const entries = new Map();
  let totalRowsScanned = 0;
  let matchedRows = 0;
  let invalidRows = 0;
  for (const file of spec.files) {
    const result = results.get(file.id);
    totalRowsScanned += result.totalRowsScanned;
    invalidRows += result.invalidRows;
    matchedRows += result.rows.length;
    for (const row of result.rows) {
      const key = spec.mode.keyColumns.map((column) => row.values[column]);
      const encoded = encodeKey(key);
      if (!entries.has(encoded)) entries.set(encoded, { key, rows: new Map() });
      const byFile = entries.get(encoded).rows;
      if (!byFile.has(file.id)) byFile.set(file.id, []);
      byFile.get(file.id).push(row);
    }
  }

  const compareColumns = (spec.compareColumns === '*' ? columns : spec.compareColumns)
    .filter((column) => !spec.mode.keyColumns.includes(column));
  const changed = [];
  const missing = [];
  const duplicates = [];
  let identicalKeys = 0;
  let changedKeys = 0;
  for (const [, entry] of [...entries].sort(([left], [right]) => (left > right) - (left < right))) {
    const duplicateFiles = spec.files
      .filter(({ id }) => (entry.rows.get(id)?.length ?? 0) > 1)
      .map(({ id }) => id);
    if (duplicateFiles.length > 0) {
      if (spec.duplicateKeyPolicy === 'fail') {
        throw new CompareError('DUPLICATE_KEY', `duplicate business key in files ${duplicateFiles.join(', ')}`);
      }
      duplicates.push({ key: entry.key, files: duplicateFiles });
      continue;
    }

    const presentFiles = spec.files.filter(({ id }) => entry.rows.has(id)).map(({ id }) => id);
    if (presentFiles.length !== spec.files.length) {
      missing.push({
        key: entry.key,
        sheetName: entry.rows.values().next().value[0].sheetName,
        presentFiles,
        missingFiles: spec.files.filter(({ id }) => !entry.rows.has(id)).map(({ id }) => id),
        baselineRelation: entry.rows.has(spec.baseline) ? 'DELETED' : 'ADDED'
      });
      continue;
    }

    const baseline = entry.rows.get(spec.baseline)[0];
    let keyChanged = false;
    for (const column of compareColumns) {
      const rule = ruleForColumn(spec, column);
      if (spec.files.every(({ id }) => equalValues(baseline.values[column], entry.rows.get(id)[0].values[column], rule))) continue;
      keyChanged = true;
      changed.push({
        key: entry.key,
        sheetName: baseline.sheetName,
        column,
        files: Object.fromEntries(spec.files.map(({ id }) => {
          const row = entry.rows.get(id)[0];
          return [id, { value: row.values[column], rowNumber: row.rowNumber }];
        }))
      });
    }
    if (keyChanged) changedKeys += 1;
    else identicalKeys += 1;
  }

  return {
    summary: {
      files: spec.files.length,
      totalRowsScanned,
      matchedRows,
      identicalKeys,
      changedKeys,
      missingKeys: missing.length,
      duplicateKeys: duplicates.length,
      invalidRows
    },
    changed,
    missing,
    duplicates
  };
}
