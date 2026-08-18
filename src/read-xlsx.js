import ExcelJS from 'exceljs';

import { matchesFilter, normalizeValue } from './normalize.js';

export class InputError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'InputError';
    this.code = code;
  }
}

function requiredColumns(spec) {
  const columns = [...spec.mode.keyColumns];
  if (spec.compareColumns !== '*') columns.push(...spec.compareColumns);
  columns.push(...spec.filters.map(({ column }) => column));
  return [...new Set(columns)];
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

function readHeaders(sheet, headerRow) {
  const row = sheet.getRow(headerRow);
  const headers = [];
  const seen = new Map();
  for (let index = 1; index <= sheet.columnCount; index += 1) {
    const text = row.getCell(index).text;
    if (text === '') continue;
    const normalized = text.normalize('NFKC');
    const duplicate = seen.get(normalized);
    if (duplicate) {
      throw new InputError('HEADER_DUPLICATED', `duplicate header ${duplicate.text} at columns ${duplicate.index} and ${index}`);
    }
    const header = { text, normalized, index };
    seen.set(normalized, header);
    headers.push(header);
  }
  return headers;
}

function mappingRank(header, standard, aliases) {
  if (header.text === standard) return 0;
  if (header.normalized === standard.normalize('NFKC')) return 1;
  if ((aliases[standard] ?? []).some((alias) => header.text === alias || header.normalized === alias.normalize('NFKC'))) return 2;
  return -1;
}

function mapColumns(headers, standards, aliases) {
  const byStandard = new Map();
  const byHeader = new Map(headers.map((header) => [header.index, []]));
  for (const header of headers) {
    for (const standard of standards) {
      const rank = mappingRank(header, standard, aliases);
      if (rank < 0) continue;
      byHeader.get(header.index).push({ standard, rank });
    }
  }
  for (const header of headers) {
    const candidates = byHeader.get(header.index);
    if (candidates.length === 0) continue;
    const best = Math.min(...candidates.map(({ rank }) => rank));
    const matches = candidates.filter(({ rank }) => rank === best);
    if (matches.length > 1) {
      throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `source column ${header.text} maps to multiple standard columns`);
    }
    const { standard, rank } = matches[0];
    const sources = byStandard.get(standard) ?? [];
    sources.push({ header, rank });
    byStandard.set(standard, sources);
  }
  const mapping = new Map();
  for (const [standard, candidates] of byStandard) {
    const best = Math.min(...candidates.map(({ rank }) => rank));
    const matches = candidates.filter(({ rank }) => rank === best);
    if (matches.length > 1) {
      throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `multiple source columns map to standard column ${standard}`);
    }
    mapping.set(standard, matches[0].header.index);
  }
  return mapping;
}

export async function readFileRows(file, spec, standardColumns = null) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(file.path);
  } catch {
    throw new InputError('INPUT_ERROR', `cannot read XLSX input for ${file.id}`);
  }

  const sheet = workbook.getWorksheet(spec.sheet.name);
  if (!sheet) {
    throw new InputError('SHEET_NOT_FOUND', `sheet ${spec.sheet.name} not found for ${file.id}; available: ${workbook.worksheets.map(({ name }) => name).join(', ')}`);
  }

  const headers = readHeaders(sheet, spec.sheet.headerRow);
  const required = requiredColumns(spec);
  const baselineColumns = standardColumns ?? (spec.compareColumns === '*'
    ? headers.map(({ text }) => text)
    : required);
  const columns = spec.compareColumns === '*' ? [...new Set([...baselineColumns, ...required])] : required;
  const mapping = mapColumns(headers, columns, spec.columnAliases ?? {});
  for (const column of columns) {
    if (!mapping.has(column)) {
      throw new InputError('COLUMN_MISSING', `column ${column} is missing from ${file.id}`);
    }
  }

  const rows = [];
  let invalidRows = 0;
  let totalRowsScanned = 0;
  for (let rowNumber = spec.sheet.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    totalRowsScanned += 1;
    const row = sheet.getRow(rowNumber);
    const values = Object.fromEntries(columns.filter((column) => mapping.has(column)).map((column) => [
      column,
      normalizeValue(row.getCell(mapping.get(column)).value, ruleForColumn(spec, column))
    ]));
    if (!spec.filters.every((filter) => matchesFilter(values[filter.column], filter, ruleForColumn(spec, filter.column)))) continue;
    if (spec.mode.keyColumns.some((column) => values[column]?.[0] === 'blank')) {
      invalidRows += 1;
      continue;
    }
    rows.push({ fileId: file.id, sheetName: sheet.name, rowNumber, values });
  }

  return { columns, rows, invalidRows, totalRowsScanned };
}
