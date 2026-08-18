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

function consumeTier(headers, standards, matches, mapping) {
  const byHeader = new Map(headers.map((header) => [header.index, standards.filter((standard) => matches(header, standard))]));
  const byStandard = new Map(standards.map((standard) => [standard, headers.filter((header) => matches(header, standard))]));
  for (const [index, candidates] of byHeader) {
    if (candidates.length > 1) {
      throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `source column ${headers.find((header) => header.index === index).text} maps to multiple standard columns`);
    }
  }
  for (const [standard, candidates] of byStandard) {
    if (candidates.length > 1) {
      throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `multiple source columns map to standard column ${standard}`);
    }
  }
  for (const header of headers) {
    const [standard] = byHeader.get(header.index);
    if (standard === undefined) continue;
    mapping.set(standard, header.index);
  }
  return {
    headers: headers.filter((header) => byHeader.get(header.index).length === 0),
    standards: standards.filter((standard) => byStandard.get(standard).length === 0)
  };
}

function aliasGraph(headers, standards, aliases) {
  const byHeader = new Map(headers.map((header) => [header.index, []]));
  const byStandard = new Map(standards.map((standard) => [standard, []]));
  for (const header of headers) {
    for (const standard of standards) {
      if (!(aliases[standard] ?? []).some((alias) => header.text === alias || header.normalized === alias.normalize('NFKC'))) continue;
      byHeader.get(header.index).push(standard);
      byStandard.get(standard).push(header);
    }
  }
  return { byHeader, byStandard };
}

function resolveAliases(headers, standards, aliases, mapping) {
  let remainingHeaders = headers;
  let remainingStandards = standards;
  while (true) {
    const graph = aliasGraph(remainingHeaders, remainingStandards, aliases);
    const forced = remainingHeaders.filter((header) => {
      const candidates = graph.byHeader.get(header.index);
      return candidates.length === 1 && graph.byStandard.get(candidates[0]).length === 1;
    });
    if (forced.length === 0) break;
    for (const header of forced) mapping.set(graph.byHeader.get(header.index)[0], header.index);
    const usedHeaders = new Set(forced.map(({ index }) => index));
    const usedStandards = new Set(forced.map((header) => graph.byHeader.get(header.index)[0]));
    remainingHeaders = remainingHeaders.filter(({ index }) => !usedHeaders.has(index));
    remainingStandards = remainingStandards.filter((standard) => !usedStandards.has(standard));
  }

  const graph = aliasGraph(remainingHeaders, remainingStandards, aliases);
  const activeHeaders = remainingHeaders.filter((header) => graph.byHeader.get(header.index).length > 0);
  const activeStandards = remainingStandards.filter((standard) => graph.byStandard.get(standard).length > 0);
  if (activeHeaders.length === 0) return;
  if (activeHeaders.length !== activeStandards.length) {
    if (activeHeaders.length < activeStandards.length) {
      throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `source column ${activeHeaders[0].text} maps to multiple standard columns`);
    }
    throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `multiple source columns map to standard column ${activeStandards[0]}`);
  }

  const orderedStandards = [...activeStandards].sort((left, right) => graph.byStandard.get(left).length - graph.byStandard.get(right).length);
  const solutions = [];
  // ponytail: stop after two solutions; use a bipartite matcher only if wide alias graphs become a measured bottleneck.
  const search = (position, usedHeaders, solution) => {
    if (solutions.length > 1) return;
    if (position === orderedStandards.length) {
      solutions.push(solution);
      return;
    }
    const standard = orderedStandards[position];
    for (const header of graph.byStandard.get(standard)) {
      if (usedHeaders.has(header.index)) continue;
      search(position + 1, new Set([...usedHeaders, header.index]), new Map([...solution, [standard, header.index]]));
    }
  };
  search(0, new Set(), new Map());
  if (solutions.length !== 1) {
    const header = activeHeaders.find(({ index }) => graph.byHeader.get(index).length > 1);
    if (header) throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `source column ${header.text} maps to multiple standard columns`);
    throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `multiple source columns map to standard column ${activeStandards[0]}`);
  }
  for (const [standard, index] of solutions[0]) mapping.set(standard, index);
}

function mapColumns(headers, standards, aliases) {
  const mapping = new Map();
  let remaining = consumeTier(headers, standards, (header, standard) => header.text === standard, mapping);
  remaining = consumeTier(remaining.headers, remaining.standards, (header, standard) => header.normalized === standard.normalize('NFKC'), mapping);
  resolveAliases(remaining.headers, remaining.standards, aliases, mapping);
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
  const aliases = spec.columnAliases ?? {};
  let columns;
  let mapping;
  if (standardColumns === null && spec.compareColumns === '*') {
    mapping = mapColumns(headers, required, aliases);
    for (const column of required) {
      if (!mapping.has(column)) throw new InputError('COLUMN_MISSING', `column ${column} is missing from ${file.id}`);
    }
    const canonicalByHeader = new Map([...mapping].map(([column, index]) => [index, column]));
    columns = headers.map((header) => canonicalByHeader.get(header.index) ?? header.text);
    for (const header of headers) {
      if (!canonicalByHeader.has(header.index)) mapping.set(header.text, header.index);
    }
  } else {
    const baselineColumns = standardColumns ?? required;
    columns = spec.compareColumns === '*' ? [...new Set([...baselineColumns, ...required])] : required;
    mapping = mapColumns(headers, columns, aliases);
  }
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
