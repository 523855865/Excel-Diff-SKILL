import { openStreamingWorkbook } from './exceljs-stream-compat.js';
import { matchesFilter, normalizeValue } from './normalize.js';
import { PartitionError } from './partitions.js';

export class InputError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'InputError';
    this.code = code;
  }
}

function requiredColumns(spec) {
  const columns = spec.mode.type === 'row' || spec.mode.type === 'multiset' ? [] : [...spec.mode.keyColumns];
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

function readHeaders(row) {
  const headers = [];
  const seen = new Map();
  for (let index = 1; index <= (row?.cellCount ?? 0); index += 1) {
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

function resolveColumns(row, file, spec, standardColumns) {
  const headers = readHeaders(row);
  const required = requiredColumns(spec);
  const aliases = spec.columnAliases ?? {};
  let columns;
  let mapping;
  if (standardColumns === null && spec.compareColumns === '*') {
    mapping = mapColumns(headers, [...new Set([...required, ...Object.keys(aliases)])], aliases);
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
  return { columns, mapping };
}

export async function scanFileRows(file, spec, standardColumns = null, onRow = async () => {}, onScan = async () => {}, options = {}) {
  const availableSheets = [];
  let columns;
  let mapping;
  let foundSheet = false;
  let invalidRows = 0;
  let totalRowsScanned = 0;
  let matchedRows = 0;
  let callbackFailure;
  let compatibility;

  const processRow = async (row, rowNumber, sheetName) => {
    const values = Object.fromEntries(columns.filter((column) => mapping.has(column)).map((column) => {
      const cell = row.getCell(mapping.get(column));
      const value = compatibility.cellValue(cell);
      try {
        return [column, normalizeValue(value, ruleForColumn(spec, column))];
      } catch (error) {
        if (error instanceof TypeError && error.message === 'unsupported cell value') {
          throw new InputError('CELL_VALUE_UNSUPPORTED', `unsupported cell value in ${file.id} row ${rowNumber} column ${column}`);
        }
        throw error;
      }
    }));
    let matches = true;
    for (const filter of spec.filters) {
      try {
        if (!matchesFilter(values[filter.column], filter, ruleForColumn(spec, filter.column))) matches = false;
      } catch (error) {
        if (error instanceof TypeError) throw new InputError('FILTER_INVALID', `invalid filter for column ${filter.column}`);
        throw error;
      }
      if (!matches) break;
    }
    if (!matches) return;
    if (spec.mode.type !== 'row' && spec.mode.type !== 'multiset'
      && spec.mode.keyColumns.some((column) => values[column]?.[0] === 'blank')) {
      invalidRows += 1;
      return;
    }
    matchedRows += 1;
    try {
      await onRow({ fileId: file.id, sheetName, rowNumber, values }, columns);
    } catch (error) {
      callbackFailure = error;
      throw error;
    }
  };

  const reportScan = async (rowNumber, cellCount) => {
    try {
      await onScan({ fileId: file.id, rowNumber, cellCount });
    } catch (error) {
      callbackFailure = error;
      throw error;
    }
  };

  const consume = async (workbook) => {
    for await (const sheet of workbook) {
      availableSheets.push(sheet.name);
      if (sheet.name !== spec.sheet.name) continue;
      foundSheet = true;
      compatibility.prepareWorksheet(sheet);
      let previousRowNumber = spec.sheet.headerRow;
      for await (const row of sheet) {
        try {
          if (row.number < spec.sheet.headerRow) continue;
          if (row.number === spec.sheet.headerRow) {
            ({ columns, mapping } = resolveColumns(row, file, spec, standardColumns));
            continue;
          }
          if (columns === undefined) ({ columns, mapping } = resolveColumns(null, file, spec, standardColumns));
          for (let rowNumber = previousRowNumber + 1; rowNumber < row.number; rowNumber += 1) {
            totalRowsScanned += 1;
            await reportScan(rowNumber, 0);
            await processRow({ getCell: () => ({ value: null }) }, rowNumber, sheet.name);
          }
          totalRowsScanned += 1;
          await reportScan(row.number, row.cellCount);
          await processRow(row, row.number, sheet.name);
          previousRowNumber = row.number;
        } finally {
          compatibility.releaseRow(row.number);
        }
      }
      if (columns === undefined) ({ columns, mapping } = resolveColumns(null, file, spec, standardColumns));
    }
  };

  let failure;
  try {
    compatibility = await openStreamingWorkbook(file.path, spec.sheet.name, {
      maxTempBytes: spec.resources?.maxTempBytes,
      tempBudget: options.tempBudget,
      check: options.check,
      openFile: options.openFile,
      remove: options.remove
    });
    await consume(compatibility.workbook);
  } catch (error) {
    if (error instanceof InputError || error instanceof PartitionError || error === callbackFailure
      || error?.code === 'RUNTIME_LIMIT_EXCEEDED' || error?.code === 'ENOSPC' || error?.code === 'EDQUOT') {
      failure = error;
    } else {
      failure = new InputError('INPUT_ERROR', `cannot read XLSX input for ${file.id}`);
      if (error?.cleanupError) failure.cleanupError = error.cleanupError;
    }
    throw failure;
  } finally {
    if (compatibility) {
      let cleanupError;
      for (let attempt = 0; attempt < 2; attempt += 1) {
        try {
          await compatibility.close();
          cleanupError = undefined;
          break;
        } catch (error) {
          cleanupError = error;
        }
      }
      if (cleanupError) {
        if (!failure) throw cleanupError;
        if (failure && (typeof failure === 'object' || typeof failure === 'function')) {
          try { failure.cleanupError = cleanupError; } catch {}
        }
      }
    }
  }

  if (!foundSheet) {
    throw new InputError('SHEET_NOT_FOUND', `sheet ${spec.sheet.name} not found for ${file.id}; available: ${availableSheets.join(', ')}`);
  }
  return { columns, invalidRows, totalRowsScanned, matchedRows };
}

export async function readFileRows(file, spec, standardColumns = null) {
  const rows = [];
  const result = await scanFileRows(file, spec, standardColumns, async (row) => rows.push(row));
  return {
    columns: result.columns,
    rows,
    invalidRows: result.invalidRows,
    totalRowsScanned: result.totalRowsScanned
  };
}
