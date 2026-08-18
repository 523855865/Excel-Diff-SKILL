import ExcelJS from 'exceljs';
import colCache from 'exceljs/lib/utils/col-cache.js';
import parseSax from 'exceljs/lib/utils/parse-sax.js';
import sharedFormulaUtils from 'exceljs/lib/utils/shared-formula.js';
import excelUtils from 'exceljs/lib/utils/utils.js';
import relationshipTypes from 'exceljs/lib/xlsx/rel-type.js';
import { createRequire } from 'node:module';
import { StringDecoder } from 'node:string_decoder';

import { matchesFilter, normalizeValue } from './normalize.js';

const require = createRequire(import.meta.resolve('exceljs'));
const { slideFormula } = sharedFormulaUtils;
const unzipper = require('unzipper');

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

function xmlAttribute(tag, name) {
  return new RegExp(`(?:^|\\s)${name}=(["'])(.*?)\\1`).exec(tag)?.[2];
}

function addressInRange(address, reference) {
  const cell = colCache.decodeAddress(address);
  const range = colCache.decode(reference);
  return cell.row >= range.top && cell.row <= range.bottom && cell.col >= range.left && cell.col <= range.right;
}

async function* expandSharedFormulas(iterator) {
  const decoder = new StringDecoder('utf8');
  const masters = new Map();
  let buffer = '';
  let cellAddress;

  const drain = (final = false) => {
    let output = '';
    while (buffer !== '') {
      const tagStart = buffer.indexOf('<');
      if (tagStart < 0) {
        output += buffer;
        buffer = '';
        break;
      }
      if (tagStart > 0) {
        output += buffer.slice(0, tagStart);
        buffer = buffer.slice(tagStart);
        continue;
      }

      const tagEnd = buffer.indexOf('>');
      if (tagEnd < 0) {
        if (final) {
          output += buffer;
          buffer = '';
        }
        break;
      }
      const tag = buffer.slice(0, tagEnd + 1);
      if (/^<c(?:\s|>)/.test(tag)) cellAddress = xmlAttribute(tag, 'r');

      if (/^<f(?:\s|>)/.test(tag) && xmlAttribute(tag, 't') === 'shared') {
        const sharedId = xmlAttribute(tag, 'si');
        const selfClosing = /\/>$/.test(tag);
        const close = selfClosing ? tagEnd + 1 : buffer.indexOf('</f>', tagEnd + 1);
        if (close < 0) break;
        const formula = selfClosing ? '' : buffer.slice(tagEnd + 1, close);
        const consumed = selfClosing ? tagEnd + 1 : close + 4;
        if (formula !== '') {
          masters.set(sharedId, { address: cellAddress, formula, ref: xmlAttribute(tag, 'ref') });
          output += buffer.slice(0, consumed);
        } else {
          const master = masters.get(sharedId);
          if (!master || !master.ref || !addressInRange(cellAddress, master.ref)) throw new Error('invalid shared formula');
          output += `<f>${slideFormula(master.formula, master.address, cellAddress)}</f>`;
        }
        buffer = buffer.slice(consumed);
        continue;
      }

      output += tag;
      buffer = buffer.slice(tagEnd + 1);
    }
    return output;
  };

  for await (const chunk of iterator) {
    buffer += decoder.write(chunk);
    const output = drain();
    if (output !== '') yield output;
  }
  buffer += decoder.end();
  const output = drain(true);
  if (output !== '') yield output;
}

async function readHyperlinks(entries, workbook, sheetName) {
  const sheet = workbook.model.sheets.find(({ name }) => name === sheetName);
  const relationship = sheet && workbook.workbookRels.find(({ Id }) => Id === sheet.rId);
  if (!relationship) return new Map();

  const sheetPath = relationship.Target.startsWith('/') ? relationship.Target.slice(1) : `xl/${relationship.Target}`;
  const separator = sheetPath.lastIndexOf('/');
  const relationshipPath = `${sheetPath.slice(0, separator)}/_rels/${sheetPath.slice(separator + 1)}.rels`;
  if (!entries.has(sheetPath) || !entries.has(relationshipPath)) return new Map();

  const targets = new Map();
  for await (const events of parseSax(entries.get(relationshipPath).stream())) {
    for (const { eventType, value } of events) {
      if (eventType === 'opentag' && value.name === 'Relationship'
        && value.attributes.Type === relationshipTypes.Hyperlink) {
        targets.set(value.attributes.Id, value.attributes.Target);
      }
    }
  }
  if (targets.size === 0) return new Map();

  // ponytail: hyperlink-bearing sheets are decompressed twice because ExcelJS 4.4 emits refs after rows; replace when upstream provides row-time targets.
  const hyperlinks = new Map();
  for await (const events of parseSax(entries.get(sheetPath).stream())) {
    for (const { eventType, value } of events) {
      if (eventType !== 'opentag' || value.name !== 'hyperlink') continue;
      const target = targets.get(value.attributes['r:id']);
      if (target) hyperlinks.set(value.attributes.ref, { hyperlink: target, tooltip: value.attributes.tooltip });
    }
  }
  return hyperlinks;
}

export async function scanFileRows(file, spec, standardColumns = null, onRow = async () => {}) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookReader(file.path, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'cache'
  });

  const availableSheets = [];
  let columns;
  let mapping;
  let foundSheet = false;
  let invalidRows = 0;
  let totalRowsScanned = 0;
  let matchedRows = 0;
  let callbackFailure;
  let hyperlinks = new Map();

  const processRow = async (row, rowNumber, sheetName) => {
    const values = Object.fromEntries(columns.filter((column) => mapping.has(column)).map((column) => {
      const cell = row.getCell(mapping.get(column));
      const formula = cell.formula;
      const hyperlink = hyperlinks.get(cell.address);
      let value = hyperlink ? { text: cell.text, ...hyperlink }
        : cell.value?.sharedFormula ? { ...cell.value, formula } : cell.value;
      if (value?.formula && typeof value.result === 'number' && excelUtils.isDateFmt(cell.numFmt)) {
        value = { ...value, result: excelUtils.excelToDate(value.result, workbook.properties.model?.date1904) };
      }
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
      await onRow({ fileId: file.id, sheetName, rowNumber, values });
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
      sheet.iterator = expandSharedFormulas(sheet.iterator);
      let previousRowNumber = spec.sheet.headerRow;
      for await (const row of sheet) {
        if (row.number < spec.sheet.headerRow) continue;
        if (row.number === spec.sheet.headerRow) {
          ({ columns, mapping } = resolveColumns(row, file, spec, standardColumns));
          continue;
        }
        if (columns === undefined) ({ columns, mapping } = resolveColumns(null, file, spec, standardColumns));
        for (let rowNumber = previousRowNumber + 1; rowNumber < row.number; rowNumber += 1) {
          totalRowsScanned += 1;
          await processRow({ getCell: () => ({ value: null }) }, rowNumber, sheet.name);
        }
        totalRowsScanned += 1;
        await processRow(row, row.number, sheet.name);
        previousRowNumber = row.number;
      }
      if (columns === undefined) ({ columns, mapping } = resolveColumns(null, file, spec, standardColumns));
    }
  };

  try {
    // ExcelJS can encounter worksheet ZIP entries before workbook metadata, so preload metadata without reading rows.
    const archive = await unzipper.Open.file(file.path);
    const entries = new Map(archive.files.map((entry) => [entry.path, entry]));
    await workbook._parseRels(entries.get('xl/_rels/workbook.xml.rels').stream());
    await workbook._parseWorkbook(entries.get('xl/workbook.xml').stream());
    if (entries.has('xl/sharedStrings.xml')) {
      for await (const unused of workbook._parseSharedStrings(entries.get('xl/sharedStrings.xml').stream())) void unused;
    }
    if (entries.has('xl/styles.xml')) await workbook._parseStyles(entries.get('xl/styles.xml').stream());
    hyperlinks = await readHyperlinks(entries, workbook, spec.sheet.name);
    await consume(workbook);
  } catch (error) {
    if (error instanceof InputError || error === callbackFailure) throw error;
    throw new InputError('INPUT_ERROR', `cannot read XLSX input for ${file.id}`);
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
