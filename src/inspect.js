import { constants } from 'node:fs';
import { access, realpath, stat } from 'node:fs/promises';
import { extname, resolve } from 'node:path';

import { openStreamingWorkbook } from './exceljs-stream-compat.js';
import { encodeKey, normalizeValue } from './normalize.js';
import { InputError } from './read-xlsx.js';

function invalid(message) {
  return new InputError('INPUT_ERROR', message);
}

async function resolveFiles(paths) {
  if (!Array.isArray(paths) || paths.length < 2) throw invalid('at least two XLSX inputs are required');
  const files = [];
  const seen = new Set();
  for (const input of paths) {
    if (typeof input !== 'string' || extname(input).toLowerCase() !== '.xlsx') {
      throw invalid(`input file must use .xlsx: ${String(input)}`);
    }
    let file;
    try {
      file = await realpath(resolve(input));
      await access(file, constants.R_OK);
      if (!(await stat(file)).isFile()) throw new Error('not a file');
    } catch {
      throw invalid(`cannot read XLSX input: ${resolve(input)}`);
    }
    if (seen.has(file)) throw invalid(`duplicate input path: ${file}`);
    seen.add(file);
    files.push(file);
  }
  return files;
}

function readHeaders(row) {
  const headers = [];
  for (let index = 1; index <= (row?.cellCount ?? 0); index += 1) {
    const raw = row.getCell(index).text;
    if (raw === '') continue;
    headers.push({
      index,
      raw,
      normalized: raw.normalize('NFKC'),
      types: {},
      blankCount: 0,
      sampleBlankCount: 0,
      sampleValues: []
    });
  }
  return headers;
}

function duplicateHeaderRisks(headers) {
  const groups = new Map();
  for (const header of headers) {
    const group = groups.get(header.normalized) ?? [];
    group.push({ index: header.index, raw: header.raw });
    groups.set(header.normalized, group);
  }
  return [...groups]
    .filter(([, columns]) => columns.length > 1)
    .map(([normalized, columns]) => ({ code: 'HEADER_DUPLICATED', normalized, columns }));
}

function canonicalCell(compatibility, row, header, file) {
  try {
    return normalizeValue(row ? compatibility.cellValue(row.getCell(header.index)) : null, {
      emptyEqualsNull: false,
      caseSensitive: true,
      formulaMode: 'formula-and-cached-result'
    });
  } catch (error) {
    if (error instanceof TypeError) {
      throw new InputError('CELL_VALUE_UNSUPPORTED', `unsupported cell value in ${file} column ${header.normalized}`);
    }
    throw error;
  }
}

async function inspectFile(file, options) {
  let compatibility;
  let failure;
  try {
    compatibility = await openStreamingWorkbook(file, options.sheet);
    const sheets = (compatibility.workbook.model.sheets ?? []).map(({ name, state }) => ({
      name,
      visibility: state ?? 'visible'
    }));
    const target = sheets.find(({ name }) => name === options.sheet);
    if (!target) {
      throw new InputError('SHEET_NOT_FOUND', `sheet ${options.sheet} not found; available: ${sheets.map(({ name }) => name).join(', ')}`);
    }

    let headers = [];
    let sampledRows = 0;
    let typedRows = 0;
    let formulaCells = 0;
    let foundSheet = false;
    let previousRowNumber = options.headerRow;

    const consumeDataRow = (row) => {
      const sampled = sampledRows < options.sampleRows;
      const countTypes = options.fullTypes || sampled;
      if (!sampled && !countTypes) return;
      for (const header of headers) {
        const value = canonicalCell(compatibility, row, header, file);
        if (countTypes) {
          header.types[value[0]] = (header.types[value[0]] ?? 0) + 1;
          if (value[0] === 'blank') header.blankCount += 1;
        }
        if (sampled) {
          header.sampleValues.push(encodeKey([value]));
          if (value[0] === 'blank') header.sampleBlankCount += 1;
        }
      }
      if (countTypes) typedRows += 1;
      if (sampled) sampledRows += 1;
    };

    for await (const worksheet of compatibility.workbook) {
      if (worksheet.name !== options.sheet) continue;
      foundSheet = true;
      compatibility.prepareWorksheet(worksheet);
      for await (const row of worksheet) {
        try {
          for (let index = 1; index <= row.cellCount; index += 1) {
            if (row.getCell(index).formula !== undefined) formulaCells += 1;
          }
          if (row.number < options.headerRow) continue;
          if (row.number === options.headerRow) {
            headers = readHeaders(row);
            continue;
          }
          for (let rowNumber = previousRowNumber + 1; rowNumber < row.number; rowNumber += 1) consumeDataRow(null);
          consumeDataRow(row);
          previousRowNumber = row.number;
        } finally {
          compatibility.releaseRow(row.number);
        }
      }
    }
    if (!foundSheet) {
      throw new InputError('SHEET_NOT_FOUND', `sheet ${options.sheet} not found; available: ${sheets.map(({ name }) => name).join(', ')}`);
    }

    const risks = duplicateHeaderRisks(headers);
    if (formulaCells > 0) risks.push({ code: 'FORMULA_CELLS', count: formulaCells });
    if (compatibility.mergeCount > 0) risks.push({ code: 'MERGED_CELLS', count: compatibility.mergeCount });
    const duplicated = new Set(risks
      .filter(({ code }) => code === 'HEADER_DUPLICATED')
      .map(({ normalized }) => normalized));

    // ponytail: only rank single-column candidates; add bounded composite search when real workbooks prove it is needed.
    const keyCandidates = headers.flatMap((header) => {
      const duplicates = header.sampleValues.length - new Set(header.sampleValues).size;
      if (sampledRows === 0 || header.sampleBlankCount > 0 || duplicates > 0 || duplicated.has(header.normalized)) return [];
      return [{ columns: [header.normalized], duplicateRate: duplicates / sampledRows, sampledRows }];
    });

    return {
      file,
      sheets,
      sheet: { ...target, headerRow: options.headerRow },
      sampledRows,
      headers: headers.map(({ index, raw, normalized, types, blankCount }) => ({
        index,
        raw,
        normalized,
        types,
        emptyRatio: typedRows === 0 ? 0 : blankCount / typedRows
      })),
      keyCandidates,
      risks
    };
  } catch (error) {
    failure = error instanceof InputError ? error : invalid(`cannot read XLSX input: ${file}`);
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
        try { failure.cleanupError = cleanupError; } catch {}
      }
    }
  }
}

export async function inspectFiles(paths, { sheet, headerRow = 1, sampleRows = 10_000, fullTypes = false } = {}) {
  if (typeof sheet !== 'string' || sheet === '') throw invalid('sheet is required');
  if (!Number.isInteger(headerRow) || headerRow < 1) throw invalid('headerRow must be a positive integer');
  if (!Number.isInteger(sampleRows) || sampleRows < 1) throw invalid('sampleRows must be a positive integer');
  if (typeof fullTypes !== 'boolean') throw invalid('fullTypes must be boolean');
  const pathsToInspect = await resolveFiles(paths);
  const files = [];
  for (const file of pathsToInspect) files.push(await inspectFile(file, { sheet, headerRow, sampleRows, fullTypes }));
  return {
    status: 'INSPECTED',
    sampleRows,
    fullTypes,
    files
  };
}
