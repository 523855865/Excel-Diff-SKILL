import ExcelJS from 'exceljs';
import colCache from 'exceljs/lib/utils/col-cache.js';
import parseSax from 'exceljs/lib/utils/parse-sax.js';
import sharedFormulaUtils from 'exceljs/lib/utils/shared-formula.js';
import excelUtils from 'exceljs/lib/utils/utils.js';
import relationshipTypes from 'exceljs/lib/xlsx/rel-type.js';
import { readSync } from 'node:fs';
import { mkdtemp, open, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StringDecoder } from 'node:string_decoder';
import unzipper from 'unzipper';

import { PartitionError } from './partitions.js';

const { slideFormula } = sharedFormulaUtils;
const sharedStringIndexBytes = 16;

async function writeAll(file, buffer, position) {
  let written = 0;
  while (written < buffer.length) {
    const result = await file.write(buffer, written, buffer.length - written, position + written);
    if (result.bytesWritten === 0) throw new Error('shared string spool write returned zero bytes');
    written += result.bytesWritten;
  }
}

function readExact(fd, buffer, position) {
  let read = 0;
  while (read < buffer.length) {
    const bytes = readSync(fd, buffer, read, buffer.length - read, position + read);
    if (bytes === 0) throw new Error('invalid shared string spool');
    read += bytes;
  }
}

async function retryCleanup(action) {
  let failure;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await action();
    } catch (error) {
      failure = error;
    }
  }
  throw failure;
}

async function createSharedStringTable(options) {
  const openFile = options.openFile ?? open;
  const remove = options.remove ?? rm;
  const directory = await mkdtemp(join(tmpdir(), 'excel-diff-shared-'));
  let data;
  let index;
  try {
    data = await openFile(join(directory, 'data'), 'w+');
    index = await openFile(join(directory, 'index'), 'w+');
  } catch (error) {
    let cleanupError;
    for (const handle of [data, index]) {
      if (!handle) continue;
      try {
        await retryCleanup(() => handle.close());
      } catch (closeError) {
        cleanupError ??= closeError;
      }
    }
    try {
      await retryCleanup(() => remove(directory, { recursive: true, force: true }));
    } catch (removeError) {
      cleanupError ??= removeError;
    }
    if (cleanupError && error && (typeof error === 'object' || typeof error === 'function')) {
      try { error.cleanupError = cleanupError; } catch {}
    }
    throw error;
  }
  let count = 0;
  let offset = 0;
  let reserved = 0;
  let dataClosed = false;
  let indexClosed = false;
  let removed = false;
  let released = false;
  const reserve = options.tempBudget
    ? (bytes) => options.tempBudget.reserveExternal(bytes)
    : (bytes) => {
      if (reserved + bytes > (options.maxTempBytes ?? Infinity)) {
        throw new PartitionError('TEMP_LIMIT_EXCEEDED', 'temporary data exceeds resources.maxTempBytes');
      }
    };
  const values = new Proxy(Object.create(null), {
    get(_target, property) {
      if (property === 'length') return count;
      if (typeof property !== 'string' || !/^(?:0|[1-9]\d*)$/.test(property)) return undefined;
      const item = Number(property);
      if (item >= count) return undefined;
      const location = Buffer.allocUnsafe(sharedStringIndexBytes);
      readExact(index.fd, location, item * sharedStringIndexBytes);
      const start = Number(location.readBigUInt64LE(0));
      const length = location.readUInt32LE(8);
      const encoded = Buffer.allocUnsafe(length);
      readExact(data.fd, encoded, start);
      return JSON.parse(encoded.toString('utf8'));
    }
  });
  return {
    directory,
    values,
    get bytes() { return reserved; },
    async append(item) {
      const encoded = Buffer.from(JSON.stringify(item), 'utf8');
      const bytes = encoded.length + sharedStringIndexBytes;
      reserve(bytes);
      reserved += bytes;
      const location = Buffer.alloc(sharedStringIndexBytes);
      location.writeBigUInt64LE(BigInt(offset), 0);
      location.writeUInt32LE(encoded.length, 8);
      await writeAll(data, encoded, offset);
      await writeAll(index, location, count * sharedStringIndexBytes);
      offset += encoded.length;
      count += 1;
    },
    async close() {
      let firstError;
      if (!dataClosed) {
        try {
          await data.close();
          dataClosed = true;
        } catch (error) {
          firstError ??= error;
        }
      }
      if (!indexClosed) {
        try {
          await index.close();
          indexClosed = true;
        } catch (error) {
          firstError ??= error;
        }
      }
      if (dataClosed && indexClosed && !removed) {
        try {
          await remove(directory, { recursive: true, force: true });
          removed = true;
        } catch (error) {
          firstError ??= error;
        }
      }
      if (dataClosed && indexClosed && removed && !released) {
        try {
          if (reserved > 0) options.tempBudget?.releaseExternal(reserved);
          reserved = 0;
          released = true;
        } catch (error) {
          firstError ??= error;
        }
      }
      if (firstError) throw firstError;
    }
  };
}

function xmlAttribute(tag, name) {
  return new RegExp(`(?:^|\\s)${name}=(["'])(.*?)\\1`).exec(tag)?.[2];
}

function addressInRange(address, reference) {
  const cell = colCache.decodeAddress(address);
  const range = colCache.decode(reference);
  return cell.row >= range.top && cell.row <= range.bottom && cell.col >= range.left && cell.col <= range.right;
}

async function* transformWorksheet(iterator, formulaResults) {
  const decoder = new StringDecoder('utf8');
  const masters = new Map();
  const expiringMasters = new Map();
  let buffer = '';
  let cellAddress;
  let cellType;
  let cellHasFormula = false;
  let rowNumber;

  const drain = function* (final = false) {
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
      if (/^<row(?:\s|>)/.test(tag)) rowNumber = Number(xmlAttribute(tag, 'r'));
      if (/^<c(?:\s|>)/.test(tag)) {
        cellAddress = xmlAttribute(tag, 'r');
        cellType = xmlAttribute(tag, 't');
        cellHasFormula = false;
      }

      if (/^<f(?:\s|>)/.test(tag)) cellHasFormula = true;

      if (/^<f(?:\s|>)/.test(tag) && xmlAttribute(tag, 't') === 'shared') {
        const sharedId = xmlAttribute(tag, 'si');
        const selfClosing = /\/>$/.test(tag);
        const close = selfClosing ? tagEnd + 1 : buffer.indexOf('</f>', tagEnd + 1);
        if (close < 0) break;
        const formula = selfClosing ? '' : buffer.slice(tagEnd + 1, close);
        const consumed = selfClosing ? tagEnd + 1 : close + 4;
        if (formula !== '') {
          const ref = xmlAttribute(tag, 'ref');
          const master = { sharedId, address: cellAddress, formula, ref };
          masters.set(sharedId, master);
          if (ref) {
            const bottom = colCache.decode(ref).bottom;
            const expiring = expiringMasters.get(bottom) ?? new Set();
            expiring.add(master);
            expiringMasters.set(bottom, expiring);
          }
          output += buffer.slice(0, consumed);
        } else {
          const master = masters.get(sharedId);
          if (!master || !master.ref || !addressInRange(cellAddress, master.ref)) throw new Error('invalid shared formula');
          output += `<f>${slideFormula(master.formula, master.address, cellAddress)}</f>`;
        }
        buffer = buffer.slice(consumed);
        continue;
      }

      if (/^<v(?:\s|>)/.test(tag) && cellHasFormula && (cellType === 'b' || cellType === 'e')) {
        const close = buffer.indexOf('</v>', tagEnd + 1);
        if (close < 0) break;
        const rowNumber = colCache.decodeAddress(cellAddress).row;
        const rowResults = formulaResults.get(rowNumber) ?? new Map();
        rowResults.set(cellAddress, { type: cellType, value: buffer.slice(tagEnd + 1, close) });
        formulaResults.set(rowNumber, rowResults);
        output += buffer.slice(0, close + 4);
        buffer = buffer.slice(close + 4);
        continue;
      }

      if (tag === '</row>') {
        for (const master of expiringMasters.get(rowNumber) ?? []) {
          if (masters.get(master.sharedId) === master) masters.delete(master.sharedId);
        }
        expiringMasters.delete(rowNumber);
        output += tag;
        buffer = buffer.slice(tagEnd + 1);
        yield output;
        output = '';
        continue;
      }

      output += tag;
      buffer = buffer.slice(tagEnd + 1);
    }
    if (output !== '') yield output;
  };

  for await (const chunk of iterator) {
    buffer += decoder.write(chunk);
    yield* drain();
  }
  buffer += decoder.end();
  yield* drain(true);
}

async function readHyperlinks(entries, workbook, sheetName, check) {
  check();
  const sheet = workbook.model.sheets.find(({ name }) => name === sheetName);
  const relationship = sheet && workbook.workbookRels.find(({ Id }) => Id === sheet.rId);
  if (!relationship) return new Map();

  const sheetPath = relationship.Target.startsWith('/') ? relationship.Target.slice(1) : `xl/${relationship.Target}`;
  const separator = sheetPath.lastIndexOf('/');
  const relationshipPath = `${sheetPath.slice(0, separator)}/_rels/${sheetPath.slice(separator + 1)}.rels`;
  if (!entries.has(sheetPath) || !entries.has(relationshipPath)) return new Map();

  const targets = new Map();
  for await (const events of parseSax(entries.get(relationshipPath).stream())) {
    check();
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
    check();
    for (const { eventType, value } of events) {
      if (eventType !== 'opentag' || value.name !== 'hyperlink') continue;
      const target = targets.get(value.attributes['r:id']);
      if (target) hyperlinks.set(value.attributes.ref, { hyperlink: target, tooltip: value.attributes.tooltip });
    }
  }
  return hyperlinks;
}

export async function openStreamingWorkbook(filePath, sheetName, options = {}) {
  const check = options.check ?? (() => {});
  const workbook = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    worksheets: 'emit',
    sharedStrings: 'emit',
    hyperlinks: 'ignore',
    styles: 'cache'
  });
  let sharedStrings;
  try {
    check();
    const archive = await unzipper.Open.file(filePath);
    check();
    const entries = new Map(archive.files.map((entry) => [entry.path, entry]));
    await workbook._parseRels(entries.get('xl/_rels/workbook.xml.rels').stream());
    check();
    await workbook._parseWorkbook(entries.get('xl/workbook.xml').stream());
    check();
    if (entries.has('xl/sharedStrings.xml')) {
      sharedStrings = await createSharedStringTable(options);
      for await (const { text } of workbook._parseSharedStrings(entries.get('xl/sharedStrings.xml').stream())) {
        check();
        await sharedStrings.append(text);
      }
      workbook.sharedStrings = sharedStrings.values;
    } else workbook.sharedStrings = [];
    check();
    if (entries.has('xl/styles.xml')) {
      await workbook._parseStyles(entries.get('xl/styles.xml').stream());
      check();
    }

    const hyperlinks = await readHyperlinks(entries, workbook, sheetName, check);
    check();
    const formulaResults = new Map();
    return {
      workbook,
      tempDirectory: sharedStrings?.directory,
      get tempBytes() { return sharedStrings?.bytes ?? 0; },
      close: async () => sharedStrings?.close(),
      prepareWorksheet(sheet) {
        sheet.iterator = transformWorksheet(sheet.iterator, formulaResults);
      },
      cellValue(cell) {
        const formula = cell.formula;
        const hyperlink = hyperlinks.get(cell.address);
        let value = hyperlink ? { text: cell.text, ...hyperlink }
          : cell.value?.sharedFormula ? { ...cell.value, formula } : cell.value;
        const formulaResult = formulaResults.get(cell.row)?.get(cell.address);
        if (value?.formula && formulaResult?.type === 'b') {
          value = { ...value, result: formulaResult.value !== '0' };
        } else if (value?.formula && formulaResult?.type === 'e') {
          value = { ...value, result: { error: excelUtils.xmlDecode(formulaResult.value) } };
        }
        if (value?.formula && typeof value.result === 'number' && excelUtils.isDateFmt(cell.numFmt)) {
          value = { ...value, result: excelUtils.excelToDate(value.result, workbook.properties.model?.date1904) };
        }
        return value;
      },
      releaseRow(rowNumber) {
        formulaResults.delete(rowNumber);
      }
    };
  } catch (error) {
    let cleanupError;
    if (sharedStrings) {
      try {
        await retryCleanup(() => sharedStrings.close());
      } catch (closeError) {
        cleanupError = closeError;
      }
    }
    if (cleanupError && error && (typeof error === 'object' || typeof error === 'function')) {
      try { error.cleanupError = cleanupError; } catch {}
    }
    throw error;
  }
}
