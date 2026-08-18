import ExcelJS from 'exceljs';
import colCache from 'exceljs/lib/utils/col-cache.js';
import parseSax from 'exceljs/lib/utils/parse-sax.js';
import sharedFormulaUtils from 'exceljs/lib/utils/shared-formula.js';
import excelUtils from 'exceljs/lib/utils/utils.js';
import relationshipTypes from 'exceljs/lib/xlsx/rel-type.js';
import { StringDecoder } from 'node:string_decoder';
import unzipper from 'unzipper';

const { slideFormula } = sharedFormulaUtils;

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

export async function openStreamingWorkbook(filePath, sheetName) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'cache'
  });
  const archive = await unzipper.Open.file(filePath);
  const entries = new Map(archive.files.map((entry) => [entry.path, entry]));
  await workbook._parseRels(entries.get('xl/_rels/workbook.xml.rels').stream());
  await workbook._parseWorkbook(entries.get('xl/workbook.xml').stream());
  if (entries.has('xl/sharedStrings.xml')) {
    for await (const unused of workbook._parseSharedStrings(entries.get('xl/sharedStrings.xml').stream())) void unused;
  } else workbook.sharedStrings = [];
  if (entries.has('xl/styles.xml')) await workbook._parseStyles(entries.get('xl/styles.xml').stream());

  const hyperlinks = await readHyperlinks(entries, workbook, sheetName);
  const formulaResults = new Map();
  return {
    workbook,
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
}
