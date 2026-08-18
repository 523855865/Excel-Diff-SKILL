import ExcelJS from 'exceljs';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

export function makeTempDir() {
  return mkdtemp(join(tmpdir(), 'excel-diff-test-'));
}

export async function writeWorkbook(file, sheetName, rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet(sheetName).addRows(rows);
  await workbook.xlsx.writeFile(file);
}
