# Excel Diff Phase One Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Node.js CLI that validates a CompareSpec, compares two or more small XLSX files by business key, and writes `summary.json`, `changed.csv`, and `missing.csv`.

**Architecture:** Keep one deterministic pipeline: validate the spec before I/O, read each selected sheet into typed row records, aggregate by canonical business key, classify keys, then stream the bounded phase-one result arrays to artifact files. Split only at stable responsibilities: spec validation, value semantics, XLSX reading, comparison, reporting, and CLI orchestration.

**Tech Stack:** Node.js 24, JavaScript ESM, ExcelJS 4, Ajv 8, `node:test`, Node.js standard library.

---

## File map

- `package.json`: ESM package metadata, CLI entry, dependency and test commands.
- `package-lock.json`: npm-generated dependency lock.
- `schemas/compare-spec.schema.json`: strict structural CompareSpec contract.
- `src/spec.js`: schema compilation, semantic validation, and path resolution.
- `src/normalize.js`: typed canonical values, filters, equality, key encoding, and printable values.
- `src/read-xlsx.js`: sheet/header validation and row extraction.
- `src/compare.js`: key aggregation and classification.
- `src/report.js`: run directory, summary, and CSV artifacts.
- `src/cli.js`: argument parsing, pipeline orchestration, structured exits.
- `test/helpers.js`: temporary directory and XLSX fixture helpers.
- `test/spec.test.js`: structural and semantic spec validation.
- `test/normalize.test.js`: strict value semantics, filters, and key encoding.
- `test/read-xlsx.test.js`: workbook and column validation.
- `test/compare.test.js`: multi-file comparison behavior.
- `test/report.test.js`: CSV escaping and artifact layout.
- `test/cli.test.js`: end-to-end CLI and artifact contract.

### Task 1: Package and strict CompareSpec validation

**Files:**
- Create: `package.json`
- Create: `package-lock.json`
- Create: `schemas/compare-spec.schema.json`
- Create: `src/spec.js`
- Create: `test/spec.test.js`

- [ ] **Step 1: Create the package manifest**

```json
{
  "name": "excel-diff-skill",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "bin": { "excel-diff": "src/cli.js" },
  "scripts": { "test": "node --test" },
  "engines": { "node": ">=24" },
  "dependencies": { "ajv": "^8.17.1", "exceljs": "^4.4.0" }
}
```

- [ ] **Step 2: Install the two runtime dependencies**

Run: `npm install`

Expected: `package-lock.json` is generated and npm exits `0`.

- [ ] **Step 3: Write the failing spec tests**

Create `test/spec.test.js` with tests that call `loadSpec()` using JSON files in a temporary directory:

```javascript
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { loadSpec, SpecError } from '../src/spec.js';

const valid = {
  version: '1.0', baseline: 'A',
  files: [{ id: 'A', path: 'a.xlsx' }, { id: 'B', path: 'b.xlsx' }],
  sheet: { name: 'Data', headerRow: 1 },
  mode: { type: 'key', keyColumns: ['id'] },
  compareColumns: '*', filters: [],
  normalization: { emptyEqualsNull: false, caseSensitive: true, formulaMode: 'formula-and-cached-result', columns: {} },
  duplicateKeyPolicy: 'report',
  output: { directory: 'out', sampleSize: 20 }
};

async function writeSpec(value) {
  const dir = await mkdtemp(path.join(tmpdir(), 'excel-diff-spec-'));
  const file = path.join(dir, 'spec.json');
  await writeFile(file, JSON.stringify(value));
  return file;
}

test('resolves paths relative to the spec file', async () => {
  const file = await writeSpec(valid);
  const spec = await loadSpec(file);
  assert.equal(spec.files[0].path, path.join(path.dirname(file), 'a.xlsx'));
  assert.equal(spec.output.directory, path.join(path.dirname(file), 'out'));
});

test('rejects unknown properties before workbook I/O', async () => {
  await assert.rejects(loadSpec(await writeSpec({ ...valid, script: 'x' })), SpecError);
});

test('rejects duplicate ids, duplicate paths, and missing baseline', async () => {
  await assert.rejects(loadSpec(await writeSpec({ ...valid, files: [{ id: 'A', path: 'a.xlsx' }, { id: 'A', path: 'a.xlsx' }] })), /duplicate file id/i);
  await assert.rejects(loadSpec(await writeSpec({ ...valid, baseline: 'C' })), /baseline/i);
});
```

- [ ] **Step 4: Run the test and verify the missing module failure**

Run: `node --test test/spec.test.js`

Expected: FAIL with `ERR_MODULE_NOT_FOUND` for `src/spec.js`.

- [ ] **Step 5: Implement the strict schema and semantic validation**

Create `schemas/compare-spec.schema.json` with `additionalProperties: false` at every object level. Define these exact fields and enums:

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "type": "object",
  "additionalProperties": false,
  "required": ["version", "baseline", "files", "sheet", "mode", "compareColumns", "duplicateKeyPolicy", "output"],
  "properties": {
    "version": { "const": "1.0" },
    "baseline": { "type": "string", "minLength": 1 },
    "files": {
      "type": "array", "minItems": 2,
      "items": {
        "type": "object", "additionalProperties": false,
        "required": ["id", "path"],
        "properties": { "id": { "type": "string", "minLength": 1 }, "path": { "type": "string", "minLength": 1 } }
      }
    },
    "sheet": {
      "type": "object", "additionalProperties": false,
      "required": ["name", "headerRow"],
      "properties": { "name": { "type": "string", "minLength": 1 }, "headerRow": { "type": "integer", "minimum": 1 } }
    },
    "mode": {
      "type": "object", "additionalProperties": false,
      "required": ["type", "keyColumns"],
      "properties": { "type": { "const": "key" }, "keyColumns": { "type": "array", "minItems": 1, "uniqueItems": true, "items": { "type": "string", "minLength": 1 } } }
    },
    "compareColumns": { "oneOf": [{ "const": "*" }, { "type": "array", "minItems": 1, "uniqueItems": true, "items": { "type": "string", "minLength": 1 } }] },
    "columnAliases": { "type": "object", "additionalProperties": { "type": "array", "uniqueItems": true, "items": { "type": "string", "minLength": 1 } } },
    "filters": {
      "type": "array",
      "items": {
        "type": "object", "additionalProperties": false,
        "required": ["column", "operator"],
        "properties": {
          "column": { "type": "string", "minLength": 1 },
          "operator": { "enum": ["eq", "ne", "gt", "gte", "lt", "lte", "in", "notIn", "contains", "startsWith", "endsWith", "isNull", "isNotNull", "between"] },
          "value": {}, "values": { "type": "array" }
        }
      }
    },
    "normalization": {
      "type": "object", "additionalProperties": false,
      "properties": {
        "emptyEqualsNull": { "type": "boolean", "default": false },
        "caseSensitive": { "type": "boolean", "default": true },
        "formulaMode": { "enum": ["formula", "cached-result", "formula-and-cached-result"], "default": "formula-and-cached-result" },
        "columns": {
          "type": "object",
          "additionalProperties": {
            "type": "object", "additionalProperties": false,
            "properties": { "trim": { "type": "boolean" }, "caseSensitive": { "type": "boolean" }, "emptyEqualsNull": { "type": "boolean" }, "numericTolerance": { "type": "number", "minimum": 0 } }
          }
        }
      }
    },
    "duplicateKeyPolicy": { "enum": ["report", "fail"] },
    "output": {
      "type": "object", "additionalProperties": false,
      "required": ["directory"],
      "properties": { "directory": { "type": "string", "minLength": 1 }, "sampleSize": { "type": "integer", "minimum": 0, "default": 20 } }
    }
  }
}
```

Create `src/spec.js` with the complete validation flow:

```javascript
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import Ajv2020 from 'ajv/dist/2020.js';

const schema = JSON.parse(await readFile(new URL('../schemas/compare-spec.schema.json', import.meta.url)));
const validate = new Ajv2020({ allErrors: true, useDefaults: true }).compile(schema);

export class SpecError extends Error {
  constructor(message) {
    super(message);
    this.code = 'SPEC_INVALID';
  }
}

export async function loadSpec(specPath) {
  let parsed;
  try {
    parsed = JSON.parse(await readFile(specPath, 'utf8'));
  } catch (error) {
    throw new SpecError(`cannot read spec: ${error.message}`);
  }
  const spec = structuredClone(parsed);
  if (!validate(spec)) {
    const message = validate.errors.map(error => `${error.instancePath || '/'} ${error.message}`).join('; ');
    throw new SpecError(message);
  }
  const base = path.dirname(path.resolve(specPath));
  const ids = new Set();
  const paths = new Set();
  spec.files = spec.files.map(file => {
    if (ids.has(file.id)) throw new SpecError(`duplicate file id: ${file.id}`);
    ids.add(file.id);
    const resolved = path.resolve(base, file.path);
    if (paths.has(resolved)) throw new SpecError(`duplicate file path: ${resolved}`);
    if (path.extname(resolved).toLowerCase() !== '.xlsx') throw new SpecError(`input must be .xlsx: ${resolved}`);
    paths.add(resolved);
    return { ...file, path: resolved };
  });
  if (!ids.has(spec.baseline)) throw new SpecError(`baseline does not reference a file: ${spec.baseline}`);
  spec.filters ??= [];
  spec.columnAliases ??= {};
  spec.normalization ??= {};
  spec.normalization.columns ??= {};
  spec.output.directory = path.resolve(base, spec.output.directory);
  if (paths.has(spec.output.directory)) throw new SpecError('output directory overlaps an input file');
  for (const filter of spec.filters) {
    if (['isNull', 'isNotNull'].includes(filter.operator)) continue;
    if (['in', 'notIn', 'between'].includes(filter.operator)) {
      const values = filter.values ?? (Array.isArray(filter.value) ? filter.value : null);
      const validLength = filter.operator === 'between' ? values?.length === 2 : (values?.length ?? 0) > 0;
      if (!validLength) throw new SpecError(`${filter.operator} requires ${filter.operator === 'between' ? 'two' : 'one or more'} values`);
      continue;
    }
    if (!Object.hasOwn(filter, 'value')) throw new SpecError(`${filter.operator} requires value`);
  }
  return spec;
}
```

This formats schema errors without reading any workbook.

- [ ] **Step 6: Run the focused tests**

Run: `node --test test/spec.test.js`

Expected: 3 tests PASS.

- [ ] **Step 7: Commit the contract**

```bash
git add package.json package-lock.json schemas/compare-spec.schema.json src/spec.js test/spec.test.js
git commit -m "feat: validate compare specs"
```

### Task 2: Typed values, filters, and canonical keys

**Files:**
- Create: `src/normalize.js`
- Create: `test/normalize.test.js`

- [ ] **Step 1: Write failing value-semantics tests**

Create `test/normalize.test.js` that imports `normalizeValue`, `equalValues`, `encodeKey`, and `matchesFilter`. Assert:

```javascript
assert.notDeepEqual(normalizeValue(null, {}), normalizeValue('', {}));
assert.notDeepEqual(normalizeValue('001', {}), normalizeValue(1, {}));
assert.deepEqual(normalizeValue(' A ', { trim: true, caseSensitive: false }), ['string', 'a']);
assert.equal(equalValues(['number', 10], ['number', 10.005], { numericTolerance: 0.01 }), true);
assert.notEqual(encodeKey([['string', 'a|b'], ['string', 'c']]), encodeKey([['string', 'a'], ['string', 'b|c']]));
assert.equal(matchesFilter(['string', 'Finance'], { operator: 'startsWith', value: 'Fin' }), true);
assert.deepEqual(normalizeValue({ formula: 'A1+1', result: 2 }, { formulaMode: 'formula-and-cached-result' }), ['formula', ['A1+1', ['number', 2]]]);
```

- [ ] **Step 2: Verify the module is missing**

Run: `node --test test/normalize.test.js`

Expected: FAIL with `ERR_MODULE_NOT_FOUND` for `src/normalize.js`.

- [ ] **Step 3: Implement the minimum typed-value API**

Create `src/normalize.js` with these exports and behavior:

```javascript
export function normalizeValue(value, rule = {}) {
  if (value == null || (rule.emptyEqualsNull && value === '')) return ['blank', null];
  if (value instanceof Date) return ['date', value.toISOString()];
  if (typeof value === 'string') {
    let text = rule.trim ? value.trim() : value;
    if (rule.caseSensitive === false) text = text.toLocaleLowerCase();
    return ['string', text];
  }
  if (typeof value === 'number') return ['number', value];
  if (typeof value === 'boolean') return ['boolean', value];
  if (value.error) return ['error', value.error];
  if (value.richText) return normalizeValue(value.richText.map(part => part.text).join(''), rule);
  if (value.formula) {
    const result = normalizeValue(value.result, { ...rule, formulaMode: undefined });
    if (rule.formulaMode === 'formula') return ['formula', value.formula];
    if (rule.formulaMode === 'cached-result') return result;
    return ['formula', [value.formula, result]];
  }
  return ['string', String(value)];
}

export function equalValues(left, right, rule = {}) {
  if (left[0] === 'number' && right[0] === 'number' && rule.numericTolerance != null) {
    return Math.abs(left[1] - right[1]) <= rule.numericTolerance;
  }
  return JSON.stringify(left) === JSON.stringify(right);
}

export function encodeKey(values) {
  return JSON.stringify(values);
}

export function matchesFilter(value, filter, rule = {}) {
  const literal = item => normalizeValue(item, rule);
  const same = item => equalValues(value, literal(item));
  const scalar = value[1];
  const list = filter.values ?? (Array.isArray(filter.value) ? filter.value : []);
  const comparable = item => {
    const other = literal(item);
    return other[0] === value[0] ? [scalar, other[1]] : null;
  };

  switch (filter.operator) {
    case 'eq': return same(filter.value);
    case 'ne': return !same(filter.value);
    case 'in': return list.some(same);
    case 'notIn': return !list.some(same);
    case 'isNull': return value[0] === 'blank';
    case 'isNotNull': return value[0] !== 'blank';
    case 'contains': return typeof scalar === 'string' && scalar.includes(String(filter.value));
    case 'startsWith': return typeof scalar === 'string' && scalar.startsWith(String(filter.value));
    case 'endsWith': return typeof scalar === 'string' && scalar.endsWith(String(filter.value));
    case 'between': {
      if (list.length !== 2) return false;
      const low = comparable(list[0]);
      const high = comparable(list[1]);
      return low != null && high != null && scalar >= low[1] && scalar <= high[1];
    }
    default: {
      const pair = comparable(filter.value);
      if (pair == null) return false;
      if (filter.operator === 'gt') return pair[0] > pair[1];
      if (filter.operator === 'gte') return pair[0] >= pair[1];
      if (filter.operator === 'lt') return pair[0] < pair[1];
      if (filter.operator === 'lte') return pair[0] <= pair[1];
      return false;
    }
  }
}

export function printValue([type, value]) {
  if (type === 'blank') return '';
  return type === 'formula' || typeof value === 'object' ? JSON.stringify(value) : String(value);
}
```

Use `value.richText.map(part => part.text).join('')` for ExcelJS rich text. For formula objects, honor all three formula modes and recursively normalize the cached result. Apply `trim`, then case folding only to strings. When `emptyEqualsNull` is true, normalize the empty string to the same `['blank', null]` representation as a blank cell.

- [ ] **Step 4: Run focused tests**

Run: `node --test test/normalize.test.js`

Expected: all assertions PASS.

- [ ] **Step 5: Commit value semantics**

```bash
git add src/normalize.js test/normalize.test.js
git commit -m "feat: normalize spreadsheet values"
```

### Task 3: XLSX rows and deterministic column mapping

**Files:**
- Create: `test/helpers.js`
- Create: `src/read-xlsx.js`
- Create: `test/read-xlsx.test.js`

- [ ] **Step 1: Add the workbook fixture helper**

Create `test/helpers.js` exporting `makeTempDir()` and `writeWorkbook(file, sheetName, rows)`. Use `new ExcelJS.Workbook()`, `worksheet.addRows(rows)`, and `workbook.xlsx.writeFile(file)`; do not commit binary fixtures.

```javascript
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';

export function makeTempDir() {
  return mkdtemp(path.join(tmpdir(), 'excel-diff-test-'));
}

export async function writeWorkbook(file, sheetName, rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  sheet.addRows(rows);
  await workbook.xlsx.writeFile(file);
}
```

- [ ] **Step 2: Write failing reader tests**

Create `test/read-xlsx.test.js` with generated workbooks that assert:

- `readFileRows()` returns typed values plus `fileId`, `sheetName`, and original `rowNumber`.
- row order is preserved only as provenance.
- a missing sheet throws `SHEET_NOT_FOUND` and lists available sheets.
- duplicate headers throw `HEADER_DUPLICATED` with both positions.
- missing key, compare, or filter columns throw `COLUMN_MISSING`.
- NFKC-equivalent headers and explicit `columnAliases` map to the same standard column.
- two source columns mapping to one standard column throw `COLUMN_MAPPING_AMBIGUOUS`.

- [ ] **Step 3: Verify the reader module is missing**

Run: `node --test test/read-xlsx.test.js`

Expected: FAIL with `ERR_MODULE_NOT_FOUND` for `src/read-xlsx.js`.

- [ ] **Step 4: Implement `readFileRows(file, spec, standardColumns)`**

Create `src/read-xlsx.js`. Load the workbook with ExcelJS, select `spec.sheet.name`, build a raw header list from `spec.sheet.headerRow`, and map headers by exact text, then NFKC text, then explicit aliases. Reject ambiguity instead of taking the first match.

For every later row:

1. Build a standard-column object of normalized typed values using global defaults merged with `normalization.columns[column]`.
2. Apply every filter; skip the row unless all match.
3. If any key value is blank, increment `invalidRows` and omit the row.
4. Return `{ fileId, sheetName, rowNumber, values }`.

Return `{ columns, rows, invalidRows }`. When `compareColumns` is `"*"`, the baseline file establishes `standardColumns`; each later file must resolve every standard column.

Use this implementation shape, with every helper local to `src/read-xlsx.js`:

```javascript
import ExcelJS from 'exceljs';
import { matchesFilter, normalizeValue } from './normalize.js';

export class InputError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

const text = value => value == null ? '' : String(value);
const folded = value => text(value).normalize('NFKC');

function ruleFor(spec, column) {
  return {
    emptyEqualsNull: spec.normalization.emptyEqualsNull ?? false,
    caseSensitive: spec.normalization.caseSensitive ?? true,
    formulaMode: spec.normalization.formulaMode ?? 'formula-and-cached-result',
    ...spec.normalization.columns[column]
  };
}

function resolveColumn(raw, standards, aliases) {
  const exact = standards.filter(column => column === raw);
  if (exact.length === 1) return exact[0];
  const nfkc = standards.filter(column => folded(column) === folded(raw));
  if (nfkc.length === 1) return nfkc[0];
  const alias = standards.filter(column => (aliases[column] ?? []).some(value => folded(value) === folded(raw)));
  if (alias.length === 1) return alias[0];
  if (alias.length > 1 || nfkc.length > 1) throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `ambiguous column: ${raw}`);
  return null;
}

export async function readFileRows(file, spec, standardColumns = null) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(file.path);
  } catch (error) {
    throw new InputError('INPUT_ERROR', `cannot read ${file.id}: ${error.message}`);
  }
  const sheet = workbook.getWorksheet(spec.sheet.name);
  if (!sheet) throw new InputError('SHEET_NOT_FOUND', `${spec.sheet.name}; available: ${workbook.worksheets.map(item => item.name).join(', ')}`);
  const header = sheet.getRow(spec.sheet.headerRow);
  const rawHeaders = Array.from({ length: header.cellCount }, (_, index) => text(header.getCell(index + 1).value));
  const foldedHeaders = rawHeaders.map(folded);
  const duplicates = rawHeaders.filter((value, index) => value && foldedHeaders.indexOf(folded(value)) !== index);
  if (duplicates.length) throw new InputError('HEADER_DUPLICATED', `duplicate header: ${[...new Set(duplicates)].join(', ')}`);

  const standards = standardColumns ?? rawHeaders.filter(Boolean);
  const positions = new Map();
  rawHeaders.forEach((raw, index) => {
    const standard = standardColumns ? resolveColumn(raw, standards, spec.columnAliases) : raw;
    if (!standard) return;
    if (positions.has(standard)) throw new InputError('COLUMN_MAPPING_AMBIGUOUS', `multiple columns map to ${standard}`);
    positions.set(standard, index + 1);
  });
  const compareColumns = spec.compareColumns === '*' ? standards : spec.compareColumns;
  const required = new Set([...spec.mode.keyColumns, ...compareColumns, ...spec.filters.map(filter => filter.column)]);
  for (const column of required) {
    if (!positions.has(column)) throw new InputError('COLUMN_MISSING', `${file.id} is missing ${column}`);
  }

  const rows = [];
  let invalidRows = 0;
  let totalRowsScanned = 0;
  for (let rowNumber = spec.sheet.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    totalRowsScanned += 1;
    const row = sheet.getRow(rowNumber);
    const values = Object.fromEntries(standards.map(column => [
      column, normalizeValue(row.getCell(positions.get(column)).value, ruleFor(spec, column))
    ]));
    if (!spec.filters.every(filter => matchesFilter(values[filter.column], filter, ruleFor(spec, filter.column)))) continue;
    if (spec.mode.keyColumns.some(column => values[column][0] === 'blank')) {
      invalidRows += 1;
      continue;
    }
    rows.push({ fileId: file.id, sheetName: sheet.name, rowNumber, values });
  }
  return { columns: standards, rows, invalidRows, totalRowsScanned };
}
```

- [ ] **Step 5: Run reader tests**

Run: `node --test test/read-xlsx.test.js`

Expected: all reader tests PASS.

- [ ] **Step 6: Commit the reader**

```bash
git add src/read-xlsx.js test/helpers.js test/read-xlsx.test.js
git commit -m "feat: read keyed xlsx rows"
```

### Task 4: Aggregate and classify business keys

**Files:**
- Create: `src/compare.js`
- Create: `test/compare.test.js`

- [ ] **Step 1: Write failing multi-file comparison tests**

Create `test/compare.test.js` using generated A/B/C workbooks. Verify this exact scenario:

- key `1` appears in every file with the same values but different row order: `IDENTICAL`.
- key `2` appears in every file with one changed department: one changed key and one field detail.
- key `3` exists only in baseline A: one missing key with `baselineRelation: 'DELETED'`.
- key `4` exists only in B and C: one missing key with `baselineRelation: 'ADDED'`.
- key `5` is duplicated in B: one duplicate key; in `report` mode it produces no changed or missing detail.
- `duplicateKeyPolicy: 'fail'` rejects with `DUPLICATE_KEY`.

- [ ] **Step 2: Verify the comparison module is missing**

Run: `node --test test/compare.test.js`

Expected: FAIL with `ERR_MODULE_NOT_FOUND` for `src/compare.js`.

- [ ] **Step 3: Implement `compare(spec)`**

Create `src/compare.js` that:

1. Reads the baseline first to establish standard columns.
2. Reads every remaining file against those columns.
3. Builds `Map<encodeKey(keyValues), Map<fileId, RowRecord[]>>`.
4. Sorts encoded keys before classification so output is deterministic.
5. Applies duplicate, missing, identical, and changed classification in that order.
6. Compares only non-key selected columns with `equalValues()` and their column rule.
7. Returns `{ summary, changed, missing, duplicates }` without writing files.

The returned summary must contain `files`, `totalRowsScanned`, `matchedRows`, `identicalKeys`, `changedKeys`, `missingKeys`, `duplicateKeys`, and `invalidRows`. A changed entry has `{ key, sheetName, column, files: { [fileId]: { value, rowNumber } } }`. A missing entry has `{ key, sheetName, presentFiles, missingFiles, baselineRelation }`.

Implement the classifier directly:

```javascript
import { encodeKey, equalValues } from './normalize.js';
import { readFileRows } from './read-xlsx.js';

export class CompareError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

export async function compare(spec) {
  const baseline = spec.files.find(file => file.id === spec.baseline);
  const ordered = [baseline, ...spec.files.filter(file => file.id !== spec.baseline)];
  const loaded = [];
  let columns;
  for (const file of ordered) {
    const result = await readFileRows(file, spec, columns);
    columns ??= result.columns;
    loaded.push(result);
  }

  const records = new Map();
  for (const result of loaded) {
    for (const row of result.rows) {
      const key = encodeKey(spec.mode.keyColumns.map(column => row.values[column]));
      const byFile = records.get(key) ?? new Map();
      const rows = byFile.get(row.fileId) ?? [];
      rows.push(row);
      byFile.set(row.fileId, rows);
      records.set(key, byFile);
    }
  }

  const changed = [];
  const missing = [];
  const duplicates = [];
  let identicalKeys = 0;
  for (const key of [...records.keys()].sort()) {
    const byFile = records.get(key);
    const duplicateFiles = spec.files.filter(file => (byFile.get(file.id)?.length ?? 0) > 1).map(file => file.id);
    if (duplicateFiles.length) {
      if (spec.duplicateKeyPolicy === 'fail') throw new CompareError('DUPLICATE_KEY', `${key}: ${duplicateFiles.join(', ')}`);
      duplicates.push({ key, files: duplicateFiles });
      continue;
    }
    const presentFiles = spec.files.filter(file => byFile.has(file.id)).map(file => file.id);
    if (presentFiles.length !== spec.files.length) {
      const missingFiles = spec.files.filter(file => !byFile.has(file.id)).map(file => file.id);
      missing.push({
        key, sheetName: spec.sheet.name, presentFiles, missingFiles,
        baselineRelation: byFile.has(spec.baseline) ? 'DELETED' : 'ADDED'
      });
      continue;
    }
    const compareColumns = (spec.compareColumns === '*' ? columns : spec.compareColumns)
      .filter(column => !spec.mode.keyColumns.includes(column));
    let keyChanged = false;
    for (const column of compareColumns) {
      const rows = Object.fromEntries(spec.files.map(file => [file.id, byFile.get(file.id)[0]]));
      const first = rows[spec.baseline].values[column];
      const rule = { ...spec.normalization, ...spec.normalization.columns[column] };
      if (spec.files.every(file => equalValues(first, rows[file.id].values[column], rule))) continue;
      keyChanged = true;
      changed.push({
        key, sheetName: spec.sheet.name, column,
        files: Object.fromEntries(spec.files.map(file => [file.id, {
          value: rows[file.id].values[column], rowNumber: rows[file.id].rowNumber
        }]))
      });
    }
    if (!keyChanged) identicalKeys += 1;
  }

  return {
    summary: {
      files: spec.files.length,
      totalRowsScanned: loaded.reduce((sum, item) => sum + item.totalRowsScanned, 0),
      matchedRows: loaded.reduce((sum, item) => sum + item.rows.length, 0),
      identicalKeys,
      changedKeys: new Set(changed.map(item => item.key)).size,
      missingKeys: missing.length,
      duplicateKeys: duplicates.length,
      invalidRows: loaded.reduce((sum, item) => sum + item.invalidRows, 0)
    },
    changed, missing, duplicates
  };
}
```

- [ ] **Step 4: Run comparison tests**

Run: `node --test test/compare.test.js`

Expected: all comparison tests PASS with the documented counts.

- [ ] **Step 5: Commit the comparison engine**

```bash
git add src/compare.js test/compare.test.js
git commit -m "feat: compare rows by business key"
```

### Task 5: Write run artifacts

**Files:**
- Create: `src/report.js`
- Create: `test/report.test.js`

- [ ] **Step 1: Write failing report tests**

Create `test/report.test.js` with an in-memory comparison result containing commas, quotes, and newlines. Assert that `writeReport()`:

- creates one run-id child directory under the configured output directory;
- writes valid `summary.json` with `status: 'COMPLETED'` and relative artifact names;
- writes `changed.csv` and `missing.csv` with UTF-8 text;
- doubles quotes and quotes fields containing commas, quotes, CR, or LF;
- emits dynamic `<fileId>.value` and `<fileId>.row` columns in spec file order.

- [ ] **Step 2: Verify the report module is missing**

Run: `node --test test/report.test.js`

Expected: FAIL with `ERR_MODULE_NOT_FOUND` for `src/report.js`.

- [ ] **Step 3: Implement artifact writing**

Create `src/report.js` with:

```javascript
import { randomUUID } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { printValue } from './normalize.js';

export function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

export async function writeReport(spec, result) {
  const runId = `${new Date().toISOString().replace(/\D/g, '').slice(0, 14)}Z-${randomUUID().slice(0, 8)}`;
  const directory = path.join(spec.output.directory, runId);
  await mkdir(directory, { recursive: true });
  const fileIds = spec.files.map(file => file.id);
  const changedHeader = ['key', 'sheet', 'column', ...fileIds.flatMap(id => [`${id}.value`, `${id}.row`])];
  const changedRows = result.changed.map(item => [
    item.key, item.sheetName, item.column,
    ...fileIds.flatMap(id => [printValue(item.files[id].value), item.files[id].rowNumber])
  ]);
  const missingRows = result.missing.map(item => [
    item.key, item.sheetName, item.presentFiles.join('|'), item.missingFiles.join('|'), item.baselineRelation
  ]);
  const toCsv = rows => `${rows.map(row => row.map(csvCell).join(',')).join('\n')}\n`;
  await writeFile(path.join(directory, 'changed.csv'), toCsv([changedHeader, ...changedRows]));
  await writeFile(path.join(directory, 'missing.csv'), toCsv([
    ['key', 'sheet', 'presentFiles', 'missingFiles', 'baselineRelation'], ...missingRows
  ]));
  const summary = {
    status: 'COMPLETED', runId, ...result.summary,
    artifacts: { changed: 'changed.csv', missing: 'missing.csv' }
  };
  await writeFile(path.join(directory, 'summary.json'), `${JSON.stringify(summary, null, 2)}\n`);
  return { directory, summary };
}
```

Use a run ID built from an ISO UTC timestamp without punctuation plus the first eight characters of `randomUUID()`. Write CSV headers even when there are no details. Sort details by encoded key and then column. Write `summary.json` last, adding `status`, `runId`, and `artifacts` to the engine summary.

- [ ] **Step 4: Run report tests**

Run: `node --test test/report.test.js`

Expected: all report tests PASS.

- [ ] **Step 5: Commit reporting**

```bash
git add src/report.js test/report.test.js
git commit -m "feat: write comparison artifacts"
```

### Task 6: CLI end-to-end closure

**Files:**
- Create: `src/cli.js`
- Create: `test/cli.test.js`
- Modify: `README.md`

- [ ] **Step 1: Write the failing CLI test**

Create `test/cli.test.js`. Generate two workbooks and one spec, invoke `process.execPath` with `src/cli.js compare --spec <path>` through `spawnSync`, and assert:

- exit status `0`;
- stdout parses as one JSON object with `status: 'COMPLETED'`;
- stderr is empty;
- the returned run directory contains all three artifacts;
- an invalid spec exits `2` and stderr contains a JSON error with `code: 'SPEC_INVALID'`;
- a missing input file exits `4` with `code: 'INPUT_ERROR'`;
- an unknown command or missing `--spec` exits `2`.

- [ ] **Step 2: Verify the CLI is missing**

Run: `node --test test/cli.test.js`

Expected: FAIL because `src/cli.js` does not exist.

- [ ] **Step 3: Implement the CLI**

Create executable `src/cli.js` with a small manual parser accepting only `compare --spec <path>`. Call `loadSpec()`, `compare()`, and `writeReport()` in order. Print only `JSON.stringify(summary)` to stdout. Print `{ status: 'FAILED', code, message }` JSON to stderr.

Map `SpecError` and usage errors to exit `2`, missing/unreadable/XLSX input errors to `4`, and unexpected errors to `6`. Do not print stack traces unless `EXCEL_DIFF_DEBUG=1`.

Use this complete orchestration:

```javascript
#!/usr/bin/env node
import { compare, CompareError } from './compare.js';
import { InputError } from './read-xlsx.js';
import { writeReport } from './report.js';
import { loadSpec, SpecError } from './spec.js';

class UsageError extends Error {
  constructor(message) {
    super(message);
    this.code = 'SPEC_INVALID';
  }
}

async function main(args) {
  if (args.length !== 3 || args[0] !== 'compare' || args[1] !== '--spec') {
    throw new UsageError('usage: excel-diff compare --spec <compare-spec.json>');
  }
  const spec = await loadSpec(args[2]);
  const result = await compare(spec);
  const report = await writeReport(spec, result);
  process.stdout.write(`${JSON.stringify({ ...report.summary, directory: report.directory })}\n`);
}

main(process.argv.slice(2)).catch(error => {
  const knownInput = error instanceof InputError || error instanceof CompareError;
  const exitCode = error instanceof SpecError || error instanceof UsageError ? 2 : knownInput ? 4 : 6;
  const code = error.code ?? (knownInput ? 'INPUT_ERROR' : 'INTERNAL_ERROR');
  process.stderr.write(`${JSON.stringify({ status: 'FAILED', code, message: error.message })}\n`);
  if (process.env.EXCEL_DIFF_DEBUG === '1' && error.stack) process.stderr.write(`${error.stack}\n`);
  process.exitCode = exitCode;
});
```

- [ ] **Step 4: Document the runnable phase-one flow**

Replace the one-line README with installation, the exact CLI command, a complete two-file key-mode CompareSpec example, output paths, strict-default warnings, phase-one limits, and `npm test`.

````markdown
# Excel Diff

Deterministic small-file XLSX comparison by business key.

```bash
npm install
node src/cli.js compare --spec compare-spec.json
```

```json
{
  "version": "1.0",
  "baseline": "A",
  "files": [{ "id": "A", "path": "A.xlsx" }, { "id": "B", "path": "B.xlsx" }],
  "sheet": { "name": "Data", "headerRow": 1 },
  "mode": { "type": "key", "keyColumns": ["id"] },
  "compareColumns": "*",
  "filters": [],
  "normalization": {
    "emptyEqualsNull": false,
    "caseSensitive": true,
    "formulaMode": "formula-and-cached-result",
    "columns": {}
  },
  "duplicateKeyPolicy": "report",
  "output": { "directory": "output", "sampleSize": 20 }
}
```

Each run writes `summary.json`, `changed.csv`, and `missing.csv` under `output/<run-id>/`.

Defaults are strict: whitespace, case, blank cells, empty strings, numbers, and numeric strings remain distinct. Formula results are read from the workbook cache and are not recalculated.

Phase one loads selected worksheets into memory and supports only `key` mode. Streaming partitions, `row`, `multiset`, and Agent wrappers follow in later phases.

```bash
npm test
```
````

- [ ] **Step 5: Run the end-to-end test**

Run: `node --test test/cli.test.js`

Expected: all CLI tests PASS.

- [ ] **Step 6: Run the complete verification suite**

Run: `npm test`

Expected: every test file PASS and process exits `0`.

Run: `npm pack --dry-run`

Expected: package contents include `src/`, `schemas/`, and `README.md`; no test output or temporary XLSX files are included.

- [ ] **Step 7: Commit the runnable phase**

```bash
git add src/cli.js test/cli.test.js README.md
git commit -m "feat: add excel diff cli"
```

### Task 7: Final contract and hygiene verification

**Files:**
- Modify only if a verification failure proves a defect in a phase-one file.

- [ ] **Step 1: Check the working tree without touching unrelated files**

Run: `git status --short`

Expected: pre-existing `.DS_Store` and `graphify-out/` may remain untracked; no generated XLSX, output directory, or npm debug log appears.

- [ ] **Step 2: Re-run all tests from a clean dependency install**

Run: `npm ci`

Expected: dependency install exits `0` using `package-lock.json`.

Run: `npm test`

Expected: all tests PASS.

- [ ] **Step 3: Exercise a real generated comparison**

Run the CLI against the fixtures created by `test/cli.test.js` through the test itself, then read its summary and both CSV artifacts. Confirm counts match the test scenario and every changed row includes sheet, key, column, values, and original row numbers.

- [ ] **Step 4: Verify no accidental scope expansion**

Run: `git diff --stat 939e55f..HEAD`

Expected: only the phase-one package, schema, source, tests, README, and implementation plan changed. No streaming partitions, `row`/`multiset`, Skill wrapper, plugin, UI, or database code exists.
