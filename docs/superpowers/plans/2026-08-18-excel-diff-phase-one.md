# Excel Diff 阶段一实施计划

> **供自动化执行代理使用：** 必须使用 `superpowers:subagent-driven-development`（推荐）或 `superpowers:executing-plans` 子技能，逐项实施本计划。步骤使用复选框（`- [ ]`）跟踪状态。

**目标：** 构建一个 Node.js CLI，用于校验 CompareSpec、按业务主键比较两个或更多小型 XLSX 文件，并写出 `summary.json`、`changed.csv` 和 `missing.csv`。

**架构：** 保持一条确定性处理链：在 I/O 前校验规则，将每个选定工作表读取为带类型的行记录，按规范业务主键聚合并分类，然后把阶段一的结果数组写入产物文件。只按稳定职责拆分：规则校验、值语义、XLSX 读取、比较、报告和 CLI 编排。

**技术栈：** Node.js 24、JavaScript ESM、ExcelJS 4、Ajv 8、`node:test`、Node.js 标准库。

---

## 文件映射

- `package.json`：ESM 包元数据、CLI 入口、依赖和测试命令。
- `package-lock.json`：npm 生成的依赖锁定文件。
- `schemas/compare-spec.schema.json`：严格的 CompareSpec 结构契约。
- `src/spec.js`：Schema 编译、语义校验和路径解析。
- `src/normalize.js`：带类型的规范值、过滤、相等判断、主键编码和可打印值。
- `src/read-xlsx.js`：工作表及表头校验与行提取。
- `src/compare.js`：主键聚合与分类。
- `src/report.js`：运行目录、摘要和 CSV 产物。
- `src/cli.js`：参数解析、处理链编排和结构化退出。
- `test/helpers.js`：临时目录和 XLSX 测试数据辅助函数。
- `test/spec.test.js`：规则结构与语义校验。
- `test/normalize.test.js`：严格值语义、过滤和主键编码。
- `test/read-xlsx.test.js`：工作簿和列校验。
- `test/compare.test.js`：多文件比较行为。
- `test/report.test.js`：CSV 转义和产物布局。
- `test/cli.test.js`：CLI 端到端流程和产物契约。

### 任务 1：包配置与严格 CompareSpec 校验

**文件：**
- 新建：`package.json`
- 新建：`package-lock.json`
- 新建：`schemas/compare-spec.schema.json`
- 新建：`src/spec.js`
- 新建：`test/spec.test.js`

- [ ] **步骤 1：创建包清单**

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

- [ ] **步骤 2：安装两个运行时依赖**

运行：`npm install`

预期：生成 `package-lock.json`，npm 以状态码 `0` 退出。

- [ ] **步骤 3：编写预期失败的规则测试**

创建 `test/spec.test.js`，使用临时目录中的 JSON 文件调用 `loadSpec()`：

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

- [ ] **步骤 4：运行测试并确认因模块缺失而失败**

运行：`node --test test/spec.test.js`

预期：测试失败，并针对 `src/spec.js` 报出 `ERR_MODULE_NOT_FOUND`。

- [ ] **步骤 5：实现严格 Schema 与语义校验**

创建 `schemas/compare-spec.schema.json`，在每层对象上设置 `additionalProperties: false`，并精确定义以下字段和枚举：

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

创建 `src/spec.js`，实现完整校验流程：

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

该流程在不读取任何工作簿的前提下格式化 Schema 错误。

- [ ] **步骤 6：运行聚焦测试**

运行：`node --test test/spec.test.js`

预期：3 个测试通过。

- [ ] **步骤 7：提交规则契约**

```bash
git add package.json package-lock.json schemas/compare-spec.schema.json src/spec.js test/spec.test.js
git commit -m "feat: validate compare specs"
```

### 任务 2：带类型的值、过滤与规范主键

**文件：**
- 新建：`src/normalize.js`
- 新建：`test/normalize.test.js`

- [ ] **步骤 1：编写预期失败的值语义测试**

创建 `test/normalize.test.js`，导入 `normalizeValue`、`equalValues`、`encodeKey` 和 `matchesFilter`，并断言：

```javascript
assert.notDeepEqual(normalizeValue(null, {}), normalizeValue('', {}));
assert.notDeepEqual(normalizeValue('001', {}), normalizeValue(1, {}));
assert.deepEqual(normalizeValue(' A ', { trim: true, caseSensitive: false }), ['string', 'a']);
assert.equal(equalValues(['number', 10], ['number', 10.005], { numericTolerance: 0.01 }), true);
assert.notEqual(encodeKey([['string', 'a|b'], ['string', 'c']]), encodeKey([['string', 'a'], ['string', 'b|c']]));
assert.equal(matchesFilter(['string', 'Finance'], { operator: 'startsWith', value: 'Fin' }), true);
assert.deepEqual(normalizeValue({ formula: 'A1+1', result: 2 }, { formulaMode: 'formula-and-cached-result' }), ['formula', ['A1+1', ['number', 2]]]);
```

- [ ] **步骤 2：确认模块缺失**

运行：`node --test test/normalize.test.js`

预期：测试失败，并针对 `src/normalize.js` 报出 `ERR_MODULE_NOT_FOUND`。

- [ ] **步骤 3：实现最小的带类型值 API**

创建 `src/normalize.js`，提供以下导出和行为：

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

使用 `value.richText.map(part => part.text).join('')` 处理 ExcelJS 富文本。公式对象需遵守三种公式模式，并递归规范化缓存结果。只对字符串先执行 `trim`，再按规则转换大小写。当 `emptyEqualsNull` 为 `true` 时，将空字符串规范化为与空单元格相同的 `['blank', null]`。

- [ ] **步骤 4：运行聚焦测试**

运行：`node --test test/normalize.test.js`

预期：所有断言通过。

- [ ] **步骤 5：提交值语义实现**

```bash
git add src/normalize.js test/normalize.test.js
git commit -m "feat: normalize spreadsheet values"
```

### 任务 3：XLSX 行读取与确定性列映射

**文件：**
- 新建：`test/helpers.js`
- 新建：`src/read-xlsx.js`
- 新建：`test/read-xlsx.test.js`

- [ ] **步骤 1：添加工作簿测试数据辅助函数**

创建 `test/helpers.js`，导出 `makeTempDir()` 和 `writeWorkbook(file, sheetName, rows)`。使用 `new ExcelJS.Workbook()`、`worksheet.addRows(rows)` 和 `workbook.xlsx.writeFile(file)`，不提交二进制测试文件。

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

- [ ] **步骤 2：编写预期失败的读取器测试**

创建 `test/read-xlsx.test.js`，用动态生成的工作簿断言：

- `readFileRows()` 返回带类型的值以及 `fileId`、`sheetName` 和原始 `rowNumber`。
- 行顺序只作为溯源信息保留。
- 缺少工作表时抛出 `SHEET_NOT_FOUND`，并列出可用工作表。
- 重复表头抛出 `HEADER_DUPLICATED`，并给出两个位置。
- 缺少主键列、比较列或过滤列时抛出 `COLUMN_MISSING`。
- NFKC 等价表头和显式 `columnAliases` 均映射到同一标准列。
- 两个源列映射到同一标准列时抛出 `COLUMN_MAPPING_AMBIGUOUS`。

- [ ] **步骤 3：确认读取器模块缺失**

运行：`node --test test/read-xlsx.test.js`

预期：测试失败，并针对 `src/read-xlsx.js` 报出 `ERR_MODULE_NOT_FOUND`。

- [ ] **步骤 4：实现 `readFileRows(file, spec, standardColumns)`**

创建 `src/read-xlsx.js`。使用 ExcelJS 加载工作簿，选择 `spec.sheet.name`，从 `spec.sheet.headerRow` 构建原始表头列表，并依次按原文、NFKC 文本和显式别名映射表头。遇到歧义时拒绝处理，不能直接选择第一个结果。

对后续每一行：

1. 合并全局默认规则与 `normalization.columns[column]`，构建标准列到规范化带类型值的对象。
2. 应用全部过滤条件；只有全部匹配才保留该行。
3. 任一主键值为空时增加 `invalidRows`，并忽略该行。
4. 返回 `{ fileId, sheetName, rowNumber, values }`。

返回 `{ columns, rows, invalidRows }`。当 `compareColumns` 为 `"*"` 时，由基准文件确定 `standardColumns`；后续每个文件都必须解析出所有标准列。

采用以下实现结构，并将全部辅助函数保留在 `src/read-xlsx.js` 内：

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

- [ ] **步骤 5：运行读取器测试**

运行：`node --test test/read-xlsx.test.js`

预期：所有读取器测试通过。

- [ ] **步骤 6：提交读取器实现**

```bash
git add src/read-xlsx.js test/helpers.js test/read-xlsx.test.js
git commit -m "feat: read keyed xlsx rows"
```

### 任务 4：聚合并分类业务主键

**文件：**
- 新建：`src/compare.js`
- 新建：`test/compare.test.js`

- [ ] **步骤 1：编写预期失败的多文件比较测试**

创建 `test/compare.test.js`，使用动态生成的 A/B/C 工作簿验证以下精确场景：

- 主键 `1` 在所有文件中都存在，值相同但行顺序不同：分类为 `IDENTICAL`。
- 主键 `2` 在所有文件中都存在，但部门字段不同：产生一个变化主键和一条字段明细。
- 主键 `3` 仅存在于基准文件 A：产生一个缺失主键，且 `baselineRelation: 'DELETED'`。
- 主键 `4` 仅存在于 B 和 C：产生一个缺失主键，且 `baselineRelation: 'ADDED'`。
- 主键 `5` 在 B 中重复：产生一个重复主键；在 `report` 模式下不生成变化或缺失明细。
- `duplicateKeyPolicy: 'fail'` 时以 `DUPLICATE_KEY` 拒绝执行。

- [ ] **步骤 2：确认比较模块缺失**

运行：`node --test test/compare.test.js`

预期：测试失败，并针对 `src/compare.js` 报出 `ERR_MODULE_NOT_FOUND`。

- [ ] **步骤 3：实现 `compare(spec)`**

创建 `src/compare.js`，完成以下处理：

1. 先读取基准文件以确定标准列。
2. 按这些标准列读取其余每个文件。
3. 构建 `Map<encodeKey(keyValues), Map<fileId, RowRecord[]>>`。
4. 分类前对编码后的主键排序，保证输出确定性。
5. 依次执行重复、缺失、一致和变化分类。
6. 只使用 `equalValues()` 和对应列规则比较选中的非主键列。
7. 返回 `{ summary, changed, missing, duplicates }`，不在此处写文件。

返回的摘要必须包含 `files`、`totalRowsScanned`、`matchedRows`、`identicalKeys`、`changedKeys`、`missingKeys`、`duplicateKeys` 和 `invalidRows`。变化项结构为 `{ key, sheetName, column, files: { [fileId]: { value, rowNumber } } }`；缺失项结构为 `{ key, sheetName, presentFiles, missingFiles, baselineRelation }`。

直接实现分类器：

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

- [ ] **步骤 4：运行比较测试**

运行：`node --test test/compare.test.js`

预期：所有比较测试按文档规定的计数通过。

- [ ] **步骤 5：提交比较引擎**

```bash
git add src/compare.js test/compare.test.js
git commit -m "feat: compare rows by business key"
```

### 任务 5：写出运行产物

**文件：**
- 新建：`src/report.js`
- 新建：`test/report.test.js`

- [ ] **步骤 1：编写预期失败的报告测试**

创建 `test/report.test.js`，构造包含逗号、引号和换行符的内存比较结果，并断言 `writeReport()`：

- 在配置的输出目录下创建一个运行 ID 子目录；
- 写出有效的 `summary.json`，其中包含 `status: 'COMPLETED'` 和相对产物名；
- 以 UTF-8 文本写出 `changed.csv` 和 `missing.csv`；
- 将引号转义为双引号，并为包含逗号、引号、CR 或 LF 的字段加引号；
- 按规则中的文件顺序生成动态 `<fileId>.value` 和 `<fileId>.row` 列。

- [ ] **步骤 2：确认报告模块缺失**

运行：`node --test test/report.test.js`

预期：测试失败，并针对 `src/report.js` 报出 `ERR_MODULE_NOT_FOUND`。

- [ ] **步骤 3：实现产物写出**

创建 `src/report.js`：

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

运行 ID 由去除标点的 ISO UTC 时间戳和 `randomUUID()` 前八个字符组成。即使没有明细，也要写出 CSV 表头。先按编码主键、再按列名排序明细。最后写 `summary.json`，并在引擎摘要上增加 `status`、`runId` 和 `artifacts`。

- [ ] **步骤 4：运行报告测试**

运行：`node --test test/report.test.js`

预期：所有报告测试通过。

- [ ] **步骤 5：提交报告实现**

```bash
git add src/report.js test/report.test.js
git commit -m "feat: write comparison artifacts"
```

### 任务 6：完成 CLI 端到端闭环

**文件：**
- 新建：`src/cli.js`
- 新建：`test/cli.test.js`
- 修改：`README.md`

- [ ] **步骤 1：编写预期失败的 CLI 测试**

创建 `test/cli.test.js`。生成两个工作簿和一份规则，通过 `spawnSync` 使用 `process.execPath` 调用 `src/cli.js compare --spec <path>`，并断言：

- 退出状态为 `0`；
- 标准输出可解析为一个包含 `status: 'COMPLETED'` 的 JSON 对象；
- 标准错误为空；
- 返回的运行目录包含三种产物；
- 非法规则以状态码 `2` 退出，标准错误包含 `code: 'SPEC_INVALID'` 的 JSON 错误；
- 输入文件缺失时以状态码 `4` 退出，并包含 `code: 'INPUT_ERROR'`；
- 未知命令或缺少 `--spec` 时以状态码 `2` 退出。

- [ ] **步骤 2：确认 CLI 缺失**

运行：`node --test test/cli.test.js`

预期：由于 `src/cli.js` 不存在而测试失败。

- [ ] **步骤 3：实现 CLI**

创建可执行的 `src/cli.js`，使用小型手写解析器且只接受 `compare --spec <path>`。依次调用 `loadSpec()`、`compare()` 和 `writeReport()`。标准输出只打印 `JSON.stringify(summary)`，标准错误打印 `{ status: 'FAILED', code, message }` JSON。

将 `SpecError` 和用法错误映射到退出码 `2`，文件缺失、不可读或 XLSX 输入错误映射到 `4`，意外错误映射到 `6`。只有 `EXCEL_DIFF_DEBUG=1` 时才打印堆栈。

使用以下完整编排代码：

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

- [ ] **步骤 4：记录阶段一可运行流程**

将单行 README 替换为安装方法、准确的 CLI 命令、完整的两文件主键模式 CompareSpec 示例、输出路径、严格默认值提示、阶段一限制和 `npm test`。

````markdown
# Excel Diff

按业务主键确定性比较小型 XLSX 文件。

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

每次运行都会在 `output/<run-id>/` 下写出 `summary.json`、`changed.csv` 和 `missing.csv`。

默认规则严格：空格、大小写、空单元格、空字符串、数字和数字字符串均保持区别。公式结果读取工作簿中的缓存值，不重新计算。

阶段一将选定工作表加载到内存，且只支持 `key` 模式。流式分区、`row`、`multiset` 和 Agent 包装将在后续阶段实现。

```bash
npm test
```
````

- [ ] **步骤 5：运行端到端测试**

运行：`node --test test/cli.test.js`

预期：所有 CLI 测试通过。

- [ ] **步骤 6：运行完整验证套件**

运行：`npm test`

预期：所有测试文件通过，进程以状态码 `0` 退出。

运行：`npm pack --dry-run`

预期：包内容包含 `src/`、`schemas/` 和 `README.md`，不包含测试输出或临时 XLSX 文件。

- [ ] **步骤 7：提交可运行的阶段一实现**

```bash
git add src/cli.js test/cli.test.js README.md
git commit -m "feat: add excel diff cli"
```

### 任务 7：最终契约与仓库卫生验证

**文件：**
- 只有验证失败证明阶段一文件存在缺陷时才修改。

- [ ] **步骤 1：检查工作区且不触碰无关文件**

运行：`git status --short`

预期：已有的 `.DS_Store` 和 `graphify-out/` 可以继续保持未跟踪；不得出现生成的 XLSX、输出目录或 npm 调试日志。

- [ ] **步骤 2：基于全新依赖安装重新运行全部测试**

运行：`npm ci`

预期：使用 `package-lock.json` 安装依赖，并以状态码 `0` 退出。

运行：`npm test`

预期：所有测试通过。

- [ ] **步骤 3：执行一次真实的动态生成数据比较**

通过测试本身，使用 CLI 比较 `test/cli.test.js` 创建的测试数据，然后读取摘要和两个 CSV 产物。确认计数与测试场景一致，并且每条变化记录都包含工作表、主键、列、值和原始行号。

- [ ] **步骤 4：确认没有意外扩大范围**

运行：`git diff --stat 939e55f..HEAD`

预期：只有阶段一的包配置、Schema、源码、测试、README 和实施计划发生变化；不存在流式分区、`row`/`multiset`、Skill 包装、插件、UI 或数据库代码。
