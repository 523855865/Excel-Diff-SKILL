# CSV Value and Type Columns Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make report CSV keys human-readable and split each changed value into separate value and type columns.

**Architecture:** Keep all typed tuples unchanged inside normalization and comparison. Convert typed values only at the `writeReport()` output boundary, using one small recursive projection for formula and hyperlink payloads and the existing reversible CSV-injection encoding.

**Tech Stack:** Node.js 24 ESM, `node:test`, Node.js standard library

---

### Task 1: Render readable CSV keys and split values from types

**Files:**
- Modify: `src/report.js:13-45`
- Modify: `test/report.test.js:59-220`
- Modify: `README.md:34-40`

- [ ] **Step 1: Write failing report tests for the new CSV contract**

Update the main report assertion so every file contributes `value,type,row`, values no longer contain the outer typed tuple, and single-field keys are direct values:

```js
assert.deepEqual(parseCsv(changedCsv), [
  ['key', 'sheet', 'column', 'B.value', 'B.type', 'B.row', 'A.value', 'A.type', 'A.row', 'C.value', 'C.type', 'C.row'],
  ['a', 'Data', 'alpha', '2', 'number', '2', '1', 'number', '2', '1', 'number', '2'],
  ['a', 'Data', 'zeta', '2', 'number', '2', '1', 'number', '2', '1', 'number', '2'],
  ['z', 'Data', 'beta', 'a,"b\nc', 'string', '3', JSON.stringify({ formula: '=SUM(A1:A2)', result: 2 }), 'formula', '4', '', 'blank', '5']
]);
assert.equal(missingCsv, 'key,sheet,presentFiles,missingFiles,baselineRelation\nhas|pipe,Data,"[""B|east"",""A""]","[""C|west""]",ADDED\n');
```

Add one focused case for composite keys and complex values:

```js
test('writes untyped composite keys and complex value payloads', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const result = reportResult({
    changed: [{
      key: [typed('string', '001'), typed('number', 2)],
      sheetName: 'Data',
      column: 'link',
      files: {
        A: { value: typed('hyperlink', { text: typed('string', 'Open'), target: 'https://a.test', tooltip: null }), rowNumber: 2 },
        B: { value: typed('hyperlink', { text: typed('string', 'Open'), target: 'https://b.test', tooltip: 'new' }), rowNumber: 2 }
      }
    }],
    missing: [{
      key: [typed('string', '001'), typed('number', 2)],
      sheetName: 'Data', presentFiles: ['A'], missingFiles: ['B'], baselineRelation: 'DELETED'
    }]
  });
  const report = await writeReport(reportSpec(directory, ['A', 'B']), result);
  const [changed, missing] = await Promise.all([
    readFile(join(report.directory, 'changed.csv'), 'utf8'),
    readFile(join(report.directory, 'missing.csv'), 'utf8')
  ]);
  assert.deepEqual(parseCsv(changed)[1], [
    JSON.stringify(['001', 2]), 'Data', 'link',
    JSON.stringify({ text: 'Open', target: 'https://a.test', tooltip: null }), 'hyperlink', '2',
    JSON.stringify({ text: 'Open', target: 'https://b.test', tooltip: 'new' }), 'hyperlink', '2'
  ]);
  assert.equal(parseCsv(missing)[1][0], JSON.stringify(['001', 2]));
});
```

Add a direct-key/value injection regression:

```js
test('protects direct keys and values from spreadsheet activation', async (t) => {
  const directory = await makeTempDir();
  t.after(() => rm(directory, { recursive: true, force: true }));
  const report = await writeReport(reportSpec(directory, ['A', 'B']), reportResult({
    changed: [{
      key: [typed('string', '=changed-key')], sheetName: 'Data', column: 'value',
      files: {
        A: { value: typed('string', '\t=changed-value'), rowNumber: 2 },
        B: { value: typed('string', 'safe'), rowNumber: 2 }
      }
    }],
    missing: [{
      key: [typed('string', '\nmissing-key')], sheetName: 'Data',
      presentFiles: ['A'], missingFiles: ['B'], baselineRelation: 'DELETED'
    }]
  }));
  const [changedHeader, changedRow] = parseCsv(await readFile(join(report.directory, 'changed.csv'), 'utf8'));
  const [, missingRow] = parseCsv(await readFile(join(report.directory, 'missing.csv'), 'utf8'));
  assert.equal(changedRow[0], `json:${JSON.stringify('=changed-key')}`);
  assert.equal(changedRow[3], `json:${JSON.stringify('\t=changed-value')}`);
  assert.equal(missingRow[0], `json:${JSON.stringify('\nmissing-key')}`);
  assert.equal(changedHeader.concat(changedRow, missingRow).some((value) => /^[=+\-@\t\r\n]/.test(value)), false);
});
```

Update the empty-report header assertion exactly:

```js
assert.equal(
  await readFile(join(first.directory, 'changed.csv'), 'utf8'),
  'key,sheet,column,B.value,B.type,B.row,A.value,A.type,A.row,C.value,C.type,C.row\n'
);
```

Update the existing scalar preservation assertions to read three columns per file:

```js
assert.deepEqual(
  values.map((_, index) => [row[3 + index * 3], row[4 + index * 3]]),
  [['', 'blank'], ['', 'string'], ['1', 'number'], ['1', 'string'], ['true', 'boolean'], ['true', 'string']]
);
assert.equal(formulaCell, `json:${JSON.stringify('=HYPERLINK("https://example.test","open")')}`);
```

- [ ] **Step 2: Run the focused tests and verify RED**

Run:

```bash
node --test test/report.test.js
```

Expected: FAIL because `changed.csv` still uses `value,row`, typed JSON values, and typed JSON keys.

- [ ] **Step 3: Implement the minimum output-boundary projection**

In `src/report.js`, add a recursive projection that removes typed wrappers without changing comparison data:

```js
function untypedValue([type, value]) {
  if (type === 'blank') return '';
  if (type === 'formula' && Array.isArray(value)) {
    return { formula: value[0], result: untypedValue(value[1]) };
  }
  if (type === 'hyperlink') {
    return { ...value, text: untypedValue(value.text) };
  }
  return value;
}

function safeValue(value) {
  const text = value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value ?? '');
  return safeText(text);
}

function displayKey(key) {
  const values = key.map(untypedValue);
  return safeValue(values.length === 1 ? values[0] : values);
}
```

Use it only while building report rows:

```js
const changedRows = [
  ['key', 'sheet', 'column', ...fileIds.flatMap((fileId) => [
    safeText(`${fileId}.value`), safeText(`${fileId}.type`), safeText(`${fileId}.row`)
  ])],
  ...changed.map((entry) => [
    displayKey(entry.key), safeText(entry.sheetName), safeText(entry.column),
    ...fileIds.flatMap((fileId) => {
      const [type] = entry.files[fileId].value;
      return [safeValue(untypedValue(entry.files[fileId].value)), type, entry.files[fileId].rowNumber];
    })
  ])
];
```

Use `displayKey(entry.key)` for `missingRows` as well. Do not mutate `entry.key`, typed values, sorting, summary generation, or artifact paths.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
node --test test/report.test.js
```

Expected: all report tests pass with the new headers, readable keys, split type columns, complex payloads, and injection protection.

- [ ] **Step 5: Update the Chinese README contract**

Replace the old typed-tuple sentence with the public format:

```markdown
- `changed.csv`：字段差异；每个文件按 `<fileId>.value`、`<fileId>.type`、`<fileId>.row` 分列，值列不包含类型包装；
- `missing.csv`：新增、删除的业务键；单字段 key 直接显示值，复合 key 显示不含类型包装的 JSON 数组。
```

State that dangerous key and value text also uses the existing reversible `json:<JSON string>` encoding.

- [ ] **Step 6: Run full verification**

Run:

```bash
npm test
git diff --check
npm pack --dry-run --cache /private/tmp/excel-diff-npm-cache
```

Expected: all tests pass; diff check is clean; package still contains README, schema, and runtime `src` files only.

- [ ] **Step 7: Commit the implementation**

```bash
git add src/report.js test/report.test.js README.md
git commit -m "feat: split csv values and types"
```
