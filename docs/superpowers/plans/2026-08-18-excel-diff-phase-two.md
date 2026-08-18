# Excel Diff Phase Two Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the phase-one in-memory path with a streaming, disk-partitioned comparison engine that supports `key`, `row`, and `multiset` modes, bounded resources, atomic reports, failure summaries, and opt-in progress events.

**Architecture:** Keep the existing CompareSpec validation, normalization, column mapping, typed values, and CSV safety rules. Stream each workbook once into deterministic SHA-256 JSONL partitions, process one bounded partition at a time (recursively repartitioning oversized buckets), and send result records to a report sink so the CLI never retains all rows or all differences in memory. Keep `compare(spec)` collecting results only as a compatibility/test wrapper; the CLI uses the streaming sink.

**Tech Stack:** Node.js 24 ESM, ExcelJS streaming reader, Ajv JSON Schema, Node.js standard library (`crypto`, `fs`, `readline`, `os`, `stream/promises`), `node:test`

---

### Task 1: Extend CompareSpec and stream workbook rows

**Files:**
- Modify: `schemas/compare-spec.schema.json`
- Modify: `src/spec.js`
- Modify: `src/read-xlsx.js`
- Modify: `test/spec.test.js`
- Modify: `test/read-xlsx.test.js`

- [ ] **Step 1: Write failing schema tests for phase-two modes and limits**

Add tests that load these mode shapes and assert defaults are applied:

```js
for (const mode of [{ type: 'row' }, { type: 'multiset' }]) {
  const loaded = await loadSpec(await writeSpec(directory, { mode }));
  assert.deepEqual(loaded.mode, mode);
  assert.deepEqual(loaded.resources, {
    maxFiles: 16,
    maxInputBytes: 10 * 1024 ** 3,
    maxRows: 10_000_000,
    maxCells: 500_000_000,
    maxTempBytes: 50 * 1024 ** 3,
    maxPartitionBytes: 64 * 1024 ** 2,
    maxRuntimeMs: 24 * 60 * 60 * 1000
  });
}
```

Also assert `key` still requires non-empty `keyColumns`, `row`/`multiset` reject `keyColumns`, every resource limit is a positive integer, and `files.length > resources.maxFiles` fails with `SPEC_INVALID` before XLSX access.

- [ ] **Step 2: Run schema tests and verify RED**

Run: `node --test test/spec.test.js`

Expected: FAIL because the schema only accepts `key` mode and has no `resources` object.

- [ ] **Step 3: Add the minimum public schema**

Use a `oneOf` for the three exact mode objects and add a defaulted `resources` object with the seven fields above. Keep `additionalProperties: false` at every object boundary. In `loadSpec()`, reject `spec.files.length > spec.resources.maxFiles` with `SpecError('file count exceeds resources.maxFiles')`.

```json
"mode": {
  "oneOf": [
    { "type": "object", "additionalProperties": false, "required": ["type", "keyColumns"], "properties": { "type": { "const": "key" }, "keyColumns": { "type": "array", "minItems": 1, "uniqueItems": true, "items": { "type": "string", "minLength": 1 } } } },
    { "type": "object", "additionalProperties": false, "required": ["type"], "properties": { "type": { "const": "row" } } },
    { "type": "object", "additionalProperties": false, "required": ["type"], "properties": { "type": { "const": "multiset" } } }
  ]
}
```

- [ ] **Step 4: Write a failing streaming scanner test**

Add a real-XLSX test for `scanFileRows(file, spec, standardColumns, onRow)` that writes 2,000 rows, records callback rows without returning a `rows` array, and checks the existing metadata contract:

```js
const seen = [];
const result = await scanFileRows(file, spec, null, async (row) => {
  seen.push([row.rowNumber, row.values['编号']]);
});
assert.equal(seen.length, 2_000);
assert.equal(result.totalRowsScanned, 2_000);
assert.equal(result.matchedRows, 2_000);
assert.equal(Object.hasOwn(result, 'rows'), false);
assert.deepEqual(result.columns, ['编号', '姓名']);
```

Retain the current `readFileRows()` tests; it becomes a small collector around `scanFileRows()` for API compatibility.

- [ ] **Step 5: Run reader tests and verify RED**

Run: `node --test test/read-xlsx.test.js`

Expected: FAIL because `scanFileRows` is not exported.

- [ ] **Step 6: Implement the streaming scanner**

Use `new ExcelJS.stream.xlsx.WorkbookReader(file.path, { worksheets: 'emit', sharedStrings: 'cache', hyperlinks: 'ignore', styles: 'ignore' })` and `for await` iteration. Resolve headers and mappings when the requested sheet reaches `headerRow`; normalize/filter each later row immediately and await `onRow(record)`. Preserve the existing `InputError` codes, alias precedence, formula/shared-formula normalization, row counts, and blank-key behavior (blank keys are invalid only in `key` mode).

Expose exactly:

```js
export async function scanFileRows(file, spec, standardColumns = null, onRow = async () => {})
export async function readFileRows(file, spec, standardColumns = null)
```

The wrapper collects rows and returns the phase-one shape; the streaming function returns `{ columns, invalidRows, totalRowsScanned, matchedRows }`.

- [ ] **Step 7: Run focused tests and commit**

Run: `node --test test/spec.test.js test/read-xlsx.test.js`

Expected: PASS.

```bash
git add schemas/compare-spec.schema.json src/spec.js src/read-xlsx.js test/spec.test.js test/read-xlsx.test.js
git commit -m "feat: stream phase two workbook rows"
```

---

### Task 2: Add bounded SHA-256 JSONL partitions

**Files:**
- Create: `src/partitions.js`
- Create: `test/partitions.test.js`

- [ ] **Step 1: Write failing partition-store tests**

Test real temporary files through this public API:

```js
const store = await PartitionStore.create({ maxOpenFiles: 2, maxTempBytes: 1024 * 1024 });
await store.append(record('alpha'));
await store.append(record('beta'));
await store.append(record('alpha'));
await store.close();
assert.deepEqual(store.partitionPaths(), [...store.partitionPaths()].sort());
assert.equal(store.bytesWritten > 0, true);
assert.equal((await readPartition(store.partitionPaths()[0])).every((item) => item.keyHash.length === 64), true);
await store.cleanup();
```

Use `record(key)` with `keyHash: sha256(JSON.stringify(key))`. Assert a tiny `maxTempBytes` throws `PartitionError` with code `TEMP_LIMIT_EXCEEDED`, no more than two cached streams remain open, and cleanup removes the run directory.

- [ ] **Step 2: Run the test and verify RED**

Run: `node --test test/partitions.test.js`

Expected: FAIL because `src/partitions.js` does not exist.

- [ ] **Step 3: Implement the partition store with standard library only**

Export:

```js
export const sha256 = (value) => createHash('sha256').update(value).digest('hex');
export class PartitionError extends Error {}
export class PartitionStore {
  static async create(options = {}) {}
  async append(record, depth = 0) {}
  async close() {}
  partitionPaths() {}
  async cleanup() {}
}
export async function* readPartition(path) {}
export async function repartition(path, depth, options) {}
```

Write compact newline-terminated JSON. Select the bucket from `record.keyHash.slice(depth * 2, depth * 2 + 2)`. Create files lazily under `mkdtemp(join(tmpdir(), 'excel-diff-'))`. Cache at most 32 append streams by default; when the limit is reached, close the least-recently-used stream before opening another. Count UTF-8 bytes before writing and fail before exceeding `maxTempBytes`. `readPartition()` uses `createReadStream` plus `readline.createInterface` and rejects malformed/blank records.

`repartition()` creates child partitions from the next hash byte, deletes the oversized parent only after all children close, and returns sorted child paths. If every record has the same full `keyHash` and the file still exceeds `maxPartitionBytes`, throw `PartitionError('HOT_KEY_TOO_LARGE', 'single key exceeds resources.maxPartitionBytes')`.

- [ ] **Step 4: Add recursive repartition tests**

Create one oversized mixed-key partition and assert `repartition()` returns more than one bounded child when hashes diverge. Create repeated records with one key and assert `HOT_KEY_TOO_LARGE`. Use byte limits below 1 KiB so the test remains fast.

- [ ] **Step 5: Run focused tests and commit**

Run: `node --test test/partitions.test.js`

Expected: PASS.

```bash
git add src/partitions.js test/partitions.test.js
git commit -m "feat: add bounded disk partitions"
```

---

### Task 3: Route key comparison through partitions and enforce resources

**Files:**
- Modify: `src/compare.js`
- Modify: `test/compare.test.js`
- Modify: `test/cli.test.js`

- [ ] **Step 1: Write failing partitioned-key tests**

Retain every phase-one key-mode assertion and add a streaming sink test:

```js
const details = { changed: [], missing: [], duplicate: [] };
const result = await comparePartitioned(spec, {
  onChanged: async (item) => details.changed.push(item),
  onMissing: async (item) => details.missing.push(item),
  onDuplicate: async (item) => details.duplicate.push(item)
});
assert.equal(Object.hasOwn(result, 'changed'), false);
assert.equal(result.summary.changedKeys, 1);
assert.equal(details.changed.length, 1);
assert.equal(result.tempDirectory, undefined);
```

Add tests setting `maxRows`, `maxCells`, `maxInputBytes`, `maxTempBytes`, and `maxRuntimeMs` below the real workload. Each must reject with a stable non-sensitive code: `ROW_LIMIT_EXCEEDED`, `CELL_LIMIT_EXCEEDED`, `INPUT_LIMIT_EXCEEDED`, `TEMP_LIMIT_EXCEEDED`, or `RUNTIME_LIMIT_EXCEEDED`. Assert temporary partitions are removed after success and failure unless `keepTemp: true`.

- [ ] **Step 2: Run focused tests and verify RED**

Run: `node --test test/compare.test.js test/cli.test.js`

Expected: FAIL because `comparePartitioned` and resource enforcement do not exist.

- [ ] **Step 3: Implement partitioned orchestration**

Expose:

```js
export async function comparePartitioned(spec, sink = {}, options = {})
export async function compare(spec)
```

The partitioned function must:

1. `stat()` all inputs and enforce `maxInputBytes` before scanning.
2. Call `scanFileRows()` once per file, reusing baseline standard columns.
3. Encode a key with the existing typed `encodeKey()`, hash it with SHA-256, encode comparable values in standard-column order, and append `{ fileId, sheetName, rowNumber, key, keyHash, rowHash, values }`.
4. Check total scanned rows, cells, elapsed monotonic time, and partition bytes after each row.
5. Close the store, process sorted partitions, recursively repartition files larger than `maxPartitionBytes`, and build only the current partition's `Map<RecordKey, Map<FileId, RowRecord[]>>`.
6. Preserve duplicate-before-missing classification, `duplicateKeyPolicy`, baseline semantics, file/column order, numeric tolerance, formula/hyperlink handling, and redacted errors.
7. Await `sink.onChanged`, `sink.onMissing`, and `sink.onDuplicate` instead of retaining details.
8. Always close streams. Remove temporary data in `finally` unless `options.keepTemp === true`; when kept, return its absolute path as `tempDirectory`.

Implement `compare(spec)` as a collector that invokes `comparePartitioned` and returns the existing `{ summary, changed, missing, duplicates }` shape so current library tests remain valid.

- [ ] **Step 4: Run all key-mode tests and commit**

Run: `node --test test/normalize.test.js test/read-xlsx.test.js test/compare.test.js test/cli.test.js`

Expected: PASS.

```bash
git add src/compare.js test/compare.test.js test/cli.test.js
git commit -m "feat: compare keys through disk partitions"
```

---

### Task 4: Add row and multiset comparison modes

**Files:**
- Modify: `src/compare.js`
- Modify: `src/report.js`
- Modify: `test/compare.test.js`
- Modify: `test/report.test.js`

- [ ] **Step 1: Write failing row-mode tests**

Use two real workbooks where the second inserts one row. Assert row mode uses the physical worksheet row as typed key evidence, compares every selected non-filter column, reports the expected downstream changes, and classifies a shorter-file tail as missing. Assert row mode never produces duplicate-key output.

```js
assert.deepEqual(result.changed[0].key, [['number', 2]]);
assert.equal(result.duplicates.length, 0);
```

- [ ] **Step 2: Write failing multiset tests**

Use three workbooks with the same full rows in different orders and assert equality. Then vary duplicate counts and assert one record:

```js
assert.deepEqual(result.multiset, [{
  values: [['string', 'Ada'], ['string', 'HR']],
  sheetName: '人员',
  counts: { A: 2, B: 1, C: 0 },
  baselineRelation: 'DELETED'
}]);
```

For a record absent from the baseline but present elsewhere, assert `baselineRelation: 'ADDED'`. Filters still apply before multiset counting.

- [ ] **Step 3: Run mode tests and verify RED**

Run: `node --test test/compare.test.js test/report.test.js`

Expected: FAIL because non-key modes have no key builder or output contract.

- [ ] **Step 4: Implement mode-specific keys and aggregation**

Use exactly these key rules:

```js
if (spec.mode.type === 'key') return spec.mode.keyColumns.map((column) => values[column]);
if (spec.mode.type === 'row') return [['number', rowNumber]];
return compareColumns.map((column) => values[column]);
```

`key` retains current semantics. `row` uses the key/row comparison path but does not exclude any compare column. `multiset` uses the full comparable typed row as the partition key, counts occurrences per file, emits nothing when every count matches, and otherwise awaits `sink.onMultiset({ values, sheetName, counts, baselineRelation })`. Add `multiset` to the `compare()` collector.

Extend `writeReport()` compatibility output with `multiset.csv` only when `spec.mode.type === 'multiset'`. Its deterministic header is `values,sheet,<fileId>.count...,baselineRelation`; values are an untyped JSON array passed through existing reversible CSV-injection protection.

- [ ] **Step 5: Run focused tests and commit**

Run: `node --test test/compare.test.js test/report.test.js`

Expected: PASS.

```bash
git add src/compare.js src/report.js test/compare.test.js test/report.test.js
git commit -m "feat: compare row and multiset modes"
```

---

### Task 5: Stream atomic reports, failure summaries, and progress events

**Files:**
- Modify: `src/report.js`
- Modify: `src/cli.js`
- Modify: `test/report.test.js`
- Modify: `test/cli.test.js`
- Modify: `README.md`

- [ ] **Step 1: Write failing atomic-report tests**

Define and test this sink API:

```js
const report = await createReportWriter(spec);
await report.onChanged(changed);
await report.onMissing(missing);
const completed = await report.complete(summary);
assert.equal(completed.summary.status, 'COMPLETED');
assert.equal((await stat(completed.directory)).isDirectory(), true);
assert.equal((await readdir(spec.output.directory)).some((name) => name.endsWith('.tmp')), false);
```

Inject a failing detail write through an optional test-only writable factory passed to `createReportWriter(spec, { createStream })`; assert `abort(error)` removes unpublished CSV staging and atomically publishes only `summary.json` with `status: 'FAILED'` and the stable error code. Do not add production test hooks to CompareSpec.

- [ ] **Step 2: Run report tests and verify RED**

Run: `node --test test/report.test.js`

Expected: FAIL because reports are written directly to their final directory and have no sink API.

- [ ] **Step 3: Implement atomic streaming reports**

Create `.<runId>.tmp` under `output.directory`, open only the detail streams needed by the selected mode, write headers immediately, and serialize each sink call with backpressure handling. `complete(summary)` closes streams, writes newline-terminated `summary.json`, then renames staging to `<runId>`. `abort(error)` closes/destroys streams, removes staging, creates a fresh failure staging directory containing only a redacted `summary.json`, and renames it atomically. Preserve `writeReport(spec, result)` as a compatibility wrapper that feeds arrays to the sink.

- [ ] **Step 4: Write failing CLI progress and keep-temp tests**

Update the CLI parser to accept:

```text
excel-diff compare --spec <path> [--progress] [--keep-temp]
```

Without `--progress`, success still writes one stdout JSON line and no stderr. With it, stderr contains newline-delimited progress objects before completion:

```js
assert.deepEqual(Object.keys(event), ['bytesWritten', 'currentFile', 'rowsScanned', 'type']);
assert.equal(event.type, 'PROGRESS');
```

Progress must be monotonic and must never include cell values or keys. `--keep-temp` includes `tempDirectory` in the final stdout object and leaves it on disk. Add a failure case proving a `FAILED` summary directory exists and no directory is labeled completed.

- [ ] **Step 5: Run CLI tests and verify RED**

Run: `node --test test/cli.test.js`

Expected: FAIL because the CLI rejects both flags and buffers the report until after comparison.

- [ ] **Step 6: Wire CLI orchestration**

Parse the two optional flags without adding a command framework. Create the report writer, pass its sink methods to `comparePartitioned`, emit throttled progress at most once per 1,000 rows plus a final event when enabled, call `complete()` on success, and call `abort()` on any error after spec loading. Keep exit codes 2/4/6 and the existing debug-stack behavior. Ensure failure stderr stays a single structured error line when progress is disabled.

- [ ] **Step 7: Update phase-two documentation**

Update `README.md` to describe all three modes, resource defaults, streaming/disk behavior, `multiset.csv`, atomic completion, failure summaries, `--progress`, `--keep-temp`, formula limitations, and the fact that Node 24 remains the supported runtime. Remove the sentence saying streaming/row/multiset are unavailable.

- [ ] **Step 8: Run full verification and commit**

Run:

```bash
npm test
git diff --check
npm pack --dry-run --cache /private/tmp/excel-diff-phase-two-npm-cache
```

Expected: every test passes, diff check is clean, and the package contains README, schema, and runtime source files without tests, docs, worktrees, or Graphify output.

```bash
git add src/report.js src/cli.js test/report.test.js test/cli.test.js README.md
git commit -m "feat: publish atomic streaming reports"
```

---

## Plan self-review

- The five tasks cover every item in technical-design section 14 phase two: streaming reader; 256 lazy partitions; partition merge and recursive repartition; duplicate/hot-key handling; resource limits; row and multiset modes; atomic output; failed summary; progress events.
- Existing normalization, filters, aliases, typed evidence, formula behavior, CSV injection protection, and phase-one library API stay in place.
- No DuckDB, UI, XLS/style comparison, Agent wrapper, inspect command, cross-platform CI, or 3-million-row benchmark is included; those remain later phases.
- Public names are consistent across tasks: `scanFileRows`, `PartitionStore`, `comparePartitioned`, `createReportWriter`, and the compatibility wrappers `readFileRows`, `compare`, `writeReport`.
