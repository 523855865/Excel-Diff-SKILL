---
name: excel-diff
description: Use when comparing two or more XLSX workbooks, inspecting sheet or header differences, selecting comparison keys, or explaining Excel Diff results and structured error codes.
---

# Excel Diff

Inspect first, resolve ambiguity with the user, then compare from a validated CompareSpec.

## Workflow

1. Confirm at least two `.xlsx` paths and the target sheet. Never infer a sheet.
2. Always inspect before comparing:

   ```sh
   excel-diff inspect --files "/abs/before.xlsx" "/abs/after.xlsx" --sheet "人员"
   # In this repository:
   node src/cli.js inspect --files "/abs/before.xlsx" "/abs/after.xlsx" --sheet "人员"
   ```

   Add `--full-types` only when complete-column type counts are necessary. Otherwise inspection samples the first 10,000 data rows.
3. Review every file's sheet visibility, raw and NFKC-normalized headers, type distribution, empty ratio, sampled key candidates and duplicate rates, plus formula, merged-cell, and duplicate-header risks.
4. Handle the structured result before continuing:
   - `INSPECTED`: propose the sheet, key, and column mapping, then ask for confirmation wherever intent remains ambiguous.
   - `NEEDS_INPUT`: explain its code and ask for the missing decision. Do not compare yet.
   - `FAILED`: quote the code and message; do not claim success. For `SHEET_NOT_FOUND`, present the sheets listed in the error and ask the user to choose.
5. Do not treat sampled key candidates or observed field changes as reliable business conclusions. If the sheet, key, column mapping, or duplicate header is ambiguous, wait for explicit confirmation before comparing.
6. Read the [bundled CompareSpec schema](references/compare-spec.schema.json), then express the confirmed rules using only its supported fields. Run:

   ```sh
   excel-diff compare --spec "/abs/compare.json"
   ```

## CompareSpec field reference

Unknown fields are rejected. `files[].path` and `output.directory` are resolved relative to the CompareSpec file.

| Field | Required | Meaning |
| --- | --- | --- |
| `version` | yes | Must be `"1.0"`. |
| `baseline` | yes | The `files[].id` used as the reference for added, deleted, and changed results. |
| `files` | yes | At least two `{id,path}` entries. IDs and resolved paths must be unique; paths must end in `.xlsx`. |
| `sheet` | yes | `{name,headerRow}` selects one worksheet and its 1-based physical header row in every file. |
| `mode` | yes | `{type:"key",keyColumns:[...]}`, `{type:"row"}`, or `{type:"multiset"}`. |
| `compareColumns` | yes | `"*"` for every baseline header, or a non-empty unique list of column names. |
| `columnAliases` | no | `{canonicalName:[alternateHeader,...]}`; exact and NFKC header matches take precedence, and ambiguous mappings fail. Default `{}`. |
| `filters` | no | Row filters applied with AND before key validation and comparison. Default `[]`. |
| `normalization` | no | Global value rules plus per-column overrides. |
| `duplicateKeyPolicy` | yes | `report` writes duplicate-key details and skips comparison for those keys; `fail` stops with `DUPLICATE_KEY`. Required in every mode. |
| `resources` | no | Positive integer resource limits; omitted values receive schema defaults. |
| `output` | yes | `{directory,sampleSize?}` for the report root and delivery sample limit. |

Nested fields and semantics:

| Field | Meaning |
| --- | --- |
| `files[].id` | Non-empty file identifier used by `baseline`, reports, and errors. |
| `files[].path` | Input XLSX path. |
| `sheet.name` | Worksheet name shared by all inputs. |
| `sheet.headerRow` | Header row number; following physical rows are data. |
| `mode.keyColumns` | Required only for `key`; typed business-key columns, excluded from value comparison. Blank keys are invalid rows. |
| `filters[].column` | Canonical column name, required in every input. |
| `filters[].operator` | `eq`, `ne`, `gt`, `gte`, `lt`, `lte`, `in`, `notIn`, `contains`, `startsWith`, `endsWith`, `isNull`, `isNotNull`, or `between`. |
| `filters[].value` | Scalar operand; for `in`, `notIn`, and `between`, it may instead contain the required array. |
| `filters[].values` | Alternative array operand for `in`, `notIn`, and `between`; never combine it with `value`. |
| `normalization.emptyEqualsNull` | Default `false`; when true, empty string and blank cells normalize to blank. |
| `normalization.caseSensitive` | Default `true`; controls string case matching. |
| `normalization.formulaMode` | `formula`, `cached-result`, or default `formula-and-cached-result`; formulas are never recalculated. |
| `normalization.columns.<name>.trim` | Trim leading and trailing string whitespace for one canonical column. |
| `normalization.columns.<name>.caseSensitive` | Override global case handling for one column. |
| `normalization.columns.<name>.emptyEqualsNull` | Override global blank handling for one column. |
| `normalization.columns.<name>.numericTolerance` | Non-negative maximum absolute difference when testing two numeric values for equality. |
| `resources.maxFiles` | Maximum input file count; default `16`. |
| `resources.maxInputBytes` | Maximum combined input size; default `10737418240`. |
| `resources.maxRows` | Maximum scanned physical rows across all files; default `10000000`. |
| `resources.maxCells` | Maximum summed scanned cell count; default `500000000`. |
| `resources.maxTempBytes` | Maximum combined temporary spool, partition, and repartition-staging bytes; default `53687091200`. |
| `resources.maxPartitionBytes` | Target maximum partition loaded in memory; default `67108864`. |
| `resources.maxRuntimeMs` | Maximum comparison duration in milliseconds; default `86400000`. |
| `output.directory` | Report root directory; must not resolve to an input file. |
| `output.sampleSize` | Default `20`; maximum detail records the Agent may read when delivering results, not a CSV output limit. |

Filter operand rules: `eq`, `ne`, `gt`, `gte`, `lt`, `lte`, `contains`, `startsWith`, and `endsWith` require one scalar `value`; `in` and `notIn` require one non-empty array in exactly one of `value` or `values`; `between` requires exactly two values in one of them; `isNull` and `isNotNull` accept neither. Literals are null, string, number, or boolean. Date literals use `YYYY-MM-DD` or timezone-qualified ISO 8601. Ordered comparisons require matching number, string, or date types, and `between` is inclusive.

Mode semantics: `key` aligns rows by typed business keys; `row` aligns by physical row number; `multiset` ignores row order and compares per-file counts of typed selected-value combinations. With `compareColumns: "*"`, every later file must map every baseline header.

## Safety and reporting

- Never use `eval`, `Function`, temporary JavaScript, or arbitrary SQL to execute user-supplied conditions.
- Never execute workbook macros, formulas, or external links; inspect formula metadata and cached values only.
- After a successful comparison, cite `summary.json` and the absolute artifact directory. Read at most `output.sampleSize` detail records across the detail artifacts when giving examples; leave complete results in the artifacts.
- On failure, cite the structured error code and do not say the comparison completed.
- In every delivery state that formulas are not recalculated, precision already lost from overlong Excel numeric cells cannot be recovered, sampled key candidates are provisional, and full duplicate-key validation occurs during comparison.
