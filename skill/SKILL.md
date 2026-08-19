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
6. Express the confirmed rules only with fields supported by `schemas/compare-spec.schema.json`. Run:

   ```sh
   excel-diff compare --spec "/abs/compare.json"
   ```

## Safety and reporting

- Never use `eval`, `Function`, temporary JavaScript, or arbitrary SQL to execute user-supplied conditions.
- After a successful comparison, cite `summary.json` and the absolute artifact directory. Read at most `output.sampleSize` detail records across the detail artifacts when giving examples; leave complete results in the artifacts.
- On failure, cite the structured error code and do not say the comparison completed.
- In every delivery state that formulas are not recalculated, precision already lost from overlong Excel numeric cells cannot be recovered, sampled key candidates are provisional, and full duplicate-key validation occurs during comparison.
