# Excel Diff（第一阶段）

使用 Node.js 24 比较两份或多份 Excel `.xlsx` 文件。先安装依赖：

```sh
npm ci
node src/cli.js compare --spec /absolute/path/to/compare.json
```

也可通过安装后的命令运行：`excel-diff compare --spec /absolute/path/to/compare.json`。

## CompareSpec

路径相对于 JSON 文件所在目录解析。第一阶段只支持业务键（`key`）模式，完整的两文件示例如下：

```json
{
  "version": "1.0",
  "baseline": "before",
  "files": [
    { "id": "before", "path": "before.xlsx" },
    { "id": "after", "path": "after.xlsx" }
  ],
  "sheet": { "name": "人员", "headerRow": 1 },
  "mode": { "type": "key", "keyColumns": ["员工编号"] },
  "compareColumns": "*",
  "duplicateKeyPolicy": "report",
  "output": { "directory": "output" }
}
```

严格默认值为：`output.sampleSize` 为 `20`；`filters` 与 `columnAliases` 为空；标准化为 `emptyEqualsNull: false`、`caseSensitive: true`、`formulaMode: "formula-and-cached-result"`。公式只读取工作簿中已有的公式和缓存值，**不会重算公式**；超长数字应以文本单元格保存，以免 Excel 的数值精度限制改变业务键或比较值。

成功时标准输出仅有一行 JSON，包含汇总字段和本次绝对输出目录 `directory`。该目录包含：

- `summary.json`：本次汇总；
- `changed.csv`：字段差异；值单元格是 typed tuple 的 JSON，例如 `["number",12]`；
- `missing.csv`：新增、删除的业务键。

为防止 CSV 公式注入，危险的标签值会编码为 `json:<JSON string>`。

## 范围与验证

第一阶段会将工作表载入内存，不提供 streaming、按行模式、multiset 比较或 Agent 包装。

```sh
npm test
```

| 退出码 | 含义 |
| --- | --- |
| 0 | 比较完成 |
| 2 | 命令用法或 CompareSpec 无效 |
| 4 | 输入文件或比较错误 |
| 6 | 未预期内部错误 |

失败时标准错误仅有一行 JSON：`{"status":"FAILED","code":"...","message":"..."}`。默认不输出堆栈；设置 `EXCEL_DIFF_DEBUG=1` 时会在该 JSON 中附加 `stack`。
