# Excel Diff

要求 Node.js 24，用于流式检查和比较两份或多份 Excel `.xlsx` 文件。

macOS/Linux：

```sh
npm ci
npm install --global .
```

Windows PowerShell：

```powershell
npm ci
npm install --global .
```

## Inspect

先检查工作表、表头、类型分布、空值率、候选键及公式、合并单元格和重复表头风险：

```sh
excel-diff inspect --files /absolute/path/before.xlsx /absolute/path/after.xlsx --sheet 人员
```

默认对前 10,000 个数据行统计类型；需要扫描全部数据行的类型时增加 `--full-types`。发现重复表头时，命令输出 `NEEDS_INPUT`，应先确认列映射再比较。

Windows PowerShell 调用示例：

```powershell
excel-diff inspect --files "C:\work\before.xlsx" "C:\work\after.xlsx" --sheet "人员" --full-types
excel-diff compare --spec "C:\work\compare.json"
```

## CompareSpec

```sh
excel-diff compare --spec /absolute/path/to/compare.json [--progress] [--keep-temp]
```

`compare` 之后的 `--spec`、`--progress` 和 `--keep-temp` 可以任意顺序出现。

路径相对于 JSON 文件所在目录解析。完整的两文件 `key` 模式示例如下：

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

支持三种模式：

- `{"type":"key","keyColumns":["员工编号"]}`：按有类型的业务键比较，键列不再作为值列比较；
- `{"type":"row"}`：按物理行号比较所有选中列；
- `{"type":"multiset"}`：忽略行顺序，按所有选中列的有类型值组合统计每个文件的出现次数。

严格默认值为：`output.sampleSize` 为 `20`；`filters` 与 `columnAliases` 为空；标准化为 `emptyEqualsNull: false`、`caseSensitive: true`、`formulaMode: "formula-and-cached-result"`。资源限制默认值为：

| 字段 | 默认值 |
| --- | ---: |
| `maxFiles` | 16 |
| `maxInputBytes` | 10,737,418,240 |
| `maxRows` | 10,000,000 |
| `maxCells` | 500,000,000 |
| `maxTempBytes` | 53,687,091,200 |
| `maxPartitionBytes` | 67,108,864 |
| `maxRuntimeMs` | 86,400,000 |

公式只读取工作簿中已有的公式和缓存值，**不会重算公式**；若缓存值过期，比较结果也会反映该过期值。超长数字应以文本单元格保存，以免 Excel 的数值精度限制改变业务键或比较值。

## 输出与进度

成功时标准输出仅有一行 JSON，包含汇总字段和本次绝对输出目录 `directory`。比较过程逐行读取 XLSX，将共享字符串和紧凑比较记录写入临时磁盘；内存中只加载当前有界分区以及具有 Excel 固定上限的元数据（样式约 65k、每个工作表的超链接约 65k）。共享字符串 spool 和比较分区共同计入 `maxTempBytes`。报告先写入输出目录下的隐藏 `.tmp` staging 目录，完成后才原子重命名为最终目录。`output.directory` 必须是单写者目录；Node.js 标准库不提供 POSIX 目录 `rename-noreplace`，因此外部进程若在最终冲突检查与 `rename` 之间抢占同一个随机 run ID，不在保证范围内。最终目录包含：

- `summary.json`：本次汇总；
- `key`/`row` 模式的 `changed.csv`、`missing.csv` 和 `duplicate-keys.csv`：字段差异、新增/删除记录以及重复键所在文件；
- `multiset` 模式的 `multiset.csv`：值组合、sheet、按 CompareSpec 文件顺序排列的 `<fileId>.count` 和 `baselineRelation`。

为防止 CSV 公式注入，危险的标签、key 和值会编码为 `json:<JSON string>`。

`--progress` 在标准错误输出 NDJSON 进度，最多每 1,000 个扫描行一次并包含最终进度；事件只有 `bytesWritten`、`currentFile`、`rowsScanned`、`type` 四个非敏感字段。`--keep-temp` 在成功 stdout JSON 中增加绝对 `tempDirectory` 并保留临时分区；失败时 stderr JSON 和失败 `summary.json` 也会返回该路径。使用者在成功或失败后都负责清理。

## 验证与错误

```sh
npm test
```

## Agent 集成

Codex 在仓库内通过 `.agents/skills/excel-diff/SKILL.md` 发现此 Skill；唯一维护源是 `skill/SKILL.md`。修改后同步并检查副本：

```sh
node scripts/sync-skills.mjs
node scripts/sync-skills.mjs --check
```

Claude Code 开发时可直接加载本地插件：

```sh
claude --plugin-dir ./plugins/excel-diff
claude plugin validate ./plugins/excel-diff
```

也可在 Claude Code 内添加本地仓库路径或远程仓库 URL，再安装插件（本仓库不声明已发布到远程 marketplace）：

```text
/plugin marketplace add <repo-path-or-url>
/plugin install excel-diff@excel-diff-tools
```

| 退出码 | 含义 |
| --- | --- |
| 0 | 检查或比较完成 |
| 2 | 命令用法或 CompareSpec 无效 |
| 3 | 检查发现歧义，需要输入（`NEEDS_INPUT`） |
| 4 | 输入文件或比较错误 |
| 5 | 资源限制超出，或磁盘空间/配额耗尽（`DISK_FULL`） |
| 6 | 未预期内部错误 |

只要输出文件系统仍可写，失败时未发布的 CSV staging 会被删除，并原子发布一个仅含脱敏 `summary.json` 的失败目录。若输出目录一开始就无法创建，物理上无法写入失败报告；若文件系统在清理或发布中变为不可写，程序会尽力清除明细，可能只留下空 staging 目录，并以脱敏 `INTERNAL_ERROR` 报告。未启用 progress 时，标准错误仅有一行 JSON：`{"status":"FAILED","code":"...","message":"..."}`。默认不输出堆栈；设置 `EXCEL_DIFF_DEBUG=1` 时会在该 stderr JSON 中附加 `stack`。
