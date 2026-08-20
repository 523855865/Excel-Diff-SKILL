# Excel Diff

## Skill 的作用

Excel Diff Skill 用于检查和比较两份或多份 Excel `.xlsx` 文件。Agent 先检查工作表、表头、类型和候选业务键，向用户确认工作表、业务键、列映射等歧义，再由 CLI 按确认后的 CompareSpec 确定性读取、比较并生成结构化结果与完整报告。

Skill 不执行宏、公式或外部链接，也不会重算公式；Agent 不应根据抽样结果擅自推断业务规则。

## 环境搭建

Agent 必须依次完成：**检查环境 → 搭建缺失环境 → 安装并验证 CLI → 安装并验证 Codex Skill 和 Claude Code 插件**。CLI 尚不可运行时，不要先安装 Skill。

### 1. 检查环境

需要 Node.js 24 或更高版本以及 npm。macOS/Linux：

```sh
node --version
npm --version
node -e 'const major = Number(process.versions.node.split(".")[0]); if (major < 24) { console.error("Node.js 24+ is required"); process.exit(1); }'
```

Windows PowerShell：

```powershell
node --version
npm --version
node -e 'const major = Number(process.versions.node.split(".")[0]); if (major < 24) { console.error("Node.js 24+ is required"); process.exit(1); }'
```

任一命令不存在或版本检查失败时，先搭建环境。Agent 在安装系统软件前应先取得用户许可，并优先复用机器已有的 Node.js 版本管理器：

| 已有工具 | 安装并启用 Node.js 24 |
| --- | --- |
| `mise` | `mise use --global node@24` |
| `fnm` | `fnm install 24`，然后执行 `fnm use 24` |
| `nvm` / nvm-windows | `nvm install 24`，然后执行 `nvm use 24` |

没有版本管理器时，使用 [Node.js 官方安装方式](https://nodejs.org/en/download)。安装后重新执行本节的全部检查，不要绕过 Node.js 版本门禁。

### 2. 安装并验证 CLI

在仓库根目录执行：

```sh
npm ci
npm test
npm install --global .
```

不要使用 `sudo npm install --global .`；如果全局 npm 目录不可写，应改用用户级 Node.js 版本管理器或用户级 npm prefix。验证命令已进入 PATH：

macOS/Linux：

```sh
command -v excel-diff
```

Windows PowerShell：

```powershell
Get-Command excel-diff
```

### 3. 安装并验证 Codex Skill

在本仓库或其子目录中启动 Codex 时，Codex 会从 `.agents/skills/excel-diff` 自动发现仓库级 Skill，无需复制文件。

需要在其他仓库中使用时，可将 Skill 安装到用户级 `$HOME/.agents/skills`。先检查目标位置；如果已存在同名 Skill，停止并让用户决定保留、升级或移除，不要直接覆盖。

macOS/Linux：

```sh
skill_source="$(pwd)/.agents/skills/excel-diff"
skill_target="${HOME}/.agents/skills/excel-diff"
test ! -e "$skill_target" || { echo "Skill already exists: $skill_target"; exit 1; }
mkdir -p "${HOME}/.agents/skills"
ln -s "$skill_source" "$skill_target"
test -f "$skill_target/SKILL.md"
```

Windows PowerShell：

```powershell
$skillSource = (Resolve-Path ".agents\skills\excel-diff").Path
$skillRoot = Join-Path $HOME ".agents\skills"
$skillTarget = Join-Path $skillRoot "excel-diff"
if (Test-Path $skillTarget) { throw "Skill already exists: $skillTarget" }
New-Item -ItemType Directory -Force -Path $skillRoot | Out-Null
New-Item -ItemType Junction -Path $skillTarget -Target $skillSource | Out-Null
if (-not (Test-Path (Join-Path $skillTarget "SKILL.md"))) { throw "Skill installation failed" }
```

Codex 通常会自动检测新增 Skill；如果未出现，重启 Codex。使用 `/skills` 查看，或在提示中输入 `$excel-diff` 显式调用。

### 4. 安装并验证 Claude Code 插件

临时加载本地插件并验证：

```sh
claude --plugin-dir ./plugins/excel-diff
claude plugin validate ./plugins/excel-diff
```

也可在 Claude Code 内添加本地仓库路径或远程仓库 URL，再安装插件（本仓库不声明已发布到远程 marketplace）：

```text
/plugin marketplace add <repo-path-or-url>
/plugin install excel-diff@excel-diff-tools
```

安装完成后，让 Claude Code 显示已加载的插件或显式调用 `excel-diff` Skill；若插件未出现，重启 Claude Code 后再检查。

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

### 属性说明

CompareSpec 只允许下表及其子字段；未在 Schema 中声明的字段会导致 `SPEC_INVALID`。`files[].path` 和 `output.directory` 均相对于 `compare.json` 所在目录解析。

| 字段 | 必填 | 类型/取值 | 作用 |
| --- | --- | --- | --- |
| `version` | 是 | 固定为 `"1.0"` | CompareSpec 版本。 |
| `baseline` | 是 | 非空字符串 | 基准文件的 `files[].id`；新增、删除和值变化均相对它判断。 |
| `files` | 是 | 至少两个对象 | 参与比较的 XLSX 文件；`id` 和解析后的 `path` 都必须唯一。 |
| `sheet` | 是 | 对象 | 指定所有文件要读取的工作表和表头行。 |
| `mode` | 是 | `key`、`row` 或 `multiset` | 指定记录如何对齐。 |
| `compareColumns` | 是 | `"*"` 或非空列名数组 | 指定要比较的列。 |
| `columnAliases` | 否 | 对象，默认 `{}` | 把不同文件中的别名表头映射到统一列名。 |
| `filters` | 否 | 数组，默认 `[]` | 比较前过滤数据行；多个过滤条件之间为 AND。 |
| `normalization` | 否 | 对象 | 配置全局和逐列的值归一化、公式读取方式及数值容差。 |
| `duplicateKeyPolicy` | 是 | `report` 或 `fail` | `key` 模式发现重复业务键时记录报告或立即失败；Schema 在所有模式下都要求此字段。 |
| `resources` | 否 | 对象 | 限制输入、扫描、临时磁盘、分区和运行时间；省略时使用默认值。 |
| `output` | 是 | 对象 | 设置结果根目录和交付时可读取的明细样本数。 |

`files` 与 `sheet`：

| 字段 | 必填 | 类型/约束 | 作用 |
| --- | --- | --- | --- |
| `files[].id` | 是 | 非空字符串，唯一 | 文件标识；供 `baseline`、报告列名和错误信息引用。 |
| `files[].path` | 是 | 非空字符串，扩展名 `.xlsx` | 输入文件路径；相对路径基于 CompareSpec 目录解析。 |
| `sheet.name` | 是 | 非空字符串 | 所有输入文件中要读取的工作表名称。 |
| `sheet.headerRow` | 是 | 大于等于 1 的整数 | 表头所在的物理行号；下一行开始作为数据行扫描。 |

`mode` 与比较列：

| 字段/取值 | 作用 |
| --- | --- |
| `mode.type: "key"` | 用 `keyColumns` 的有类型值组成业务键，对齐不同文件中的记录。空业务键行记为无效行。 |
| `mode.keyColumns` | `key` 模式必填的非空、无重复列名数组；键列不会再作为值列比较。 |
| `mode.type: "row"` | 用工作表物理行号对齐记录，适合行顺序稳定的文件。 |
| `mode.type: "multiset"` | 忽略行顺序，按所有选中列的有类型值组合统计各文件出现次数。 |
| `compareColumns: "*"` | 使用基准文件的全部表头；后续文件必须能通过原名、NFKC 规范化名称或别名映射到这些列。 |
| `compareColumns: ["列A", "列B"]` | 只比较列出的列；数组不能为空且不能包含重复列名。 |

`columnAliases` 的键是统一列名，值是该列可接受的其他表头名称。例如 `{"姓名":["Name","员工姓名"]}`。程序优先使用表头原名，其次使用 NFKC 规范化名称，最后使用别名；一对多或多对一映射无法唯一确定时返回 `COLUMN_MAPPING_AMBIGUOUS`。

过滤条件对象包含以下字段：

| 字段 | 必填 | 作用 |
| --- | --- | --- |
| `filters[].column` | 是 | 要过滤的统一列名；该列必须存在于每个输入文件。 |
| `filters[].operator` | 是 | 过滤操作符，见下表。 |
| `filters[].value` | 按操作符 | 单个标量；`in`、`notIn`、`between` 也允许用它承载数组。 |
| `filters[].values` | 按操作符 | `in`、`notIn`、`between` 可使用的数组形式；不能与 `value` 同时出现。 |

| 操作符 | 值要求 | 含义 |
| --- | --- | --- |
| `eq` / `ne` | 一个标量 `value` | 等于 / 不等于。 |
| `gt` / `gte` / `lt` / `lte` | 一个标量 `value` | 同类型数字、字符串或日期的大于 / 大于等于 / 小于 / 小于等于。 |
| `in` / `notIn` | `value` 或 `values` 中恰好一个非空数组 | 在 / 不在给定集合中。 |
| `contains` / `startsWith` / `endsWith` | 一个字符串 `value` | 字符串包含 / 以其开头 / 以其结尾。 |
| `isNull` / `isNotNull` | 不允许 `value` 或 `values` | 为空 / 不为空。 |
| `between` | `value` 或 `values` 中恰好一个二元素数组 | 闭区间判断，包含上下界。 |

过滤字面量只能是 `null`、字符串、数字或布尔值。日期使用 `YYYY-MM-DD` 或带时区的 ISO 8601 字符串；过滤在键有效性检查和比较之前执行。

归一化配置：

| 字段 | 默认值 | 作用 |
| --- | --- | --- |
| `normalization.emptyEqualsNull` | `false` | 为 `true` 时把空字符串 `""` 与空单元格统一为 blank。 |
| `normalization.caseSensitive` | `true` | 字符串比较是否区分大小写。 |
| `normalization.formulaMode` | `"formula-and-cached-result"` | `formula` 只比较公式文本；`cached-result` 只比较工作簿缓存值；默认值同时比较两者。公式不会重算。 |
| `normalization.columns` | `{}` | 以统一列名为键，覆盖该列的归一化规则。 |
| `normalization.columns.<列>.trim` | 未设置 | 为 `true` 时去除该列字符串首尾空白。 |
| `normalization.columns.<列>.caseSensitive` | 继承全局值 | 覆盖该列的大小写规则。 |
| `normalization.columns.<列>.emptyEqualsNull` | 继承全局值 | 覆盖该列的空值规则。 |
| `normalization.columns.<列>.numericTolerance` | 未设置 | 非负数；比较该列两个数字是否相等时允许的最大绝对误差。 |

资源与输出配置：

| 字段 | 默认值 | 作用 |
| --- | ---: | --- |
| `resources.maxFiles` | 16 | 允许参与比较的最大文件数。 |
| `resources.maxInputBytes` | 10,737,418,240 | 所有输入 XLSX 文件大小之和上限。 |
| `resources.maxRows` | 10,000,000 | 所有文件扫描物理行数之和上限，包括被过滤和无效的行。 |
| `resources.maxCells` | 500,000,000 | 所有扫描行的单元格计数之和上限。 |
| `resources.maxTempBytes` | 53,687,091,200 | 共享字符串 spool、比较分区和重分区 staging 的临时磁盘总上限。 |
| `resources.maxPartitionBytes` | 67,108,864 | 单个内存加载分区的目标上限；无法继续拆分的超大单键会失败。 |
| `resources.maxRuntimeMs` | 86,400,000 | 整次比较允许的最长运行时间，单位毫秒。 |
| `output.directory` | 无，必填 | 结果根目录；相对路径基于 CompareSpec 目录解析，且不能等于输入文件路径。 |
| `output.sampleSize` | 20 | Agent 交付结果时最多读取的明细记录数；`0` 表示不读取明细示例，不限制产出的完整 CSV。 |

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

## Skill 维护

Skill 的唯一维护源是 `skill/SKILL.md`。修改后同步并检查 Codex 与 Claude Code 的副本：

```sh
node scripts/sync-skills.mjs
node scripts/sync-skills.mjs --check
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
