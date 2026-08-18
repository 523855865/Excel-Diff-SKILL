# Excel Diff 阶段一设计

## 目标

交付一个可从命令行运行的小文件 XLSX 对比闭环，用来验证 CompareSpec、严格比较语义和结果契约。阶段一只支持业务主键模式，所有参与比较的工作簿可整体放入内存。

## 范围

支持：

- 两个及以上 XLSX 文件。
- 一个明确指定的工作表和表头行。
- 显式基准文件与单列或复合业务主键。
- `compareColumns: "*"` 或显式比较列。
- 文档列出的过滤操作符。
- 严格默认值及列级 `trim`、大小写、空值和数值误差规则。
- 重复主键报告并跳过普通字段比较。
- `summary.json`、`changed.csv`、`missing.csv`。

不支持：

- 流式读取与磁盘分区。
- `row`、`multiset` 模式。
- `inspect`、列语义猜测和交互式歧义处理。
- Agent Skill、Claude Code 插件和发布包装。
- 容量基准、Web UI、样式和宏对比。

## 技术选择

- Node.js 24，JavaScript ESM。
- ExcelJS 读取 XLSX。
- Ajv 校验 JSON Schema。
- Node.js 内置 `node:test` 测试。
- 命令行参数、哈希、文件系统和 CSV 输出均使用 Node.js 标准库。

阶段一不引入命令行、CSV、日志或测试框架。代码保持单包，不建立 workspace、接口层或可替换后端。

## 组件

### CompareSpec Schema

`schemas/compare-spec.schema.json` 定义阶段一允许的输入，并设置 `additionalProperties: false`。CLI 在读取 XLSX 前完成校验。阶段一要求：

- `version` 为 `1.0`。
- `baseline` 必须引用唯一的文件 ID。
- `files` 至少两个，文件 ID 与路径不得重复。
- `sheet.name`、`sheet.headerRow` 明确给出。
- `mode.type` 只能是 `key`，且 `keyColumns` 非空。
- `duplicateKeyPolicy` 只能是 `report` 或 `fail`。
- 输出目录必须明确给出。

JSON Schema 负责结构校验；跨字段约束由一段小型语义校验完成。

### CLI

阶段一提供一个命令：

```bash
node src/cli.js compare --spec compare-spec.json
```

标准输出只写最终摘要 JSON。诊断信息写标准错误。退出码沿用技术文档：成功为 `0`，Spec 无效为 `2`，输入或 XLSX 错误为 `4`，内部错误为 `6`。

### XLSX 读取与列映射

ExcelJS 读取指定工作表。指定表头行转换为列名到列号的映射，并拒绝：

- 重复表头。
- 缺失主键列。
- 缺失显式比较列或过滤列。
- 不同文件间无法建立一致标准列集合。

数据行保留源文件 ID、工作表名和原始行号，以便差异可追溯。

### 规范化与过滤

每个单元格转换为带类型的规范值。默认不裁剪、不忽略大小写、不混淆空字符串、空单元格和 `null`，也不把字符串数字转换成数字。

公式默认同时比较公式文本和文件内已有的缓存结果；ExcelJS 不负责重新计算公式。

过滤发生在主键生成前。主键为空或无法解析的行计入 `invalidRows`；阶段一不额外输出 `invalid-rows.csv`。

复合主键使用带类型值数组的规范 JSON 编码，不使用分隔符拼接。

### 内存聚合与比较

读取结果聚合为：

```text
Map<RecordKey, Map<FileId, RowRecord[]>>
```

每个主键按以下顺序分类：

1. 任一文件出现重复主键：`DUPLICATE_KEY`；`fail` 立即失败，`report` 计数后跳过普通比较。
2. 仅部分文件存在：`MISSING_IN_FILES`，并计算相对基准的 `ADDED` 或 `DELETED`。
3. 所有文件都存在且字段一致：`IDENTICAL`。
4. 所有文件都存在且字段不同：`CHANGED`，逐字段写明细。

字段比较先处理数值误差，再比较规范编码。行顺序不参与 `key` 模式结果。

### 输出

每次运行在配置目录下创建独立 `run-id` 子目录并写出：

- `summary.json`：状态、计数、输入文件数和明细路径。
- `changed.csv`：一行代表一个主键和字段差异，包含工作表、各文件的值和原行号。
- `missing.csv`：一行代表一个缺失主键，包含存在文件、缺失文件和基准关系。

CSV 使用 UTF-8 和 RFC 4180 风格转义。阶段一不实现原子输出与失败摘要；这两项随阶段二的流式写入一起完成。

## 错误处理

- 所有 Spec 错误必须在读取第一个工作簿前返回。
- 不允许静默选择重复表头、重复文件 ID、重复路径或重复主键。
- 输入路径必须是可读 `.xlsx` 文件。
- 输出目录解析后不得等于或覆盖任一输入文件。
- Excel 公式只读取文件内已有的公式文本和缓存结果，不重新计算。
- 日志和错误默认不打印完整单元格值。

## 测试

使用一个 `node:test` 集成测试文件动态生成小型 XLSX，不提交二进制 fixture。覆盖：

- Spec 结构与跨字段校验。
- 两文件完全一致且行顺序不同。
- 复合主键无拼接碰撞。
- 字段变化及 CSV 转义。
- 相对基准的新增和删除。
- 重复主键的 `report` 与 `fail`。
- 严格空值、字符串数字、trim、大小写和数值误差。
- 缺失 Sheet、重复表头和缺失主键列。

验收命令：

```bash
npm test
node src/cli.js compare --spec <test-spec.json>
```

## 完成标准

- 全新安装后 `npm test` 通过。
- 示例 XLSX 可通过 CLI 生成三种结果文件。
- 行顺序变化不影响主键模式结果。
- 每条字段差异包含文件、工作表、原行号、主键和字段。
- 非法 Spec、重复主键和列结构问题均以明确错误退出。
- 实现没有阶段二及以后才需要的抽象或依赖。
