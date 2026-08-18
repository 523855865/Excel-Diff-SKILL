# CSV 值与类型分列设计

## 目标

让 `changed.csv` 和 `missing.csv` 更适合人工阅读，同时保留 Excel Diff 内部的带类型比较语义。

## 输出契约

`changed.csv` 的固定列仍为 `key,sheet,column`。每个输入文件按 CompareSpec 中的顺序输出三列：

```text
<fileId>.value,<fileId>.type,<fileId>.row
```

例如两文件比较的表头为：

```text
key,sheet,column,before.value,before.type,before.row,after.value,after.type,after.row
```

- `.value`：只展示规范值的内容，不再展示 `[type,value]` 包装。
- `.type`：展示该文件中单元格的类型，例如 `string`、`number`、`date`、`blank`、`formula` 或 `hyperlink`。
- `.row`：保持原始 Excel 行号。

简单值直接写入 CSV。公式、超链接等复合值写为不含类型包装的 JSON；公式保留公式文本和缓存结果，超链接保留显示文本、目标地址和 tooltip。

## key 展示

- 单字段 key：直接展示不含类型包装的值，例如 `001`。
- 复合 key：展示不含类型包装的 JSON 数组，例如 `["001","CN"]`。
- `changed.csv` 与 `missing.csv` 使用同一规则。

内部仍使用带类型 key 做聚合和比较；这里只改变报告展示，因此字符串 `"1"` 与数字 `1` 仍不会被错误匹配。

## 安全与确定性

key 和 `.value` 拆除 JSON typed tuple 后，字符串可能以 `= + - @ TAB CR LF` 开头。所有可由工作簿或 CompareSpec 控制的 CSV 文本继续经过现有可逆 `json:<JSON string>` 编码，避免电子表格公式注入。

文件顺序、差异排序、CSV 转义、运行目录和 `summary.json` 均保持不变。

## 测试

- 锁定新的动态表头顺序：`value,type,row`。
- 验证字符串、数字、空值、公式和超链接的值与类型分列。
- 验证单字段和复合 key 在两个 CSV 中均不含 typed tuple。
- 验证拆分后的 key/value 仍执行 CSV 注入防护。
- 完整测试继续覆盖多文件顺序和空报告表头。

## 非目标

- 不改变规范化、过滤、主键匹配或差异分类。
- 不同时保留旧 typed tuple CSV 格式。
- 不新增配置项、兼容模式或新的输出文件。
