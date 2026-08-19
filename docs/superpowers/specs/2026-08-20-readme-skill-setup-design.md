# README Skill 环境与安装指引设计

## 目标

让 Agent 阅读 README 后，先理解 Excel Diff Skill 的职责，再按固定顺序完成环境检查、缺失环境搭建、CLI 安装、Skill 安装和验证。

## 文档结构

仅修改 `README.md`，在现有命令说明之前增加以下内容：

1. **Skill 作用**：说明 Agent 负责检查工作簿、识别并确认工作表、业务键和列映射等歧义；CLI 负责确定性读取、比较和生成结构化报告。
2. **Agent 首次使用顺序**：明确要求依次执行“检查环境 → 搭建缺失环境 → 安装并验证 CLI → 安装并验证 Skill”，不得在 CLI 尚不可运行时直接安装 Skill。
3. **环境检查与搭建**：检查 Git、Node.js 24 或更高版本、npm；环境不满足时优先复用已有的 `mise`、`fnm` 或 `nvm`，否则引导使用 Node.js 官方安装方式。分别给出 macOS/Linux 与 Windows PowerShell 检查命令。
4. **CLI 安装与验证**：运行 `npm ci`、`npm test` 和 `npm install --global .`，再用系统命令确认 `excel-diff` 已进入 PATH。
5. **Codex Skill 安装**：说明仓库内运行时可由 `.agents/skills/excel-diff` 自动发现；跨仓库使用时安装到 `$HOME/.agents/skills`，并在未出现时重启 Codex。
6. **Claude Code 插件安装**：保留本地 `--plugin-dir` 方式和 marketplace 方式，并把它们放在 CLI 可运行之后。

## 边界

- 不新增安装脚本、依赖、配置或自动下载命令。
- 不改变 CLI、CompareSpec、比较语义或现有 Skill 内容。
- 不声称仓库已发布到远程 marketplace。
- 不覆盖用户已有的同名 Skill；安装命令必须允许 Agent 先检查目标位置。

## 验证

- README 中出现完整且唯一的首次使用顺序。
- README 命令与 `package.json`、`.agents/skills`、`.claude-plugin/marketplace.json` 和 `plugins/excel-diff` 的实际结构一致。
- 运行 `npm test`、`node scripts/sync-skills.mjs --check` 和 `git diff --check`。
