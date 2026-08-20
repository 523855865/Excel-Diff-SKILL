# Node.js 22 最低版本设计

## 目标

将项目声明和安装文档中的最低 Node.js 版本从 24 调整为 22，保持所有版本提示一致。

## 变更

- 将 `package.json` 及 `package-lock.json` 根包的 `engines.node` 改为 `>=22`。
- 将 README 中环境要求、版本检查、错误提示和版本管理器示例统一改为 Node.js 22。
- 不修改依赖、业务代码或 CI。

## 验证

- 先用现有测试证明旧声明不满足 Node.js 22 约束，再修改声明。
- 使用 Node.js 22 运行完整 `npm test`。
- 搜索并确认版本要求相关位置不再残留 Node.js 24 或 `>=24`。
