# Node.js 22 Minimum Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将项目最低 Node.js 版本及安装文档统一调整为 22。

**Architecture:** 版本契约由 `package.json` 定义，`package-lock.json` 保存根包副本，README 提供相同的人工安装门禁。用现有 `node:test` 分发测试锁定三处一致性。

**Tech Stack:** Node.js ESM、`node:test`、npm package metadata、Markdown

---

### Task 1: 锁定并更新最低 Node.js 版本

**Files:**
- Modify: `test/distribution.test.js`
- Modify: `package.json:16-18`
- Modify: `package-lock.json:18-20`
- Modify: `README.md:15-39`

- [ ] **Step 1: 写失败测试**

在 `test/distribution.test.js` 增加测试，读取 `package.json`、`package-lock.json` 和 README，断言根包引擎均为 `>=22`，README 不再包含 Node.js 24 门禁：

```js
test('package metadata and README require Node.js 22 or newer', async () => {
  const [pkg, lock, readme] = await Promise.all([
    json('package.json'),
    json('package-lock.json'),
    readFile(join(root, 'README.md'), 'utf8')
  ]);

  assert.equal(pkg.engines.node, '>=22');
  assert.equal(lock.packages[''].engines.node, '>=22');
  assert.match(readme, /需要 Node\.js 22 或更高版本/);
  assert.doesNotMatch(readme, /Node\.js 24|node@24|install 24|use 24|major < 24/);
});
```

- [ ] **Step 2: 验证测试按预期失败**

Run: `node --test test/distribution.test.js`

Expected: FAIL，实际值仍为 `>=24`。

- [ ] **Step 3: 最小实现**

- 将 `package.json` 和 `package-lock.json` 根包引擎改为 `"node": ">=22"`。
- 将 README 第 1 节中的版本说明、两个检查命令、表头及 mise/fnm/nvm 示例从 24 改为 22。

- [ ] **Step 4: 验证局部测试通过**

Run: `node --test test/distribution.test.js`

Expected: 3 tests pass，0 fail。

- [ ] **Step 5: 用 Node.js 22 完整验证**

Run: `node --version && npm test && ! rg -n 'Node\\.js 24|node@24|install 24|use 24|major < 24|\">=24\"' README.md package.json package-lock.json`

Expected: Node.js `v22.17.0`，169 tests pass，搜索无结果。

- [ ] **Step 6: 提交实现**

```bash
git add README.md package.json package-lock.json test/distribution.test.js
git commit -m "docs: support Node.js 22 and newer"
```
