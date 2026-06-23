# Harness 载体索引

> 本文件列出 excel-mcp-server 所有 Harness 载体及其职责。

## 载体清单

| # | 载体 | 路径 | 层级 | 职责 |
|---|------|------|------|------|
| 1 | 身份声明 | `AGENTS.md` | L0 | 根级身份、Turn 流程、项目事实、记忆索引 |
| 2 | 项目指令 | `CLAUDE.md` | L0 | AI 编码指令（若存在） |
| 3 | 公理 | `.omp/rules/axioms.md` | L3 | 不可违反的硬约束 |
| 4 | 质量门禁 | `.omp/rules/quality-gates.md` | L4 | 质量门禁规则（docstring 等） |
| 5 | 自我状态 | `.omp/memory/self-state.md` | L0 | 我是谁、任务路由、当前状态 |
| 6 | 决策记录 | `.omp/memory/decisions.md` | L1 | 架构决策日志（D编号） |
| 7 | 陷阱记录 | `.omp/memory/pitfalls.md` | L1 | 已知陷阱（P编号） |
| 8 | 用户偏好 | `.omp/memory/preferences.md` | L1 | 用户偏好 |
| 9 | 关系图谱 | `.omp/memory/relationships.md` | L1 | 依赖关系图 |
| 10 | 不变量体系 | `.omp/memory/invariants.md` | L3 | INV-1~32 + AX 映射 |
| 11 | 判据检查 | `.omp/extensions/criterion-check.ts` | L2 | PostToolUse 判据 hook |
| 12 | pre-commit | `.githooks/pre-commit` | L2 | Git pre-commit 入口 |
| 13 | 门禁逻辑 | `tools/git-pre-commit-gate.js` | L2 | pre-commit 检查实现 |
| 14 | hooks 配置 | `tools/setup-hooks.js` | L2 | 自动配置 core.hooksPath |
| 15 | L5 自举检查 | `tools/check-harness-l5.js` | L5 | 进化引擎健康检查 |
| 16 | 进化引擎 | `tools/harness-l5/evolve.js` | L5 | L5 进化主循环 |
| 17 | 受保护评估 | `tools/harness-l5/protected-evaluator.js` | L5 | 受保护文件评估器 |
| 18 | 载体注册表 | `data/carrier-registry.json` | L5 | 载体注册表 |
| 19 | 评估器注册表 | `data/evaluator-registry.json` | L5 | 评估器注册表 |
| 20 | 进化日志 | `data/evolution-log.json` | L5 | 进化历史日志 |
| 21 | 信任锚 | `data/trust-anchor.sha256` | L5 | SHA-256 信任锚 |
| 22 | Ruff 判据 | `scripts/ruff-pass.sh` | L1 | Loop I1 ruff 检查 |
| 23 | 不变量判据 | `scripts/invariant-pass.sh` | L1 | Loop I1 不变量检查 |
| 24 | 输出验证 | `scripts/verify-output.py` | L2 | Loop I2 独立验证器 |
| 25 | 外层对比 | `scripts/compare-holdout.ts` | L3 | Loop I3 holdout 对比 + 回滚 |
