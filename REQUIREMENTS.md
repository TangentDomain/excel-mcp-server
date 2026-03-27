# ExcelMCP 需求池

> **只保留OPEN/IN_PROGRESS的需求，已完成的需求见 [ARCHIVED.md](ARCHIVED.md)**

## 进化线路图（持续迭代线）

| 线路 | REQ | 关注点 |
|------|-----|--------|
| 🤖 AI体验优化 | REQ-025 | instructions/docstring/返回值统一/错误结构化/合并重复工具 |
| 📚 文档与门面 | REQ-026 | README优化/GitHub门面/使用示例/竞品对比/Changelog |
| 🎯 工具描述 | REQ-006 | AI选工具准确率 |
| 🔧 工程治理 | REQ-010 | 代码质量 |

## OPEN（待实现）

### REQ-028 [P1] FROM子查询支持
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...` 语法
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询，至少4个测试

### REQ-015 [P1] 性能优化（写入）
- **描述**：openpyxl write_only模式，减少写入内存和时间
- **验收**：大表批量写入性能提升可测量，至少3个测试

### REQ-012 [P1] 兼容性验证
- **描述**：多客户端实际测试（Cursor、Claude Desktop等）
- **验收**：至少2个主流MCP客户端验证通过

## 文档规范

| 文档 | 职责 | 谁维护 | 更新时机 |
|------|------|--------|---------|
| `REQUIREMENTS.md` | 当前活跃需求（OPEN/IN_PROGRESS） | 子代理 | 需求状态变化时 |
| `ARCHIVED.md` | 已完成/取消的需求 | 子代理 | 需求标DONE/CANCELLED时移入 |
| `.cron-focus.md` | 当前优先级和临时指令 | **仅CEO** | CEO手动更新 |
| `.cron-prompt.md` | 子代理执行规则和约束 | **仅CEO** | 规则变更时 |
| `README.md` / `README_CN.md` | 用户文档 | 子代理 | 功能变化时 |
