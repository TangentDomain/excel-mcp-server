# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/ROADMAP.md)，已归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求

### REQ-025 [P1] AI体验优化线（持续迭代，不关闭）
- **关注点**：instructions优化（已完成）、docstring优化（持续）、返回值统一（进行中）、错误信息结构化、大结果截断（已完成）、合并重复工具（preview/assess已完成，get_headers待合并）

### REQ-026 [P1] 文档与门面优化线（持续迭代，不关闭）
- **关注点**：README 30秒上手教程、GitHub门面、使用示例、竞品对比、Changelog

- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过

- **描述**：openpyxl write_only模式，减少写入内存和时间
- **完成**：v1.5.3，所有修改操作支持streaming参数，copy-modify-write方案

### REQ-012 [P1] 兼容性验证
- **描述**：多客户端实际测试（Cursor、Claude Desktop等）

### REQ-006 [P1] 工具描述持续优化（持续迭代，不关闭）
### REQ-010 [P1] 工程治理（持续迭代，不关闭）
