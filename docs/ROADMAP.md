# ROADMAP — 路线图

> CEO维护，子代理参考。每个Phase有完成标志，达标即进入下一阶段。

## Phase 1：核心能力 ✅
- SQL引擎（45项功能）、安全加固、calamine性能、跨文件JOIN
- 完成标志：100/100连续10轮 → ✅ 连续18轮

## Phase 2：体验打磨（当前）
- ✅ 返回值统一（_wrap统一结构 + success/data/meta/error_code）
- ✅ 错误结构化（StructuredSQLError + error_code + hint + suggested_fix）
- ✅ 合并重复工具（preview/assess合并、get_headers合并）
- ✅ FROM子查询支持
- ✅ 文档门面（README优化、30秒上手教程、CHANGELOG）
- ✅ 写入性能优化（write_only流式写入，create/import/merge已完成）
- ✅ 多客户端兼容性验证（Cursor、Claude Desktop、VSCode MCP、OpenAI ChatGPT Plugin 100%通过）
- ✅ write_only覆盖修改操作（excel_write_only_override，流式+openpyxl双模式）
- 完成标志：AI选工具准确率>95% + 多客户端验证通过 ✅

## Phase 3：生态扩展（规划）
- 配置模板系统、VSCode插件、社区建设
- 完成标志：GitHub star > 100 + 至少1个社区贡献
