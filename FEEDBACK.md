# FEEDBACK.md

## OPEN-#1 [P0] GROUP BY 聚合逻辑 bug

**状态**：✅ 已转REQ-061（2026-04-05完成）
**紧急度**：CEO 明确要求优先修复，不能绕过

### 约束
验证规则已写在 REQUIREMENTS.md REQ-052 的 notes 中。子代理处理此 REQ 时必须遵守：修复前后都跑验证代码，只有输出 FIXED 才能标 DONE。

### 修复方向
- 文件：`src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`
- 方法：`_apply_group_by_aggregation`
- 不要再修数据加载、中文替换、WHERE 逻辑——这些都没问题
- 必须直接修 `_apply_group_by_aggregation` 方法
