# FEEDBACK.md

## OPEN-#1 [P0] GROUP BY 聚合逻辑 bug（第10次尝试）

**状态**：待执行
**紧急度**：CEO 明确要求优先修复，不能绕过

### ！！铁律！！
**REQ-052 禁止标记为 DONE 除非通过验证！** 验证方法写在 REQUIREMENTS.md 的 notes 中。上游问题已修复，BUG 在 _apply_group_by_aggregation 方法内部。

### 精确线索
- 文件：`src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`
- 方法：`_apply_group_by_aggregation`
- 复现 SQL：`SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID`
- 文件：`/tmp/MapEvent.xlsx`，sheet `MapEvent`
- 传入 _execute_query 的 DataFrame 完全正确：58行，TriggerPoint_PathID unique=[1,2]，dtype=uint8
- 手动 pandas groupby 结果正确（30行，全部符合 WHERE）
- 但 _apply_group_by_aggregation 返回了 [38, 589, 58] 等不存在于数据中的值
- **不要**再修数据加载、中文替换、WHERE 逻辑——这些都没问题
- **必须**直接修 _apply_group_by_aggregation 方法
