# FEEDBACK.md

## {编号} DOCSTRING-001 docstring_quality_assessment - 文档完整性严重不达标
- **严重程度**：高
- **工具**：docstring_quality_assessment
- **参数**：{"total_functions": 506, "missing_args_sections": 541, "compliance_rate": "-6.9%"}
- **期望**：所有公共函数都应有完整的Args/Parameters和Returns文档段，文档覆盖率达到90%以上
- **实际**：506个函数中发现541个文档问题，合规率仅-6.9%（负数表示问题数量超过函数总数）
- **修复建议**：批量修复所有函数的docstring，确保包含Args/Parameters和Returns段，建立自动化docstring检查机制，设定90%以上合规率目标
- **状态**：✅ 已转REQ-062（2026-04-05创建）

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

### FEEDBACK-013: 主会话 sed 修复引入的函数签名回归（P0）
- **问题**：主会话用 `sed -i 's/, df -> str:$/df: pd.DataFrame) -> str:/g'` 批量修复12处 `df -> pd.DataFrame:` 语法错误时，对部分函数签名引入了回归：
  1. `_apply_order_by(self, parsed_sql, df)` — 参数名被改成 `df: pd.DataFrame)` 但方法签名定义不同
  2. `_apply_join_clause(self, joins, left_df)` — 同理，参数被错误替换
  3. `_evaluate_case_expression` 中 `row` 变量未定义
- **根因**：sed 替换太暴力，没有区分不同函数签名的差异
- **要求**：逐个检查被我 sed 替换过的函数签名，对照 git diff 恢复正确签名。测试用例：
  - `SELECT name FROM employees ORDER BY salary DESC`
  - `SELECT name, CASE WHEN salary > 30000 THEN '高' ELSE '低' END FROM employees`
  - `SELECT e.name FROM employees e JOIN orders o ON e.name = o.customer`
- **状态**：P0
