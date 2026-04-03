# FEEDBACK.md — 跨模块反馈通道

## OPEN

### OPEN-#1: GROUP BY 聚合错误 — 部分行被归入不符合WHERE条件的分组

**来源**: CEO实测（MapEvent.xlsx）
**严重级别**: P0
**复现**:
```sql
SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt
FROM MapEvent
WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100
GROUP BY 显示路径ID, 显示位置ID
```
**预期**: 所有结果行的 显示路径ID∈{1,2} 且 显示位置ID<100
**实际**: 出现 [36, 569, 58] — 显示路径ID=36 不在IN(1,2)中，显示位置ID=569 不<100
**说明**: 58行数据被错误地聚合到一个完全不符合WHERE条件的分组。GROUP BY的分组键计算存在bug，导致部分行的值被错误映射。同样在步骤2(位置分布)中出现不存在的路径3。
**文件**: `src/api/advanced_sql_query.py` — GROUP BY逻辑
**状态**: 已转REQ（REQ-052）第268轮

## CLOSED

## #1 excel-mcp-server - 发现217个函数缺少Args/Parameters段
- **严重程度**：高
- **工具**：docstring_analysis
- **状态**：已转REQ（REQ-049）第263轮

## #2 excel-mcp-server - ExcelOperations类API方法缺失
- **严重程度**：高
- **工具**：api_consistency_check
- **状态**：已转REQ（误报，方法名不同但功能完整）第263轮

## #3 excel-mcp-server - 文档完整性严重不达标
- **严重程度**：高
- **工具**：documentation_quality_assessment
- **状态**：已转REQ（REQ-049）第263轮
