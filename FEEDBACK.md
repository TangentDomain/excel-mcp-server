# FEEDBACK.md — 跨模块反馈通道

## OPEN

### OPEN-#1: GROUP BY bug 精确线索 — 待修复

**严重级别**: P0

**复现**:
```sql
SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt
FROM MapEvent
WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100
GROUP BY 显示路径ID, 显示位置ID
```

**预期**: 所有结果行的 显示路径ID∈{1,2} 且 显示位置ID<100  
**实际**: 出现 [38, 589, 58] — 显示路径ID=38 不在IN(1,2)中，显示位置ID=589 不<100
**说明**: 58行数据被错误地聚合到一个完全不符合WHERE条件的分组。GROUP BY的分组键计算存在bug，导致部分行的值被错误映射。

**数据文件**: `/tmp/MapEvent.xlsx`
**状态**: 已修复（REQ-052）第276轮

**精确根因线索（CEO已验证）**:
1. SQL: `SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID`
2. 结果最后一行 `[38, 589, 58]` — 58行数据被错误聚合
3. `original_rows: 379` 但 MapEvent sheet 实际只有59行（58数据+1表头）
4. ExcelMCP 把所有9个sheet（共379行）的数据混在一起了
5. 其他sheet有 显示路径ID=38, 显示位置ID=589 的数据（共58条）
6. **bug在数据加载阶段**：SQL执行时没有正确限定只读指定sheet的数据
7. 数据类型已确认全部是int（58行+1行str残留的表头行），不是类型问题

**修复方向**: `src/api/advanced_sql_query.py` — 检查 FROM 子句解析后的数据加载逻辑，确保只加载 sheet_name 对应的 DataFrame，不要把所有sheet合并

**文件**: `src/api/advanced_sql_query.py` — GROUP BY逻辑

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
