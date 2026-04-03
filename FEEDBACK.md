# FEEDBACK.md — 跨模块反馈通道

## OPEN

### OPEN-#1: GROUP BY bug 精确线索 — 待修复

**严重级别**: P0
**数据文件**: `/tmp/MapEvent.xlsx`
**状态**: 待执行（REQ-052 就是这个 bug）

**精确根因线索（CEO已验证）**:
1. SQL: `SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID`
2. 结果最后一行 `[38, 589, 58]` — 58行数据被错误聚合
3. `original_rows: 379` 但 MapEvent sheet 实际只有59行（58数据+1表头）
4. ExcelMCP 把所有9个sheet（共379行）的数据混在一起了
5. 其他sheet有 显示路径ID=38, 显示位置ID=589 的数据（共58条）
6. **bug在数据加载阶段**：SQL执行时没有正确限定只读指定sheet的数据
7. 数据类型已确认全部是int（58行+1行str残留的表头行），不是类型问题

**修复方向**: `src/api/advanced_sql_query.py` — 检查 FROM 子句解析后的数据加载逻辑，确保只加载 sheet_name 对应的 DataFrame，不要把所有sheet合并
