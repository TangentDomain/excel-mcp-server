# FEEDBACK.md

## OPEN-#1 [P0] GROUP BY 聚合错误（精确线索）

**状态**：待执行
**来源**：CEO 实测复现 + 主会话调试

<<<<<<< Updated upstream
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
=======
### 精确复现
```sql
SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID
```
- 文件：`/tmp/MapEvent.xlsx`，sheet `MapEvent`，58行数据
- 双行表头已修复：列名现在是 `TriggerPoint_PathID`（英文），dtype=uint8，unique=[1, 2]
- 中文替换已修复：`显示路径ID` → `TriggerPoint_PathID`
- **传入 `_execute_query` 的 DataFrame 完全正确**：58行，PathID 只有 1 和 2
- **但 GROUP BY 结果出现了 [38, 589, 58]**（这些值不存在于原始数据中）

### 关键发现
1. ✅ 数据加载正确（双行表头检测已修复）
2. ✅ 中文列名替换正确（`显示路径ID` → `TriggerPoint_PathID`）
3. ✅ 传入 `_execute_query` 的 DataFrame 正确
4. ❌ **bug 在 `_execute_query` 内部的 GROUP BY 实现逻辑**
5. 手动 pandas groupby 结果正确（30行，全部符合 WHERE 条件）

### 修复方向
- 检查 `_apply_group_by_aggregation` 方法
- 可能的 bug 点：列名映射（`_original_to_clean_cols`）干扰、聚合时用了错误的 DataFrame、pandas groupby 后的列名还原逻辑
- **不要**再查 WHERE 逻辑、数据加载逻辑、中文替换逻辑——这些都没问题

### 测试验证
```python
# 验证修复的代码
result = engine.execute_sql_query('/tmp/MapEvent.xlsx', 'SELECT 显示路径ID, 显示位置ID, COUNT(*) as cnt FROM MapEvent WHERE 显示路径ID IN (1, 2) AND 显示位置ID < 100 GROUP BY 显示路径ID, 显示位置ID')
bad = [r for r in result['data'][1:] if r[0] not in [1,2] or r[1] >= 100]
assert len(bad) == 0, f"BUG: {bad}"
```

### 已修复的相关问题（本轮）
1. 双行表头检测正则：`^[a-zA-Z_]` → `^[a-zA-Z_.#]`（允许 `.` 和 `#`）
2. 非空阈值：`>= 3` → `>= 2`
3. desc_map 用清洗后的列名（`_clean_dataframe` 之后构建）
4. 缓存 key 加入 sheet_name：`file_path` → `file_path|sheet_name`
>>>>>>> Stashed changes
