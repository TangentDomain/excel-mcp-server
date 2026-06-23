# 量化指标体系 (Quantification Standard)

> 不量化就无法比较，不比较就无法进步。
> 所有功能增减、缺陷修复都必须以量化指标的变化来证明价值。

## 指标维度

每个维度都是 `covered / total`，产出百分比。汇总成总评分。

### M1: 准确率 (Accuracy) — 核心

差分测试：ExcelMCP 结果 vs SQLite 真值。

| 子维度 | 说明 |
|--------|------|
| UPDATE 准确率 | 写入后读回值与 SQLite 一致的比例 |
| INSERT 准确率 | 插入后 COUNT(*) 两边一致 |
| DELETE 准确率 | 删除后行数两边一致 |
| WHERE 准确率 | 各 WHERE 条件过滤结果与 SQLite 一致 |
| 聚合准确率 | COUNT/SUM/AVG/MAX/MIN 与 SQLite 一致 |

**公式**: `准确率 = pass_count / total_cases`

### M2: 功能覆盖率 (Feature Coverage)

26 个 CLI 工具中多少被测试覆盖。

| 子维度 | 目标 |
|--------|------|
| 查询类 (9工具) | 每个工具至少 1 个测试 |
| 写入类 (7工具) | 每个工具至少 1 个读写往返测试 |
| 结构类 (6工具) | 每个工具至少 1 个测试 |
| 格式化类 (2工具) | 每个工具至少 1 个测试 |
| 文件类 (2工具) | 每个工具至少 1 个测试 |

**公式**: `覆盖率 = tested_tools / 26`

### M3: SQL 特性覆盖率 (SQL Feature Coverage)

| 特性 | 判定方式 |
|------|---------|
| SELECT * | 能返回所有列 |
| WHERE (=, >, <, <>, !=) | 条件过滤准确 |
| WHERE IN | 集合匹配准确 |
| WHERE LIKE | 模式匹配准确 |
| WHERE BETWEEN | 范围匹配准确 |
| WHERE AND/OR/NOT | 逻辑组合准确 |
| ORDER BY | 排序准确（含 ASC/DESC） |
| LIMIT/OFFSET | 分页准确 |
| COUNT/SUM/AVG/MAX/MIN | 聚合值与 SQLite 一致 |
| GROUP BY | 分组准确 |
| HAVING | 分组后过滤准确 |
| JOIN (INNER/LEFT/RIGHT/FULL) | 关联准确 |
| 子查询 | 嵌套查询准确 |
| UNION/UNION ALL | 合并准确 |
| CASE WHEN | 条件表达式准确 |
| DISTINCT | 去重准确 |
| 窗口函数 | ROW_NUMBER/RANK 准确 |
| 字符串函数 | UPPER/LOWER/LENGTH/CONCAT |
| CTE (WITH) | 公用表表达式准确 |

**公式**: `SQL覆盖率 = tested_features / 19`

### M4: 边界值覆盖率 (Edge Value Coverage)

| 边界值 | 说明 |
|--------|------|
| 整数 | 正常整数读写往返 |
| 浮点数 | 含小数精度 |
| 大整数 | > 2^31 |
| 负数 | 负值读写 |
| 零 | 0 值 |
| NULL | None 写入读出 |
| 空字符串 | '' 写入读出 |
| 中文 | 多字节字符 |
| 特殊字符 | O'Brien, 引号, 分号 |
| 极小浮点 | 0.0001 |

**公式**: `边界覆盖率 = tested_edges / 10`

### M5: 写操作安全性 (Write Safety)

| 检查项 | 判定 |
|--------|------|
| affected_rows 精确 | == SQLite rowcount |
| 文件完整性 | 非目标 sheet 不变 |
| 无匹配安全 | WHERE 无匹配时 affected_rows=0 |
| 失败安全 | 失败时 data=[], 无堆栈 |
| 幂等读取 | 同一 SELECT 两次结果一致 |
| 公式列守恒 | 公式列不被写操作覆盖 |

**公式**: `安全性 = passed_checks / 6`

### M6: 性能 (Performance)

| 指标 | 目标 |
|------|------|
| SELECT 平均延迟 | < 100ms (小文件) |
| UPDATE 平均延迟 | < 200ms |
| JOIN 平均延迟 | < 300ms |
| 大文件查询 | < 1000ms (1000+ 行) |

**公式**: `性能达标率 = met_thresholds / 4`

## 总评分

```
总评分 = (M1×0.40 + M2×0.15 + M3×0.15 + M4×0.10 + M5×0.15 + M6×0.05) × 100
```

权重理由：准确率最重要(40%)，功能/SQL覆盖各15%，写安全15%，边界值10%，性能5%。

## 趋势追踪

每次对抗运行后，一行 JSON 追加到 `data/adversarial-score.jsonl`：

```json
{
  "timestamp": "2026-06-23T13:00:00",
  "version": "1.17.0",
  "M1_accuracy": {"overall": 0.85, "update": 0.90, "insert": 1.0, "delete": 1.0, "where": 0.80, "aggregate": 0.95},
  "M2_feature_coverage": {"overall": 0.65, "tested": 17, "total": 26},
  "M3_sql_coverage": {"overall": 0.74, "tested": 14, "total": 19},
  "M4_edge_coverage": {"overall": 0.60, "tested": 6, "total": 10},
  "M5_write_safety": {"overall": 0.83, "passed": 5, "total": 6},
  "M6_performance": {"overall": 0.75, "met": 3, "total": 4},
  "total_score": 79.5,
  "failures": [{"sql": "...", "category": "where", "expected": ..., "actual": ...}]
}
```

## 迭代决策标准

- **修复优先级** = 影响的指标权重 × (当前值到 1.0 的差距)
- **功能增减判定**：新增功能必须提升某个指标 ≥ 1%；否则不加
- **退化判定**：任何指标下降 > 2% 即为退化，必须修复
