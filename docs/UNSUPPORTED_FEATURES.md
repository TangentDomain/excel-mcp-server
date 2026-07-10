# Excel MCP Server — 功能边界清单

> 最后更新: 2026-07-10
> 测试方法: 169 条差分测试（与 SQLite 交叉校验）+ 1447 条全量测试

## 📊 测试统计

- **差分测试**: 169 条，19 类别，100% 通过
- **全量测试**: 1447 passed, 3 skipped, 1 xfailed

## 🚫 不支持的功能

### 1. 聚合的聚合

```sql
-- ❌ 不支持
SELECT AVG(AVG(伤害)) FROM 技能配置
```

**原因**: SQL 标准不支持嵌套聚合

### 2. 日期/正则函数

```sql
-- ❌ 不支持
SELECT CURRENT_DATE
SELECT 技能名称 FROM 技能配置 WHERE 技能名称 RLIKE '^火'
```

**原因**: 非核心 SQL 标准，游戏配置表场景不需要

## ⚠️ 格式固有限制

### Excel 空字符串往返变 NULL

xlsx 格式无法区分空字符串 `""` 和 null，写入空字符串读取后变为 NULL。

**Workaround**: 使用 `IS NULL` 检查（同时匹配 NaN 和空字符串）

## ✅ 完全支持的功能

### 基础
- SELECT, DISTINCT, AS, `+-*/%`, 一元负号, 整数除法(截断向零)
- `t.*` qualified star（从子查询选择所有列）
- FROM 子查询 `FROM (SELECT ...) AS alias`

### 条件
- WHERE, LIKE, IN, NOT IN, BETWEEN, AND/OR
- 子查询（标量/IN/EXISTS）
- WHERE 引用 SELECT 别名（非窗口）— 物化为临时列
- WHERE 引用窗口函数别名 — 自动重写为子查询

### 聚合
- COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- COUNT(表达式)

### 排序
- ORDER BY, LIMIT, OFFSET, NULLS FIRST/LAST
- ORDER BY 位置编号, ORDER BY 表达式

### 窗口函数（17 个）
- ROW_NUMBER, RANK, DENSE_RANK, NTILE
- LAG, LEAD
- FIRST_VALUE, LAST_VALUE, NTH_VALUE
- AVG/SUM/MIN/MAX/COUNT OVER
- GROUP_CONCAT
- PARTITION BY（多列）, ROWS BETWEEN
- PERCENT_RANK, CUME_DIST

### 多表
- INNER/LEFT/RIGHT/FULL JOIN
- 同文件跨 Sheet JOIN
- 跨文件 JOIN（`表名@'文件路径'` 语法）
- 自连接（表别名）

### 高级
- CASE WHEN（简单/搜索/嵌套/嵌入表达式）
- CTE (WITH)
- EXISTS（关联子查询）
- UNION / UNION ALL / INTERSECT / EXCEPT
- NULLIF, COALESCE
- NULL 三值逻辑（与 SQLite 对齐）

### 字符串函数
- UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING

### 数学函数
- ABS, CEIL, FLOOR, SQRT, POWER, ROUND

### 写入
- UPDATE（含表达式计算、dry_run 预览、窗口函数）
- INSERT（单行/多行）
- DELETE（必须 WHERE）
- `_ROW_NUMBER_` 虚拟列（SELECT 和 UPDATE WHERE 均可用）

### run-python 沙箱
- 安全内置: abs/all/dict/list/len/print/range/str/class/super/异常类 等 ~60 个
- 安全模块: json/math/re/datetime/statistics/fractions 可导入
- 注入函数: query/update/insert/delete/ExcelOperations

---

**维护者**: tangjian
**测试环境**: Python 3.13, pandas 2.x, sqlglot 27.x
