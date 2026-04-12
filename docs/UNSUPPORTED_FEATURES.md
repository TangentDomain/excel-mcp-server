# Excel MCP Server - 功能边界清单（最新）

> 最后更新: 2026-04-13  
> 测试方法: Python API 直接调用

## 📊 测试统计

- **总测试**: 47 个复杂场景
- **完全支持**: 42 个 (89%)
- **部分支持**: 2 个 (4%)
- **不支持**: 3 个 (6%)

## 🚫 完全不支持的功能 (3个)

### 1. 高级聚合 (1个)

| 功能 | 用途 | 游戏开发应用 | 优先级 |
|------|------|-------------|--------|
| `聚合的聚合` | 嵌套聚合 | 统计分析 | 🟢 低 |

**错误示例**:
```sql
-- ❌ 不支持
SELECT AVG(AVG(伤害)) FROM 技能配置
```

**错误信息**: `不支持的表达式类型: <class 'sqlglot.expressions.Avg'>`

### 2. 类型转换 (1个)

| 功能 | 用途 | 游戏开发应用 | 优先级 |
|------|------|-------------|--------|
| `CAST()` | 类型转换 | 数据格式化 | 🟡 中 |

**状态**: ⚠️ 部分支持（某些情况可能工作）

### 3. 日期/正则 (1个)

| 功能 | 用途 | 游戏开发应用 | 优先级 |
|------|------|-------------|--------|
| `CURRENT_DATE`, `RLIKE` | 日期/正则 | 时效/模式匹配 | 🟢 低 |

**错误示例**:
```sql
-- ❌ 不支持
SELECT CURRENT_DATE
SELECT 技能名称 FROM 技能配置 WHERE 技能名称 RLIKE '^火'
```

## ⚠️ 部分支持的功能 (2个)

### 1. 元组表达式 IN

**问题**: 不支持元组表达式  
**影响**: 无法使用 `WHERE (a, b) IN ((1, 2), (3, 4))`

**错误示例**:
```sql
-- ❌ 不支持
SELECT * FROM 技能配置 WHERE (技能类型, 等级) IN (('主动', 1), ('被动', 2))
```

**错误信息**: `不支持的表达式类型: <class 'sqlglot.expressions.Tuple'>`

### 2. 双行表头列名映射

**问题**: 测试中发现使用英文列名 `attack` 失败  
**影响**: 需要使用实际列名（中文或英文）

**解决方案**: 使用正确的列名

## ✅ 完全支持的功能 (42个)

### 数学函数 (6个)
- ✅ `ABS()` - 绝对值
- ✅ `CEIL()` - 向上取整
- ✅ `FLOOR()` - 向下取整
- ✅ `SQRT()` - 平方根
- ✅ `POWER()` - 幂运算
- ✅ `ROUND()` - 四舍五入

### 聚合函数 (5个)
- ✅ `COUNT()`, `SUM()`, `AVG()`, `MIN()`, `MAX()`

### 窗口函数 (10个) - 🎉 完整支持
- ✅ **基础**: `ROW_NUMBER()`, `RANK()`, `DENSE_RANK()`, `NTILE()`
- ✅ **偏移**: `LEAD()`, `LAG()` - 🆕 新发现已支持
- ✅ **首尾**: `FIRST_VALUE()`, `LAST_VALUE()` - 🆕 新发现已支持
- ✅ **范围**: `ROWS BETWEEN` - 🆕 新发现已支持
- ✅ **多列**: `PARTITION BY` 支持多列

### 字符串函数 (6个)
- ✅ `UPPER()`, `LOWER()`, `TRIM()`, `LENGTH()`
- ✅ `CONCAT()`, `REPLACE()`, `SUBSTRING()`

### 条件逻辑 (3个)
- ✅ `CASE WHEN`, `COALESCE`, `NULLIF`

### 高级功能 (8个)
- ✅ `WITH/CTE` - 公共表表达式
- ✅ 标量子查询, `IN` 子查询, `EXISTS` 子查询
- ✅ 自连接, 多表 `JOIN`, 跨文件查询
- ✅ `UNION/UNION ALL`, `INTERSECT/EXCEPT`

### 数据操作 (4个)
- ✅ `SELECT`, `WHERE`, `ORDER BY`, `GROUP BY`, `HAVING`
- ✅ `LIMIT/OFFSET`, `DISTINCT`

### 复杂场景 (9+)
- ✅ 三层嵌套子查询
- ✅ CTE + 子查询组合
- ✅ 多窗口函数同时使用
- ✅ CASE WHEN 嵌套
- ✅ 自连接
- ✅ 窗口函数 + 子查询
- ✅ 聚合 + CASE WHEN
- ✅ 数学函数组合
- ✅ 分页查询

## 🎯 支持率对比

| 版本 | 支持率 | 不支持功能数 |
|------|--------|-------------|
| 之前 | 79% | 5个 |
| **现在** | **89%** | **3个** |

## 📝 新发现支持的功能

### 窗口函数增强（已实现但未记录）

1. **LEAD/LAG** - 访问前/后行数据
   ```sql
   SELECT 技能名称, LEAD(伤害) OVER (ORDER BY 技能ID) as 下一个 FROM 技能配置
   SELECT 技能名称, LAG(伤害) OVER (ORDER BY 技能ID) as 上一个 FROM 技能配置
   ```

2. **FIRST_VALUE/LAST_VALUE** - 窗口首尾值
   ```sql
   SELECT 技能名称, FIRST_VALUE(伤害) OVER (ORDER BY 技能ID) as 首个 FROM 技能配置
   SELECT 技能名称, LAST_VALUE(伤害) OVER (ORDER BY 技能ID) as 末个 FROM 技能配置
   ```

3. **ROWS BETWEEN** - 窗口范围
   ```sql
   SELECT 技能名称, SUM(伤害) OVER (ORDER BY 技能ID ROWS BETWEEN 1 PRECEDING AND 1 FOLLOWING) as 窗口和 FROM 技能配置
   ```

## 🚀 下一步建议

### 高优先级（如果有需求）

1. **CAST 函数增强** - 完善类型转换支持
2. **元组表达式** - 支持 `WHERE (a, b) IN (...)`

### 低优先级（可选）

1. **日期函数** - `CURRENT_DATE`, `NOW()`
2. **正则表达式** - `RLIKE`, `REGEXP`
3. **聚合的聚合** - `AVG(AVG(x))`

## 📝 测试方法

参考 `.claude.md` 中的测试方法论：

```python
from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query
)

# 测试功能
result = execute_advanced_sql_query(
    '/path/to/test.xlsx',
    "SELECT LEAD(伤害) OVER (ORDER BY 技能ID) FROM 技能配置"
)

if result['success']:
    print("✅ 功能支持")
else:
    print(f"❌ 不支持: {result.get('message')}")
```

---

**维护者**: tangjian  
**测试环境**: Python 3.12, pandas 2.x, sqlglot 27.x  
**测试报告**: [超全面测试报告](testing-reports/ULTRA_COMPREHENSIVE_TEST_REPORT.md)
