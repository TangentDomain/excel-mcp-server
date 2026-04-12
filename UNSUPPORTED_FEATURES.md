# Excel MCP Server - 功能边界清单

> 最后更新: 2026-04-13  
> 测试方法: Python API 直接调用 (参考 .claude.md)

## 📊 测试统计

- **总测试**: 38 个复杂场景
- **完全支持**: 30 个 (79%)
- **部分支持**: 3 个 (8%)
- **不支持**: 5 个 (13%)

**🎉 最新更新**: 2026-04-13
- ✅ 新增支持: ABS, CEIL, FLOOR, SQRT, POWER (5个数学函数)

## 🚫 完全不支持的功能 (5个)

### 1. 数学函数 ~~(5个)~~ ✅ 已全部实现

| 函数 | 用途 | 游戏开发应用 | 优先级 | 状态 |
|------|------|-------------|--------|------|
| ~~`ABS()`~~ | 绝对值 | 属性差值计算 | 🔴 高 | ✅ 已实现 |
| ~~`CEIL()`~~ | 向上取整 | 价格向上取整 | 🔴 高 | ✅ 已实现 |
| ~~`FLOOR()`~~ | 向下取整 | 价格向下取整 | 🔴 高 | ✅ 已实现 |
| ~~`SQRT()`~~ | 平方根 | 伤害公式计算 | 🟡 中 | ✅ 已实现 |
| ~~`POWER()`~~ | 幂运算 | 指数增长 | 🟡 中 | ✅ 已实现 |

**错误示例**:
```sql
-- ❌ 不支持
SELECT ABS(Price - 1000) FROM 装备
SELECT CEIL(Price / 100) FROM 装备
SELECT SQRT(BaseAtk) FROM 装备
```

### 2. 窗口函数高级特性 (5个)

| 函数 | 用途 | 游戏开发应用 | 优先级 |
|------|------|-------------|--------|
| `LEAD()` | 取后续行 | 历史记录对比 | 🔴 高 |
| `LAG()` | 取前继行 | 历史记录对比 | 🔴 高 |
| `FIRST_VALUE()` | 取第一个值 | 首杀记录 | 🔴 高 |
| `LAST_VALUE()` | 取最后一个值 | 最新记录 | 🔴 高 |
| `PARTITION BY 多列` | 多列分区 | 复杂分组统计 | 🟡 中 |

**错误示例**:
```sql
-- ❌ 不支持
SELECT LEAD(Price, 1) OVER (ORDER BY ID) FROM 装备
SELECT ROW_NUMBER() OVER (PARTITION BY Type, Rarity ORDER BY Price) FROM 装备
```

### 3. 其他功能 (2个)

| 功能 | 用途 | 游戏开发应用 | 优先级 |
|------|------|-------------|--------|
| `CURRENT_DATE` | 当前日期 | 时效数据 | 🟡 中 |
| `RLIKE/REGEXP` | 正则表达式 | 高级模式匹配 | 🟢 低 |

## ⚠️ 部分支持的功能 (3个)

### 1. 窗口函数 ROWS BETWEEN
- **问题**: 类型比较错误 (float vs str)
- **影响**: 无法使用窗口范围

### 2. 窗口函数多层级
- **问题**: 类型比较错误
- **影响**: 无法使用多个窗口函数

### 3. 聚合的聚合
- **问题**: 不支持嵌套聚合
- **影响**: 无法计算 `AVG(AVG(x))`

### 4. CAST 函数
- **问题**: 不支持 Cast 表达式类型
- **影响**: 无法类型转换

## ✅ 完全支持的功能 (25个)

### 数学函数
- ✅ `ROUND()` - 四舍五入
- ✅ `ABS()` - 绝对值
- ✅ `CEIL()` - 向上取整
- ✅ `FLOOR()` - 向下取整
- ✅ `SQRT()` - 平方根
- ✅ `POWER()` - 幂运算

### 聚合函数
- ✅ `COUNT()`, `SUM()`, `AVG()`, `MIN()`, `MAX()`

### 窗口函数基础
- ✅ `ROW_NUMBER()`, `RANK()`, `DENSE_RANK()`, `NTILE()`

### 字符串函数
- ✅ `UPPER()`, `LOWER()`, `TRIM()`, `LENGTH()`
- ✅ `CONCAT()`, `REPLACE()`, `SUBSTRING()`

### 数据操作
- ✅ `SELECT`, `WHERE`, `ORDER BY`, `GROUP BY`, `HAVING`
- ✅ `LIMIT/OFFSET`, `DISTINCT`
- ✅ `UNION/UNION ALL`, `INTERSECT/EXCEPT`

### 条件逻辑
- ✅ `CASE WHEN`, `COALESCE`, `NULLIF`
- ✅ `BETWEEN`, `IN`, `NOT IN`, `LIKE`

### 高级功能
- ✅ `WITH/CTE` - 公共表表达式
- ✅ 标量子查询, `IN` 子查询, `EXISTS` 子查询
- ✅ 自连接, 多表 `JOIN`, 跨文件查询

## 🎯 下一步实现建议

### 高优先级 (游戏开发常用)

~~数学函数 (ABS, CEIL, FLOOR, SQRT, POWER) - ✅ 已完成~~

**窗口函数增强** (参考现有 ROW_NUMBER 实现):
1. 修复窗口函数类型比较问题
2. 实现 `LEAD/LAG/FIRST_VALUE/LAST_VALUE`
3. 支持多列 `PARTITION BY`

### 中优先级 (增强功能)

1. 修复窗口函数 ROWS BETWEEN 范围问题
2. 支持聚合的聚合 `AVG(AVG(x))`
3. 优化 CAST 函数支持

### 低优先级 (可选)

1. 日期函数 (`CURRENT_DATE`, `NOW()`)
2. 正则表达式 (`RLIKE`, `REGEXP`)

## 📝 测试方法

参考 `.claude.md` 中的测试方法论：

```python
from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query
)

# 测试新功能
result = execute_advanced_sql_query(
    '/path/to/test.xlsx',
    "SELECT ABS(Price) FROM 装备"
)

if result['success']:
    print("✅ 功能支持")
else:
    print(f"❌ 不支持: {result.get('message')}")
```

## 🔗 相关文件

- `.claude.md` - 开发指南和测试方法论
- `src/excel_mcp_server_fastmcp/api/advanced_sql_query.py` - 主要实现文件
- `tests/` - 测试用例

---

**维护者**: tangjian  
**测试环境**: Python 3.12, pandas 2.x, sqlglot 27.x
