# SQL 标准遵循原则

## 核心原则

**只支持 SQL 标准支持的功能。SQLite 3.x 为唯一真值来源。**

### 为什么？

1. **与主流数据库保持一致** — MySQL, PostgreSQL, SQL Server, Oracle 都遵循 SQL 标准
2. **避免自创特殊行为** — 不实现 SQL 标准之外的特殊逻辑
3. **降低维护成本** — 符合标准的行为更容易维护和文档化
4. **用户可预期** — 用户已有的 SQL 知识可以直接应用

## 应用规范

### 功能开发
- ✅ **支持** — 符合 SQL 标准的功能（如窗口函数、JOIN、聚合）
- ❌ **不支持** — SQL 标准明确不支持的功能
- 💡 **透明重写** — 对 SQL 标准允许但引擎实现受限的写法，自动重写为等价形式

### 测试用例设计
- ✅ **测试用例必须符合 SQL 标准** — 不测试 SQL 不支持的写法
- ❌ **不测试边缘 bug** — SQL 标准限制不是 bug
- ✅ **差分测试** — 与 SQLite 交叉校验，169 条测试 100% 通过

## SQL 执行顺序

```
1. FROM     → 确定数据源
2. JOIN     → 表关联
3. WHERE    → 筛选行
4. GROUP BY → 分组
5. HAVING   → 分组筛选
6. 窗口函数 → 计算窗口函数
7. SELECT   → 选择列
8. ORDER BY → 排序
9. LIMIT    → 限制结果
```

## 引擎透明重写

### WHERE 引用窗口函数别名

SQL 标准中 WHERE 在窗口函数之前执行，不能直接引用窗口别名。但用户常写：

```sql
SELECT *, RANK() OVER(...) AS rk FROM t WHERE rk <= 3
```

引擎自动重写为标准等价形式（透明，用户无感）：

```sql
SELECT * FROM (SELECT *, RANK() OVER(...) AS rk FROM t) _sub WHERE rk <= 3
```

### WHERE 引用 SELECT 别名（非窗口）

SQLite 支持在 WHERE 中引用 SELECT 别名。引擎通过物化别名为临时列实现：

```sql
SELECT a+b AS sum_val FROM t WHERE sum_val > 10
-- 引擎内部：先计算 a+b 为临时列，再 WHERE 过滤
```

## 已知限制

### 不可修复（SQL/格式固有限制）

- **Excel 空字符串往返变 NULL** — xlsx 格式无法区分空字符串和 null
- **聚合的聚合** — `AVG(AVG(x))` 不支持（SQL 标准也不支持嵌套聚合）

### 已修复（原为限制，现已支持）

- ~~FROM 子查询~~ → 已支持
- ~~SELECT t.* 限定星号~~ → 已支持
- ~~WHERE 引用窗口函数别名~~ → 透明重写为子查询
- ~~WHERE 引用 SELECT 别名~~ → 物化为临时列

---

**真值来源**: SQLite 3.x
**校验方式**: calibrator 交叉校验 + 169 条差分测试
**最后更新**: 2026-07-10
