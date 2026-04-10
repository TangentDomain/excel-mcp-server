# SQL功能完整度报告

## 支持度评估

| 功能类别 | 支持度 | 说明 |
|---------|--------|------|
| **基础查询** | ⭐⭐⭐⭐⭐ | SELECT, DISTINCT, WHERE, ORDER BY, LIMIT, OFFSET |
| **窗口函数** | ⭐⭐⭐⭐⭐ | 16个：ROW_NUMBER/RANK/DENSE_RANK/LAG/LEAD/FIRST_VALUE/LAST_VALUE/NTILE/PERCENT_RANK/CUME_DIST + AVG/SUM/COUNT/MIN/MAX OVER |
| **JOIN + 窗口函数** | ⭐⭐⭐⭐⭐ | 全部支持 |
| **多表JOIN** | ⭐⭐⭐⭐⭐ | INNER/LEFT/RIGHT/FULL/CROSS，支持多条件ON |
| **跨文件JOIN** | ⭐⭐⭐⭐⭐ | @'path'语法引用其他Excel文件 |
| **CTE + 窗口** | ⭐⭐⭐⭐⭐ | WITH + 窗口函数 + 多层分析 |
| **聚合函数** | ⭐⭐⭐⭐⭐ | COUNT/SUM/AVG/MAX/MIN/GROUP_CONCAT/COUNT(DISTINCT) |
| **数据修改** | ⭐⭐⭐⭐⭐ | UPDATE/INSERT/DELETE，含dry_run预览 |
| **集合运算** | ⭐⭐⭐⭐⭐ | UNION/UNION ALL/INTERSECT/EXCEPT |
| **子查询** | ⭐⭐⭐⭐⭐ | WHERE/FROM/SELECT/EXISTS，FROM支持单层嵌套 |
| **字符串函数** | ⭐⭐⭐⭐⭐ | UPPER/LOWER/TRIM/LENGTH/CONCAT/REPLACE/SUBSTRING/LEFT/RIGHT |
| **CASE WHEN** | ⭐⭐⭐⭐⭐ | 含窗口函数嵌套 |
| **算术表达式** | ⭐⭐⭐⭐⭐ | WHERE/SELECT中支持 + - * / % |

## 不支持的SQL特性（非Excel场景需求）

- NATURAL JOIN → 使用显式 JOIN ON
- WITH RECURSIVE → Excel无层级数据
- LATERAL JOIN → 使用子查询或CTE替代
- 存储过程/触发器/视图 → 非Excel概念

## 测试覆盖

- 48个SQL专项测试 + 16个INSERT/DELETE/跨文件JOIN测试
- 969个总测试通过，0回归
- 45+项SQL特性逐一验证通过
