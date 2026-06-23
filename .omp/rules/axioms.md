---
description: ExcelMCP 顶层公理体系——从 4 条公理推导 32 条不变量
alwaysApply: true
globs: ["src/**/*.py", "tests/**/*.py"]
version: "1.0.0"
---

# 公理体系（AX）

> 从 4 条顶层公理推导出 32 条不变量（INV-1~32）。冲突时按优先级裁决。

## AX-001: SQL 标准是唯一真值来源 [优先级: 1]

所有 SQL 行为以 SQLite 3.x 为准。calibrator 交叉校验保证对齐。

**推导**: INV-2（SQL-SQLite 对齐）、INV-10~15（窗口函数/排名/空表/特殊字符/除零/LIKE）、INV-25~32（DISTINCT/HAVING/NULL/子查询/OFFSET/NOT/双行表头/_ROW_NUMBER_）

**冲突裁决**: 当 SQL 标准与 Excel 特性冲突时，SQL 标准优先，Excel 特性作为已知限制记录。

## AX-002: 写操作必须安全可回滚 [优先级: 2]

任何写操作（UPDATE/INSERT/DELETE）不能破坏文件完整性。写操作失败时必须失败安全（不部分写入）。

**推导**: INV-3（文件完整）、INV-16~18（UPDATE/INSERT/DELETE 验证）、INV-19（写操作 SQLite 对齐）、INV-20（公式列守恒）、INV-22~24（affected_rows/无匹配安全/NULL写入）

**冲突裁决**: 写操作失败时必须失败安全。

## AX-003: 失败必须安全且可诊断 [优先级: 3]

所有错误返回结构化结果，不泄露堆栈，错误可分类。

**推导**: INV-1（结果结构）、INV-5（失败安全）、INV-6（错误可分类）

**冲突裁决**: 诊断信息优先于原始异常。

## AX-004: 查询语义必须确定且幂等 [优先级: 4]

读取操作不修改状态，相同输入得相同输出。

**推导**: INV-4（行数守恒）、INV-7（幂等读取）、INV-8（LIMIT 约束）、INV-9（聚合语义）、INV-21（跨文件 JOIN 真值）

**冲突裁决**: 幂等性优先于性能优化。
