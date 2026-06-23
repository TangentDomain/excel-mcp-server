# 决策记录

> 所有设计/架构决策的不可变日志。一旦记录不修改，只追加。

## D001: FastMCP 选型

- **背景**: 需要一个 MCP 服务器框架来暴露 Excel 操作工具
- **决策**: 选择 FastMCP（Python）作为 MCP 服务器框架
- **原因**: FastMCP 是 Python MCP 生态的主流选择，天然集成 asyncio，装饰器注册工具简洁
- **影响**: 整个项目基于 FastMCP 的工具注册模式，Python 生态

## D002: SQLite 真值来源

- **背景**: Excel 不保证 SQL 语义正确性，需要一个权威参照
- **决策**: 以 SQLite 3.x 作为 SQL 行为的唯一真值来源
- **原因**: SQLite 是标准 SQL 的权威实现，calibrator 可直接交叉校验 ExcelMCP 输出
- **影响**: INV-2 及所有 SQL 语义不变量以 SQLite 为准；引入 calibrator 模块

## D003: sqlglot 解析器

- **背景**: 需要将用户 SQL 转换为 Excel 可执行的操作
- **决策**: 使用 sqlglot 作为 SQL 解析和转换引擎
- **原因**: sqlglot 支持完整 SQL AST 操作，可做方言转换，社区活跃
- **影响**: SQL 解析依赖 sqlglot；已知 LIKE 转换 bug 需 workaround（P003）

## D004: 不变量驱动开发

- **背景**: 需要系统化保证 SQL-over-Excel 引擎的正确性
- **决策**: 建立公理→不变量→测试三层体系（AX→INV→test）
- **原因**: 从 4 条公理推导 32 条不变量，每条有测试覆盖，确保正确性可机械化检验
- **影响**: .omp/rules/axioms.md + tests/invariants/ 7 个测试文件 + evolve.js 退化检测

## D005: 进化引擎

- **背景**: Harness 自身需要版本管理和自动退化检测
- **决策**: 实现 evolve.js 五阶段进化引擎（propose → evaluate → commit → rollback → review）
- **原因**: 变更必须通过不变量测试验证才能提交，失败自动回滚
- **影响**: tools/evolve.js + data/carrier-registry.json + L5 健康检查体系
