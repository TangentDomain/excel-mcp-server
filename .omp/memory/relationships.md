# 关系图谱

>项目依赖和协作关系的快速参考。

## 上游依赖（Upstream）

| 包 | 用途 | 版本约束 |
|----|------|---------|
| openpyxl | Excel 读写（.xlsx） | 主要 I/O 引擎 |
| sqlglot | SQL 解析和方言转换 | SQL AST 操作 |
| FastMCP | MCP 服务器框架 | 工具注册和协议 |
| calamine | Excel 高速读取（Rust） | 大文件场景备选 |

## 协作依赖（Collaboration）

| 包 | 用途 | 关系 |
|----|------|------|
| SQLite 3.x | SQL 真值来源 | calibrator 交叉校验 |
| pytest | 测试框架 | 不变量测试 + 常规测试 |
| ruff | Linter + Formatter | 质量门禁 |
| Bun | 运行时 | Harness 脚本执行 |

## 下游消费者（Downstream）

无。本项目是独立 MCP 服务器，无其他项目依赖它。

## 管理者（Manager）

| 实体 | 职责 |
|------|------|
| TangentDomain/excel-mcp-server | GitHub 仓库，CI/CD |
| Oh My Pi Harness | Agentify 智能体框架，.omp/ 目录管理 |
| evolve.js | Harness 进化引擎，五阶段变更管理 |
