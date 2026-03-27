# Changelog

## [Unreleased]

### 新增
- 30秒上手教程（README顶部，折叠详细安装）
- 竞品对比表（vs haris-musa/excelpython）
- 游戏策划SQL实战示例（5个真实场景）
- CHANGELOG 版本更新日志

### 文档
- 修复过时数据：44工具、1036测试

## [v1.1.0] - 2026-03-27

### 新增
- **REQ-025 返回值结构统一**：全部工具统一 `{success, message, data, meta}` 格式
- **REQ-028 FROM子查询**：`FROM (SELECT ...) AS alias` 语法完整实现
- 结构化SQL错误码：AI可自动解析并修复的error_code

### 优化
- 合并 `get_headers` 和 `get_sheet_headers` 为统一接口（45→44工具）
- preview_operation 合并到 assess_data_impact

## [v1.0.32] - 2026-03-26

### 新增
- **REQ-027 跨文件JOIN**：`@'filepath'` 语法，支持不同文件的工作表关联查询

## [v1.0.31] - 2026-03-25

### 优化
- 合并 preview_operation 到 assess_data_impact（46→45工具）
- HAVING 中文列名精确匹配建议

### 修复
- `__init__.py` 绝对导入导致测试路径问题

## [v1.0.30] - 2026-03-24

### 新增
- COALESCE / IFNULL 空值替换
- 字符串函数：UPPER, LOWER, CONCAT, REPLACE, LENGTH, SUBSTRING

## [v1.0.28] - 2026-03-23

### 新增
- **REQ-010 窗口函数**：ROW_NUMBER, RANK, DENSE_RANK + PARTITION BY
- EXISTS 关联子查询

## [v1.0.25] - 2026-03-21

### 新增
- 双行表头自动识别（中文描述 + 英文字段名）
- 智能列名建议（编辑距离匹配）

## [v1.0.20] - 2026-03-19

### 新增
- **REQ-012 CTE**：WITH ... AS 公共表表达式
- UNION / UNION ALL 合并查询

## [v1.0.15] - 2026-03-17

### 新增
- **REQ-011 子查询**：WHERE IN / NOT IN / 标量子查询
- CASE WHEN 条件表达式

## [v1.0.10] - 2026-03-15

### 新增
- **REQ-008 SQL查询引擎**：基于 sqlglot + pandas 的完整SQL支持
- **REQ-009 JOIN**：同文件内工作表关联查询
- 路径安全验证（SecurityValidator）
- 公式注入防护

## [v1.0.0] - 2026-03-12

### 初始发布
- 44个基础Excel操作工具
- python-calamine 高性能读取（Rust引擎）
- openpyxl 写入支持
- FastMCP 框架集成
