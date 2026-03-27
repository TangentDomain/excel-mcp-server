# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.3.0] - 2026-03-27

### 新增
- **REQ-015 StreamingWriter流式写入**：`core/streaming_writer.py`，calamine读取 + write_only流式写入
- **batch_insert_rows/upsert_row流式模式**：默认`streaming=True`，大文件内存占用大幅降低
- 15个StreamingWriter新测试

### 优化
- **calamine浮点数兼容**：整数→浮点数（2→2.0）标准化比较
- **自动降级**：streaming失败自动回退openpyxl传统路径
- **列宽保留**：流式写入时保留列宽设置

### 新增
- **REQ-015 写入性能优化**：新建文件使用 `write_only` 模式（流式写入），新建和格式转换、文件合并场景内存占用大幅降低
- **FROM子查询**：`FROM (SELECT ...) AS alias` 语法完整实现，12个测试覆盖（WHERE过滤、JOIN结果子查询、嵌套子查询拒绝、DISTINCT、无别名等）
- **集中式错误提示系统**：27个error_code→中文💡修复建议映射，所有工具（含Operations层）错误响应自动附加修复提示

### 优化
- **REQ-025 返回值结构统一**：全部44个工具统一 `{success, message, data, meta}` 格式，AI客户端只需检查 `success` 字段
- **REQ-025 合并重复工具**：`get_headers` 和 `get_sheet_headers` 合并为统一接口（可选 `sheet_name`）
- **REQ-025 preview/assess合并**：`preview_operation` 合并到 `assess_data_impact`
- **REQ-025 instructions统一返回格式说明**：MCP instructions 新增 📦 统一返回格式段落，告知AI客户端JSON结构
- **REQ-025 SQL错误提示精准化**：修复 OFFSET/RIGHT JOIN/FULL OUTER JOIN 误报为不支持，instructions 不支持列表与代码实现同步
- **REQ-025 SQL错误结构化**：增强 ParseError/UnsupportedError 的 AI 可修复提示，新增 `suggested_fix` 字段（自动生成修复SQL）

### 文档
- **REQ-026 英文README同步**：30-Second Setup、竞品对比表、SQL实战场景、数据修正
- **REQ-026 30秒上手教程**：前置到README最顶部，一行配置+自然语言示例
- **REQ-026 竞品对比表**：新增 vs haris-musa/excelpython 对比，突出SQL引擎+Rust性能+游戏垂直
- README 新增 FROM子查询、窗口函数、CTE 示例
- README/README.en 测试数修正（1041）

## [v1.1.0] - 2026-03-27

### 新增
- **REQ-025 返回值结构统一**：全部工具统一 `{success, message, data, meta}` 格式
- **REQ-028 FROM子查询**：`FROM (SELECT ...) AS alias` 语法
- 结构化SQL错误码：AI可自动解析并修复的 `error_code`

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
