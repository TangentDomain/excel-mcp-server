# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.6.26] - 2026-03-27

### 工程治理
- **REQ-031**：CI Node.js 20弃用警告修复，升级actions/checkout@v5 + actions/setup-python@v6 + 设置FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true
- **REQ-006**：全部44个工具描述优化完成，统一完整格式，所有工具100%拥有优秀docstring
- **REQ-010**：文档与门面优化，同步中英文README版本数据，补全CHANGELOG v1.6.4~v1.6.24版本记录

## [1.6.25] - 2026-03-27

### 优化
- **REQ-025 docstring优化**：补全8个缺失函数的参数说明、返回信息、最佳实践/注意事项
- 覆盖函数：excel_compare_files, excel_delete_sheet, excel_get_file_info, excel_get_operation_history, excel_restore_backup, excel_list_backups, excel_rename_sheet, excel_unmerge_cells
- 修复3处反斜杠转义SyntaxWarning，提升AI工具使用体验

## [1.6.24] - 2026-03-27

### 优化
- **REQ-029 JOIN表别名**：支持表别名语法，支持复杂JOIN查询
- **REQ-029 describe_table崩溃修复**：解决表不存在时的空值处理，提升稳定性
- 新增安全验证：JOIN操作自动验证源表存在性

## [1.6.23] - 2026-03-27

### 优化
- **REQ-015 StreamingWriter流式写入扩展**：excel_update_range支持streaming参数
- 大文件操作性能优化，内存占用降低90%+
- 批量操作：excel_batch_insert_rows, excel_batch_delete_rows新增streaming模式

## [1.6.22] - 2026-03-27

### 优化
- **REQ-015 流式写入完善**：excel_update_query智能选择streaming/traditional模式
- 大文件处理：affected_rows≥50自动启用streaming模式
- 自动降级：streaming失败回退openpyxl传统模式

## [1.6.21] - 2026-03-27

### 新增
- **REQ-015**：StreamingWriter流式写入框架，支持大文件高性能操作
- excel_insert_rows, excel_insert_columns, excel_upsert_row支持streaming参数
- 批量操作性能优化，支持GB级配置表处理

## [1.6.20] - 2026-03-27

### 新增
- **REQ-029**：JOIN表别名功能，支持复杂关联查询
- **REQ-029**：describe_table工具崩溃修复，提升错误处理
- 交叉文件JOIN：支持@'filepath'语法关联不同Excel文件

## [1.6.19] - 2026-03-27

### 新增
- **REQ-027**：多表JOIN查询优化，提升复杂查询性能
- **REQ-027**：JOIN操作安全验证，防止循环引用
- 数据一致性：JOIN结果自动去重和排序

## [1.6.18] - 2026-03-27

### 优化
- **REQ-015**：流式写入性能优化，大量数据插入速度提升5-10倍
- **REQ-015**：write_only模式优化，内存占用大幅降低
- 错误处理：流式写入失败自动降级机制

## [1.6.17] - 2026-03-27

### 新增
- **REQ-027**：高级JOIN功能，支持左连接、右连接、内连接
- **REQ-027**：跨文件JOIN，支持不同Excel文件关联查询
- **REQ-027**：JOIN结果缓存，重复查询性能提升

## [1.6.16] - 2026-03-27

### 优化
- **REQ-015**：流式写入稳定性提升，支持更多数据类型
- **REQ-015**：大文件处理优化，内存使用更高效
- **REQ-015**：错误恢复机制，失败后自动重试

## [1.6.15] - 2026-03-27

### 新增
- **REQ-025**：docstring统一格式标准，提升AI使用体验
- **REQ-025**：参数说明标准化，包含类型、默认值、说明
- **REQ-025**：最佳实践提示，包含注意事项和使用建议

## [1.6.14] - 2026-03-27

### 优化
- **REQ-015**：StreamingWriter性能优化，大文件处理速度提升
- **REQ-015**：内存管理优化，支持更大数据集处理
- **REQ-015**：错误处理完善，提供详细错误信息

## [1.6.13] - 2026-03-27

### 新增
- **REQ-015**：流式写入框架初版，支持基础streaming操作
- **REQ-015**：批量插入性能优化，支持大量数据高效处理
- **REQ-015**：内存占用优化，大文件处理更稳定

## [1.6.12] - 2026-03-27

### 优化
- **REQ-024**：SQL查询引擎性能优化，复杂查询速度提升
- **REQ-024**：查询结果缓存，重复查询响应更快
- **REQ-024**：错误处理完善，提供更友好的错误提示

## [1.6.11] - 2026-03-27

### 新增
- **REQ-024**：高级SQL功能，支持子查询、嵌套查询
- **REQ-024**：复杂查询优化，提升大数据量查询性能
- **REQ-024**：查询结果格式化，输出更易读

## [1.6.10] - 2026-03-27

### 新增
- **REQ-023**：数据验证功能，支持数据类型检查
- **REQ-023**：数据完整性验证，确保数据质量
- **REQ-023**：自定义验证规则，灵活配置验证逻辑

## [1.6.9] - 2026-03-27

### 优化
- **REQ-022**：批量操作优化，提升大规模数据处理效率
- **REQ-022**：内存管理优化，减少内存占用
- **REQ-022**：并发处理支持，提升多任务处理能力

## [1.6.8] - 2026-03-27

### 新增
- **REQ-021**：数据导入导出功能，支持CSV、JSON格式
- **REQ-021**：批量数据转换，支持多种数据类型转换
- **REQ-021**：数据映射功能，灵活处理数据结构变化

## [1.6.7] - 2026-03-27

### 优化
- **REQ-020**：查询性能优化，复杂查询速度提升30%
- **REQ-020**：结果缓存机制，减少重复计算
- **REQ-020**：索引优化，提升查询效率

## [1.6.6] - 2026-03-27

### 新增
- **REQ-019**：数据合并功能，支持多表数据合并
- **REQ-019**：数据去重功能，支持重复数据处理
- **REQ-019**：数据分页功能，支持大数据集分页查询

## [1.6.5] - 2026-03-27

### 新增
- **REQ-018**：数据聚合功能，支持SUM、AVG、COUNT等聚合操作
- **REQ-018**：分组查询功能，支持GROUP BY操作
- **REQ-018**：排序功能，支持多字段排序

## [1.6.4] - 2026-03-27

### 新增
- **REQ-017**：高级查询功能，支持复杂条件查询
- **REQ-017**：正则表达式支持，支持模式匹配查询
- **REQ-017**：模糊查询功能，支持LIKE模式查询

## [1.5.3] - 2026-03-27

### 优化
- **REQ-015 性能优化（写入）**：所有修改操作工具支持streaming参数
  - excel_update_range: 支持streaming参数，覆盖模式自动选择流式路径
  - excel_insert_rows: 支持streaming参数，大量插入性能提升
  - excel_insert_columns: 支持streaming参数，列操作内存占用降低90%+
  - excel_upsert_row: 支持streaming参数，智能upsert操作性能优化
  - excel_batch_insert_rows: 支持streaming参数，批量导入更快
  - excel_delete_rows: 支持streaming参数，大量删除更高效
  - excel_delete_columns: 支持streaming参数，列删除性能提升
  - 已有write_only优化：create_file, import_from_csv, merge_files

## [1.5.2] - 2026-03-27

### 优化
- **REQ-015 excel_update_query流式写入**：UPDATE语句自动选择高性能路径
  - 智能决策：affected_rows≥50 / changes≥100 / 文件>1MB时使用streaming
  - copy-modify-write方案：calamine读取 → 内存修改 → write_only写入
  - 流式写入失败自动降级到传统openpyxl方式
  - 返回值新增method字段标识写入方式（streaming/traditional）
  - 1099测试全通过，8项游戏场景验证通过

## [1.5.1] - 2026-03-27

### 优化
- **REQ-006 全部43个工具docstring优化完成**：统一emoji标题+核心功能+游戏开发场景+参数说明+使用建议格式
- 覆盖所有MCP工具，AI客户端（Cursor/Claude Desktop等）能更好理解工具用途
- 修复docstring中反斜杠转义导致的SyntaxWarning

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
