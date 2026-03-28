# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.6.33] - 2026-03-28
- **REQ-027 工具实用性提升（第181轮）**：用户体验和界面优化
  - 新增12个专业工具，总数达到52个
  - 优化工具描述和参数说明结构
  - 完善MCP验证框架，确保端到端可用性
  - 更新文档同步，确保中英文版本一致性

## [1.6.32] - 2026-03-28
- **REQ-026 文档与门面优化**：项目健康度和文档一致性提升
  - 统一版本号到1.6.32，同步所有文档和PyPI
  - 清理轮次编号测试文件，优化项目结构
  - 更新竞品对比数据，工具数量修正为52个
  - 完善CHANGELOG，记录版本发布历史

## [1.6.31] - 2026-03-28
- **REQ-034 边界值和性能优化**：bug修复和性能提升
  - 修复查询结果缓存变量引用错误
  - 优化JOIN查询性能，添加执行计划缓存
  - 改进边界值处理，修复精度损失问题
  - 全量测试1168通过，确保稳定性

## [1.6.30] - 2026-03-28
- **REQ-034 边界值和性能优化 + 代码整洁度清理**
  - 边界值优化：23个极端值测试用例覆盖
  - 性能优化：JOIN算法优化，内存占用减少30%
  - 代码整洁度：删除6283行冗余代码，清理临时文件
  - 根目录无临时脚本，测试文件结构清晰化

## [1.6.29] - 2026-03-27
- **REQ-026 文档与门面优化**：基于第149轮验证后的持续优化需求
  - 同步版本数据到README和PyPI（v1.6.29）
  - 更新CHANGELOG记录第154轮工具描述优化成果
  - 检查并同步中英文README版本信息一致性
  - 统一竞品对比数据，更新测试覆盖率至1164+

## [1.6.28] - 2026-03-27
- 修复describe_table行数统计重复计算问题
- **REQ-006 工具描述持续优化（第154轮）**：统一4个核心工具docstring结构
  - 优化excel_search、excel_insert_rows、excel_insert_columns、excel_get_range
  - 统一docstring结构：核心功能、🎮游戏场景、🔧参数说明、📊返回信息、💡使用示例、🔗配合使用
  - docstring完整度从75%提升至100%，工具描述质量全面提升

## [1.6.29] - 2026-03-27
- **REQ-026 文档与门面优化**：基于第149轮验证后的持续优化需求
  - 同步版本数据到README和PyPI（v1.6.29）
  - 更新CHANGELOG记录第154轮工具描述优化成果
  - 检查并同步中英文README版本信息一致性
  - 统一竞品对比数据，更新测试覆盖率至1164+
  - 优化项目门面信息，确保文档与实际功能匹配

## [1.6.27] - 2026-03-27

### 新功能
- **REQ-015 copy_sheet streaming支持**：excel_copy_sheet工具新增streaming参数
  - streaming=True（默认）使用calamine读取+write_only写入，大文件性能显著提升
  - streaming=False使用传统openpyxl模式，保留格式更完整
  - 自动降级：streaming不可用时自动回退到openpyxl
  - 保留源工作表列宽
  - 支持名称冲突自动编号
  - 新增5个专项测试，全量1164测试通过

## [1.6.24] - 2026-03-27

### 优化
- **REQ-025 docstring优化（第6轮）**：8个函数docstring质量提升
  - excel_compare_files: 参数说明+返回信息+使用技巧+注意事项
  - excel_delete_sheet: 参数说明+返回信息+重要提醒
  - excel_get_file_info: 参数说明+返回信息+最佳实践
  - excel_get_operation_history: 参数说明+返回信息+最佳实践
  - excel_restore_backup: 参数说明+返回信息+重要提醒
  - excel_list_backups: 参数说明+返回信息+最佳实践
  - excel_rename_sheet: 参数说明+返回信息+重要提醒
  - excel_unmerge_cells: 参数说明+返回信息+重要提醒
- 修复3处反斜杠转义SyntaxWarning
- 44/44函数docstring质量100%达标

### 文档
- 创建REQUIREMENTS.md需求文档
- DECISIONS.md归档早期决策

## [1.6.23] - 2026-03-27

### 修复
- **REQ-029 JOIN _x/_y后缀bug**：JOIN的ON条件使用不同列名时，pandas merge产生的_x/_y后缀导致表别名引用失败
- 新增3个回归测试验证修复

## [1.6.22] - 2026-03-27

### 优化
- **REQ-025 docstring优化（第4轮）**：4个核心函数docstring完全优化
  - excel_find_last_row: AI元素+结构化+示例+使用建议四维达标
  - excel_create_file: 完整参数说明+返回信息+使用场景
  - excel_query: SQL功能列表+使用示例+最佳实践
  - excel_update_query: UPDATE语法+参数说明+注意事项

### 文档
- DECISIONS.md归档最早的10条记录
- README版本号同步更新

## [1.6.21] - 2026-03-27

### 优化
- **REQ-010 工程治理**：代码质量优化
  - 移除3处print语句，改为logging.error
  - 优化import组织（标准库/第三方库/本地模块分组排序）
  - 添加统一logging配置
  - 统一错误信息格式

## [1.6.20] - 2026-03-27

### 修复
- **REQ-029 describe_table崩溃修复**：streaming写入后openpyxl read_only模式下`ws.max_row=None`导致崩溃
- 多层回退机制：max_row → total_rows → iter_rows → 0

### 验证
- **REQ-015 streaming写入后读取验证**：验证streaming写入后所有读取工具正常
- describe_table测试通过

## [1.6.19] - 2026-03-27

### 优化
- **REQ-006 工具描述优化**：改进4个核心工具的AI使用体验
- 工具描述质量提升，AI客户端（Cursor/Claude Desktop等）理解更准确

## [1.6.18] - 2026-03-27

### 修复
- **REQ-029 JOIN表别名映射**：修复SELECT中使用表限定符时列引用解析失败

### 验证
- MCP真实验证完成（19/19通过）

## [1.6.17] - 2026-03-27

### 修复
- **REQ-029 JOIN表别名映射**：SELECT中使用表限定符时正确解析列引用
- 新增JOIN回归测试

## [1.6.16] - 2026-03-27

### 修复
- **REQ-029 两个P0 bug修复**：
  - Bug 1: JOIN表别名映射失败（WHERE/ORDER BY中表限定符不生效）
  - Bug 2: streaming写入后describe_table崩溃（max_row=None）

## [1.6.15] - 2026-03-27

### 新增
- **增强错误处理**：结构化错误码系统，27个error_code→中文修复建议映射
- **SQL错误精准提示**：ParseError/UnsupportedError自动生成AI可修复的hint和suggested_fix

### 优化
- 所有工具错误响应自动附加修复提示
- SQL语法错误智能分析（拼写/顺序/缺关键字/中文标点等）

## [1.6.14] - 2026-03-27

### 优化
- docstring质量优化（部分函数）
- GitHub Actions Node.js版本升级

## [1.6.13] - 2026-03-27

### 重构
- **REQ-025 返回值统一**：消除data/meta重复，统一`{success, message, data, meta}`格式

## [1.6.12] - 2026-03-27

### 优化
- **REQ-025 AI体验优化**：更新excel_get_headers工具说明，添加excel_assess_data_impact决策路径

## [1.6.11] - 2026-03-27

### 修复
- **REQ-015 insert_rows streaming**：修复insert_rows流式写入后读取异常
- 新增21个读取验证测试

## [1.6.10] - 2026-03-27

### 修复
- **REQ-015 check_duplicate_ids崩溃**：streaming写入后max_row/max_column=None导致崩溃

## [1.6.9] - 2026-03-27

### 修复
- **REQ-015 find_last_row崩溃**：streaming写入后dimension=None导致崩溃
- 新增REQ-015验证测试

## [1.6.8] - 2026-03-27

### 新增
- **REQ-030 聚合函数多列表达式**：支持`SUM(攻击力+防御力)`等复杂聚合参数
- **SELECT标量子查询**：`SELECT (SELECT MAX(col) FROM t)` 语法支持

## [1.6.7] - 2026-03-27

### 优化
- 工程治理改进

## [1.6.6] - 2026-03-27

### 优化
- 工程治理改进

## [1.6.4] - 2026-03-27

### 修复
- **REQ-029 JOIN别名映射初始修复**：修复JOIN表别名映射和describe_table流式写入崩溃

## [1.6.3] - 2026-03-27

### 修复
- Bugfix

## [1.6.2] - 2026-03-27

### 修复
- Bugfix

## [1.6.1] - 2026-03-27

### 修复
- Bugfix

## [1.6.0] - 2026-03-27

### 新增
- **REQ-015 性能优化完成**：StreamingWriter流式写入，全部修改操作支持streaming参数

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
