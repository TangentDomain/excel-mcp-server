# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [v1.6.44] - 2026-03-28

### 优化
- **代码质量重构**：`excel_describe_table` 函数从216行拆分为6个职责单一的小函数
  - `_detect_dual_header()`：双行表头检测（32行）
  - `_collect_column_statistics()`：列统计信息收集（33行）
  - `_build_describe_columns()`：列类型推断和结果构建（33行）
  - `_resolve_row_count()`：行数统计含streaming兼容（44行）
  - `_prepare_describe_result()`：结果格式化（22行）
  - 主函数从216行降至129行（减少40%）
- 新增 `scripts/code_quality_analysis.py` 代码质量分析工具

## [v1.6.43] - 2026-03-28

### 新增
- **REQ-027 GitHub star 提升计划**：GitHub star 统计和激励系统
- GitHub star badge 显示功能，支持实时stars数量展示
- star-thanks.py 自动化脚本，动态更新star感谢信息
- README 优化，添加star引导激励机制
- 中英文README完全同步，版本信息统一

### 优化
- **REQ-026 文档优化**：自动化版本检查脚本功能完善
- 修复CHANGELOG.md版本同步问题，确保文档一致性
- 更新项目健康度自检逻辑，文档瘦身规则执行
- 版本徽章同步更新，测试数量统计修正

## [v1.6.42] - 2026-03-28

### 优化
- **REQ-026 文档与门面优化**：持续监控和改进
- 项目健康度自检机制完善，根目录垃圾文件清理
- 文档结构优化，DECISIONS.md瘦身至27行
- 中英文README内容同步检查机制
- 自动化版本检查脚本功能验证

## [v1.6.41] - 2026-03-28

### 优化
- **REQ-026 文档与门面优化**：项目健康度监控
- 全量测试1161+通过，项目健康度极佳
- MCP真实验证12/12通过，功能稳定性验证
- 文档版本一致性保持，测试数量统计更新
- CHANGELOG去重处理，数据校正完成

## [v1.6.40] - 2026-03-28

### 新增
- **REQ-027 GitHub star 提升计划**：GitHub star 统计和激励系统
- 新增 CONTRIBUTING.md 贡献指南和 GitHub 模板
- 创建 star-thanks.py 自动化脚本和统计系统

### 优化
- **REQ-026 文档优化**：自动化版本检查脚本，确保文档版本同步
- 修复README.md/README.en.md版本徽章格式问题
- 统一 pyproject.toml/__init__.py/README.md/README.en.md 版本号
- 清理 HTML 徽章格式，修复缺失的 closing brackets

## [v1.6.39] - 2026-03-27

### 新增
- **REQ-015 copy_sheet streaming支持**：excel_copy_sheet工具新增streaming参数
  - streaming=True（默认）使用calamine读取+write_only写入，大文件性能显著提升
  - streaming=False使用传统openpyxl模式，保留格式更完整
  - 自动降级：streaming不可用时自动回退到openpyxl
  - 保留源工作表列宽
  - 支持名称冲突自动编号
  - 新增5个专项测试，全量测试通过

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

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 JOIN _x/_y后缀bug**：JOIN的ON条件使用不同列名时，pandas merge产生的_x/_y后缀导致表别名引用失败
- 新增3个回归测试验证修复

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-025 docstring优化（第4轮）**：4个核心函数docstring完全优化
  - excel_find_last_row: AI元素+结构化+示例+使用建议四维达标
  - excel_create_file: 完整参数说明+返回信息+使用场景
  - excel_query: SQL功能列表+使用示例+最佳实践
  - excel_update_query: UPDATE语法+参数说明+注意事项

### 文档
- DECISIONS.md归档最早的10条记录
- README版本号同步更新

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-010 工程治理**：代码质量优化
  - 移除3处print语句，改为logging.error
  - 优化import组织（标准库/第三方库/本地模块分组排序）
  - 添加统一logging配置
  - 统一错误信息格式

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 describe_table崩溃修复**：streaming写入后openpyxl read_only模式下`ws.max_row=None`导致崩溃
- 多层回退机制：max_row → total_rows → iter_rows → 0

### 验证
- **REQ-015 streaming写入后读取验证**：验证streaming写入后所有读取工具正常
- describe_table测试通过

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-006 工具描述优化**：改进4个核心工具的AI使用体验
- 工具描述质量提升，AI客户端（Cursor/Claude Desktop等）理解更准确

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 JOIN表别名映射**：修复SELECT中使用表限定符时列引用解析失败

### 验证
- MCP真实验证完成（19/19通过）

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 JOIN表别名映射**：SELECT中使用表限定符时正确解析列引用
- 新增JOIN回归测试

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 两个P0 bug修复**：
  - Bug 1: JOIN表别名映射失败（WHERE/ORDER BY中表限定符不生效）
  - Bug 2: streaming写入后describe_table崩溃（max_row=None）

## [v1.6.39] - 2026-03-27

### 新增
- **增强错误处理**：结构化错误码系统，27个error_code→中文修复建议映射
- **SQL错误精准提示**：ParseError/UnsupportedError自动生成AI可修复的hint和suggested_fix

### 优化
- 所有工具错误响应自动附加修复提示
- SQL语法错误智能分析（拼写/顺序/缺关键字/中文标点等）

## [v1.6.39] - 2026-03-27

### 优化
- docstring质量优化（部分函数）
- GitHub Actions Node.js版本升级

## [v1.6.39] - 2026-03-27

### 重构
- **REQ-025 返回值统一**：消除data/meta重复，统一`{success, message, data, meta}`格式

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-025 AI体验优化**：更新excel_get_headers工具说明，添加excel_assess_data_impact决策路径

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-015 insert_rows streaming**：修复insert_rows流式写入后读取异常
- 新增21个读取验证测试

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-015 check_duplicate_ids崩溃**：streaming写入后max_row/max_column=None导致崩溃

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-015 find_last_row崩溃**：streaming写入后dimension=None导致崩溃
- 新增REQ-015验证测试

## [v1.6.39] - 2026-03-27

### 新增
- **REQ-030 聚合函数多列表达式**：支持`SUM(攻击力+防御力)`等复杂聚合参数
- **SELECT标量子查询**：`SELECT (SELECT MAX(col) FROM t)` 语法支持

## [v1.6.39] - 2026-03-27

### 优化
- 工程治理改进

## [v1.6.39] - 2026-03-27

### 优化
- 工程治理改进

## [v1.6.39] - 2026-03-27

### 修复
- **REQ-029 JOIN别名映射初始修复**：修复JOIN表别名映射和describe_table流式写入崩溃

## [v1.6.39] - 2026-03-27

### 修复
- Bugfix

## [v1.6.39] - 2026-03-27

### 修复
- Bugfix

## [v1.6.39] - 2026-03-27

### 修复
- Bugfix

## [v1.6.39] - 2026-03-27

### 新增
- **REQ-015 性能优化完成**：StreamingWriter流式写入，全部修改操作支持streaming参数

## [v1.6.39] - 2026-03-27

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

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-015 excel_update_query流式写入**：UPDATE语句自动选择高性能路径
  - 智能决策：affected_rows≥50 / changes≥100 / 文件>1MB时使用streaming
  - copy-modify-write方案：calamine读取 → 内存修改 → write_only写入
  - 流式写入失败自动降级到传统openpyxl方式
  - 返回值新增method字段标识写入方式（streaming/traditional）
  - 1099测试全通过，8项游戏场景验证通过

## [v1.6.39] - 2026-03-27

### 优化
- **REQ-006 全部43个工具docstring优化完成**：统一emoji标题+核心功能+游戏开发场景+参数说明+使用建议格式
- 覆盖所有MCP工具，AI客户端（Cursor/Claude Desktop等）能更好理解工具用途
- 修复docstring中反斜杠转义导致的SyntaxWarning

## [v1.6.39] - 2026-03-27

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

## [v1.6.39] - 2026-03-27

### 新增
- **REQ-025 返回值结构统一**：全部工具统一 `{success, message, data, meta}` 格式
- **REQ-028 FROM子查询**：`FROM (SELECT ...) AS alias` 语法
- 结构化SQL错误码：AI可自动解析并修复的 `error_code`

### 优化
- 合并 `get_headers` 和 `get_sheet_headers` 为统一接口（45→44工具）
- preview_operation 合并到 assess_data_impact

## [v1.6.39] - 2026-03-26

### 新增
- **REQ-027 跨文件JOIN**：`@'filepath'` 语法，支持不同文件的工作表关联查询

## [v1.6.39] - 2026-03-25

### 优化
- 合并 preview_operation 到 assess_data_impact（46→45工具）
- HAVING 中文列名精确匹配建议

### 修复
- `__init__.py` 绝对导入导致测试路径问题

## [v1.6.39] - 2026-03-24

### 新增
- COALESCE / IFNULL 空值替换
- 字符串函数：UPPER, LOWER, CONCAT, REPLACE, LENGTH, SUBSTRING

## [v1.6.39] - 2026-03-23

### 新增
- **REQ-010 窗口函数**：ROW_NUMBER, RANK, DENSE_RANK + PARTITION BY
- EXISTS 关联子查询

## [v1.6.39] - 2026-03-21

### 新增
- 双行表头自动识别（中文描述 + 英文字段名）
- 智能列名建议（编辑距离匹配）

## [v1.6.39] - 2026-03-19

### 新增
- **REQ-012 CTE**：WITH ... AS 公共表表达式
- UNION / UNION ALL 合并查询

## [v1.6.39] - 2026-03-17

### 新增
- **版本检查自动化**：自动化版本检查脚本，减少文档同步错误
- **REQ-0192 自我进化**：持续监控和优化文档同步流程

## [v1.6.39] - 2026-03-15

### 新增
- **REQ-008 SQL查询引擎**：基于 sqlglot + pandas 的完整SQL支持
- **REQ-009 JOIN**：同文件内工作表关联查询
- 路径安全验证（SecurityValidator）
- 公式注入防护

## [v1.6.39] - 2026-03-12

### 初始发布
- 44个基础Excel操作工具
- python-calamine 高性能读取（Rust引擎）
- openpyxl 写入支持
- FastMCP 框架集成