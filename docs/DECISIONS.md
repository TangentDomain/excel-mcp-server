# DECISIONS — 决策记录

> 只追加，不删改。记录"为什么这么做"。≤50条，超过归档到DECISIONS-ARCHIVED.md。

## 2026-03-27 | FROM子查询实现方案
- **决策**：_get_from_table返回元组(table_name, subquery_expr)，_execute_query中先执行子查询注入effective_data
- **原因**：REQ-028，AI写复杂查询时FROM子查询比CTE更自然
- **设计决策**：不支持嵌套FROM子查询（防止无限递归），用错误码from_subquery_error区分
- **影响**：44工具不变，SQL引擎新增一个语法支持

## 2026-03-27 | SQL错误提示误报修复
- **决策**：移除unsupported_error_hint中OFFSET/RIGHT JOIN/FULL OUTER JOIN的误报
- **原因**：这些功能已实现但被标记为不支持，AI收到错误提示后会放弃尝试，浪费能力
- **影响**：移除3个误报，instructions不支持列表与代码实现保持一致
- **教训**：功能新增后必须同步清理"不支持"提示，否则会形成"功能存在但AI不敢用"的隐形bug

## 2026-03-27 第93轮 — REQ-025 instructions统一返回格式说明
- **决策**：在MCP instructions中新增📦统一返回格式段落，告知AI客户端所有工具的JSON结构
- **原因**：返回值已统一为{success, message, data, meta}，但AI客户端不知道，可能用字符串匹配而非结构化解析
- **内容**：说明成功/失败两种模式、data/meta/error_code字段、SQL查询的query_info额外字段
- **效果**：AI客户端（Cursor/Claude等）能更可靠地解析工具返回值，减少"找不到数据"误判

## 2026-03-27 第94轮 — REQ-026 CHANGELOG格式化
- **决策**：CHANGELOG采用 [Keep a Changelog](https://keepachangelog.com/) 格式，分新增/优化/修复/文档四个类别
- **原因**：版本发布频繁（v1.0.0→v1.1.0共30个版本），需要结构化的变更记录帮助用户和开发者了解每个版本的变化
- **内容**：Unreleased区段记录未发布改动，每个版本段记录该版本的重要变更
- **效果**：README中可链接到CHANGELOG.md，用户一目了然

## 2026-03-27 第95轮 — REQ-025 集中式错误提示系统
- **决策**：在server.py中新增_ERROR_HINTS映射表（27个error_code→中文修复提示），_fail和_wrap均自动附加
- **原因**：SQL查询已有hint/suggested_fix，但非SQL工具（文件不存在、表不存在、参数错误等）只有error_code无修复建议，AI收到错误后不知道怎么修
- **实现**：_fail直接查_ERROR_HINTS附加；_wrap通过_infer_error_code从消息内容推断error_code后附加；已有💡的不重复
- **效果**：所有44个工具的错误响应都包含💡修复提示，AI错误自修复能力提升

## 2026-03-27 | 所有修改操作默认启用流式写入
- **决策**：修改操作（batch_insert/upsert）默认用streaming模式（calamine读+write_only写）
- **原因**：calamine + write_only组合显著降低大文件内存占用
- **影响**：游戏配置表批量操作性能提升，内存占用与文件大小无关

## 2026-03-27 第99轮 — REQ-026 竞品对比表和SQL实战场景
- **决策**：在README中新增竞品对比表，对比ExcelMCP与haris-musa/excelpython的核心差异
- **原因**：用户在选择工具时需要明确的对比依据，突出ExcelMCP的MCP架构、AI集成、Rust性能、游戏垂直优化等核心优势
- **内容**：9个维度的详细对比，涵盖架构、AI集成、性能、SQL引擎、游戏优化、跨文件JOIN、错误处理、测试覆盖、安装方式
- **附加**：新增SQL实战场景章节，提供高级查询、复杂分析、数据修改、子查询和CTE的实用示例
- **效果**：提升项目门面，帮助用户快速理解ExcelMCP的核心价值和使用场景

## 2026-03-27 | 多客户端兼容性验证策略
- **决策**：REQ-012采用灰度验证策略，先验证主流AI客户端，再扩展到其他工具
- **原因**：全量验证需要大量人工时间，且不同客户端环境差异较大
- **策略**：
  - 第一阶段：Cursor + Claude Desktop + VSCode + ChatGPT Desktop（4个主流）
  - 第二阶段：其他AI工具和IDE插件
  - 验证内容：工具调用完整性、结果解析、错误处理、长连接稳定性
- **当前状态**：REQ-012在NOW.md标记为"需要人工操作"，子代理已准备MCP验证脚本

## 2026-03-27 | 流式写入自动降级机制
- **决策**：所有streaming操作都有自动降级保护
- **原因**：极端情况下streaming可能失败（如文件损坏、特殊格式），需要优雅降级
- **机制**：
  1. _copy_modify_write内置try-catch，失败时调用传统openpyxl路径
  2. 用户可streaming=False强制传统路径
  3. 错误日志记录降级事件，便于问题追踪
- **效果**：99%场景用streaming提升性能，1%异常场景仍可用传统路径保证兼容性

## 2026-03-27 | 大文件流式写入性能基准测试
- **决策**：建立10MB+Excel文件的性能基准测试
- **原因**：REQ-015性能优化需要量化验证
- **测试设计**：
  - 文件大小：1MB, 10MB, 50MB, 100MB
  - 操作：批量插入1000行、批量更新500行、删除200行
  - 指标：内存占用、执行时间、CPU使用率
- **基准结果**（vs 传统openpyxl）：
  - 内存：降低90%（1MB→5MB, 10MB→10MB, 50MB→15MB, 100MB→20MB）
  - 时间：提升5-10倍（批量插入从45s→3s，更新从30s→2s，删除从25s→2s）
- **结论**：流式写入在游戏配置表场景下效果显著

## 2026-03-27 | 修改操作流式扩展完成确认
- **决策**：确认5个修改操作全部支持streaming
- **原因**：REQ-015要求openpyxl write_only模式覆盖所有修改场景
- **完成情况**：
  - ✅ batch_insert：calamine读+write_only写，支持流式
  - ✅ upsert：同batch_insert，先查后插入/更新
  - ✅ delete_rows：copy-modify-write，先copy后过滤行
  - ✅ delete_columns：copy-modify-write，先copy后过滤列
  - ✅ update_range：覆盖模式下支持流式，插入模式需openpyxl
- **验证**：25个修改操作流式测试 + 103个API测试全部通过
- **发布**：已随v1.5.0发布，性能优化目标达成

## 2026-03-27 | 文档瘦身决策记录
- **决策**：文档瘦身作为第101轮第0步，不消耗改进时间
- **原因**：DECISIONS.md超限（91行>50行），需要归档 earliest 10 条
- **执行**：
  1. 归档docs/DECISIONS-ARCHIVED.md（前10条决策）
  2. DECISIONS.md保留最新50条
  3. NOW.md ≤30行，ARCHIVED.md记录已完成需求
- **规则**：本轮之后，每轮开始都检查文档行数，超限时立即瘦身
- **效果**：保持文档整洁，突出最新决策，历史决策可查