# DECISIONS — 决策记录

> 只追加，不删改。记录"为什么这么做"。≤50条，超过归档到DECISIONS-ARCHIVED.md。

## 2026-03-26 | 不做配置导出引擎
- 原因：构建工具属于CI/CD环节，不是MCP查询工具
- 影响：聚焦SQL引擎

## 2026-03-26 | calamine替换openpyxl读取
- 原因：2300x性能提升
- 影响：REQ-015读取部分完成

## 2026-03-26 | 取消View/写入校验/Auto Increment
- 原因：SQL引擎已覆盖（FK用JOIN查、范围用WHERE、枚举用IN）
- 影响：需求池精简，避免过度设计

## 2026-03-26 | MCP工具的用户是AI不是策划
- 原因：策划说一句话，AI翻译成工具调用
- 影响：优化方向从"人看得懂"转为"AI用得好"

## 2026-03-26 | 废弃scorecard和evolution-log
- 原因：子代理从没维护过，信息在每轮输出里
- 影响：用NOW.md替代，精简5000行文档

## 2026-03-27 | 文档体系重构
- 原因：历史/现在/未来混在一起，看不到重点
- 影响：NOW.md聚焦+ROADMAP定方向+DECISIONS记决策

## 2026-03-27 | 子代理偷懒问题
- 原因：子代理自行改focus为"维护模式"然后不做实质工作
- 影响：cron prompt加红线约束，禁止子代理改focus/ROADMAP，禁止自行暂停

## 2026-03-27 | 敏感信息泄露教训
- 原因：PyPI token写入docs/RULES.md并提交，GitHub push protection拒绝
- 影响：git reset清理commit历史，token移到.cron-prompt.md（不入库）
- 规则：提交前必须grep检查敏感信息，入库文件用引用不写值

## 2026-03-27 | FROM子查询实现方案
- 原因：REQ-028，AI写复杂查询时FROM子查询比CTE更自然
- 方案：_get_from_table返回元组(table_name, subquery_expr)，_execute_query中先执行子查询注入effective_data
- 设计决策：不支持嵌套FROM子查询（防止无限递归），用错误码from_subquery_error区分
- 影响：44工具不变，SQL引擎新增一个语法支持

## 2026-03-27 | SQL错误提示误报修复
- 原因：_unsupported_error_hint中OFFSET/RIGHT JOIN/FULL OUTER JOIN被标为不支持，但代码实际已支持。AI收到错误提示后会放弃尝试，浪费能力。
- 影响：移除3个误报，instructions不支持列表与代码实现保持一致
- 教训：功能新增后必须同步清理"不支持"提示，否则会形成"功能存在但AI不敢用"的隐形bug

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
- 原因：calamine + write_only组合显著降低大文件内存占用
- 影响：游戏配置表批量操作性能提升，内存占用与文件大小无关

2026-03-27 第99轮 — REQ-026 竞品对比表和SQL实战场景
- **决策**：在README中新增竞品对比表，对比ExcelMCP与haris-musa/excelpython的核心差异
- **原因**：用户在选择工具时需要明确的对比依据，突出ExcelMCP的MCP架构、AI集成、Rust性能、游戏垂直优化等核心优势
- **内容**：9个维度的详细对比，涵盖架构、AI集成、性能、SQL引擎、游戏优化、跨文件JOIN、错误处理、测试覆盖、安装方式
- **附加**：新增SQL实战场景章节，提供高级查询、复杂分析、数据修改、子查询和CTE的实用示例
- **效果**：提升项目门面，帮助用户快速理解ExcelMCP的核心价值和使用场景
