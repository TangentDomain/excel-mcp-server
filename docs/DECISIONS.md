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
