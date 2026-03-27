# DECISIONS — 决策记录 (归档)

> 已归档最早的10条决策（2026-03-26/27），最新决策见docs/DECISIONS.md。

## 2026-03-27 | 流式写入扩展至修改操作
- **决策**：将copy-modify-write方案扩展到delete_rows/delete_columns/update_range
- **原因**：这些操作都需要读取整个文件并重写，大文件场景下openpyxl内存占用高
- **方案**：提取_copy_modify_write通用方法，接受modify_fn回调；各操作实现自己的修改逻辑
- **约束**：update_range流式仅覆盖模式生效（insert_mode=False + preserve_formulas=False），插入模式和公式保留必须用openpyxl
- **自动降级**：streaming失败自动回退openpyxl传统路径；用户可streaming=False强制传统路径
- **效果**：5个修改操作全部支持streaming（batch_insert/upsert/delete_rows/delete_columns/update_range）

## 2026-03-27 | StreamingWriter用copy-modify-write方案
- **决策**：修改操作（batch_insert/upsert）默认用streaming模式（calamine读+write_only写）
- **原因**：openpyxl load_workbook对大文件内存占用高；calamine读+write_only写内存与文件大小无关
- **权衡**：保留列宽/行高/数据值，不保留单元格格式（字体/填充/边框/合并）
- **降级**：streaming失败自动回退openpyxl传统路径；用户可streaming=False强制传统路径
- **关键发现**：calamine把整数读成浮点数（2→2.0），需要数值标准化比较

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

---
*归档时间：2026-03-27 第101轮*

> 这些决策记录了项目早期的重要设计选择，留存供历史参考。