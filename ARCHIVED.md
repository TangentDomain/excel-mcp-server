# ExcelMCP 已归档需求

> 已完成或取消的需求，仅供参考。

## DONE

### REQ-028 [P1] FROM子查询支持 ✅
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过

- REQ-000 SQL查询引擎 ✅（第3-13轮）
- REQ-000 双行表头自动识别 ✅（第3轮）
- REQ-001 游戏领域函数 ✅（README DPM数学表达式示例满足）
- REQ-002 增量更新（WHERE条件批量修改）✅（第15轮）
- REQ-003 JOIN支持（跨表关联查询）✅（第16轮）
- REQ-004 查询结果导出（JSON/CSV）✅（第14轮）
- REQ-005 excel_describe_table中文列名查询 ✅（第3轮）
- REQ-007 README文档同步更新 ✅（第17轮）
- REQ-008 git worktree隔离 ✅（cron prompt内置工作流）
- REQ-009 UPDATE事务保护 ✅（第17轮）
- REQ-011 安全加固 ✅（第18-19轮）
- REQ-013 可观测性 ✅（第29-63轮，benchmark+tracker+JSON日志+错误分类）
- REQ-016 SQL引擎增强（核心9项）✅（第46轮）
- REQ-017 Streamable HTTP + SSE传输 ✅
- REQ-018 Upsert ✅（第54轮）
- REQ-019 批量INSERT ✅（第54轮）
- REQ-023 复制Sheet ✅（第53轮）
- REQ-024 重命名列 ✅（第53轮）
- REQ-027 跨文件JOIN ✅（第80轮）

## CANCELLED

- REQ-020 View（命名查询）— SQL查询本身就能保存为文本
- REQ-021 写入校验（约束体系）— SQL引擎已覆盖
- REQ-022 Auto Increment — 策划手动管理ID更可控
### REQ-028 [P1] FROM子查询支持 ✅
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过


### REQ-028 [P1] FROM子查询支持 ✅
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过

### REQ-015 [DONE] 性能优化（写入） ✅
- **描述**：openpyxl write_only模式，减少写入内存和时间
- **完成**：v1.5.3，所有修改操作支持streaming参数，copy-modify-write方案

### 2026-03-27 新增归档
- REQ-010 文档与门面优化 ✅（R151）- 项目文档和外部接口的持续优化，同步版本数据(v1.6.26/1159测试)，补全CHANGELOG
- REQ-032 P0 bug修复 ✅（R146）
- REQ-006 工具描述持续优化 ✅（R147）- 44/44工具100%优秀，所有工具描述完整
- REQ-029 JOIN表别名 + describe_table崩溃 ✅（R142）
- REQ-015 streaming写入后读取工具验证 ✅（R142）
- REQ-025 docstring优化 ✅（R143）- 44/44函数100%达标
- REQ-026 文档与门面优化 ✅（R156）- 版本v1.6.29统一，44工具，1176测试，功能验证通过，文档完全同步
