# ExcelMCP 需求池

> 来源：客观评价 → 需求 → 实现 → 验证 → 关闭
> 每轮从 OPEN 中按优先级挑选，完成后移到 DONE

## OPEN（待实现）

### REQ-012 [P1] 兼容性验证
- **来源**：发布前必须验证 — 用户会用各种客户端和环境
- **描述**：确保在不同MCP客户端和平台上都能正常工作
- **关注点**：
  - **多客户端** — OpenClaw/Claude Desktop/Cursor/Windsurf 都能正常调用
  - **跨平台路径** — Windows反斜杠、Mac空格路径、Linux符号链接
  - **Python版本** — 3.10/3.11/3.12/3.13都能跑
  - **依赖最小化** — 只依赖必要的包，不引入重量级依赖
- **已完成**：
  - ✅ 第24轮：包元数据完善（license/authors/keywords/classifiers/urls）
  - ✅ 第24轮：移除死依赖fastmcp（代码从未import，只使用mcp.server.fastmcp）
  - ✅ 第24轮：修复误导性错误提示（pip install fastmcp → pip install mcp）
  - ✅ 第24轮：清理mypy配置中formulas/xlwings引用
  - ✅ 第24轮：验证18个模块全部正常import
  - ✅ 第24轮：验证入口函数main()正常工作
  - ✅ 第26轮：GitHub Actions CI（4 Python版本 × 3 OS矩阵测试）
  - ✅ 第26轮：CLI --version/-v 支持
- **验收**：至少在2个不同MCP客户端中验证通过，CI跑3.10和3.13
- **状态**：OPEN（多客户端验证待实际测试，CI矩阵已就绪）

### REQ-013 [P2] 可观测性
- **来源**：运维需要 — 发布后需要监控工具运行状态
- **描述**：建立性能基线和工具使用追踪
- **关注点**：
  - **性能基线** — 每轮自动跑benchmark（大表查询/批量写入/搜索），对比历史数据检测退化
  - **工具使用频率** — 记录哪些工具最常用、哪些没人用，指导后续优化方向
  - **结构化日志** — 统一日志格式，方便排查问题
  - **错误分类统计** — 按错误类型统计（文件不存在/格式错误/权限问题等）
- **验收**：benchmark自动运行并输出对比报告，日志格式统一
- **状态**：OPEN

### REQ-010 [P1] 工程治理
- **来源**：CEO要求 — 项目不仅要功能完善，工程本身也要健康
- **描述**：关注代码质量和工程健康度，不只是功能堆叠
- **关注维度**：
  - **代码复杂度** — 新增代码是否增加圈复杂度？单个函数是否过长？是否有深层嵌套？
  - **代码重复** — 不同工具间是否有重复逻辑？应提取公共方法
  - **测试质量** — 测试是真在验证行为，还是只为凑覆盖率？边界case覆盖了吗？
  - **错误处理一致性** — 41个工具的错误信息格式是否统一？失败时是否有用？
  - **依赖健康** — 依赖版本是否最新？有没有已知漏洞？
  - **性能回归** — 每轮改进是否引入性能退化？大表查询时间是否稳定？
  - **架构清晰度** — server.py是否越来越胖？API层和核心层是否职责清晰？
- **已完成**：
  - ✅ 第22轮：清理重构残留旧包+3个死依赖+8个测试import修复（-1921行）
  - ✅ 第23轮：excel_get_operation_history bug修复（Optional参数条件验证）
  - ✅ 第23轮：excel_search_directory安全补漏（添加_validate_path）
  - ✅ 第23轮：MCP说明与代码同步（JOIN功能遗漏2轮）
  - ✅ 第24轮：移除死依赖fastmcp + 修复误导错误提示 + 清理mypy残留引用
  - ✅ 第24轮：包元数据完善（license/authors/keywords/classifiers/urls）
  - ✅ 第26轮：发现错误响应格式不一致（_validate_path用message，ExcelOperations用error），记录待修复
- **验收**：每轮评价中包含工程治理评估，发现问题立即修复或建需求
- **状态**：OPEN（持续迭代，不关闭）

### REQ-008 [P2] 定时任务使用git worktree隔离测试
- **来源**：CEO建议 — 避免测试影响主工作目录
- **描述**：定时任务在 worktree 中进行开发和测试，测试通过后再合并回 develop/main
- **好处**：多个子代理可并行开发、测试不会污染主工作目录、失败可快速丢弃
- **验收**：定时任务自动创建worktree→开发→测试→合并→清理worktree
- **状态**：OPEN

### REQ-006 [P1] 工具描述持续优化
- **来源**：AI视角评价 — 工具描述直接影响AI选工具的准确率
- **描述**：根据AI实际使用反馈，持续优化工具描述的清晰度和场景引导
- **已完成的优化**：
  - ✅ 去掉Args/Returns/Example开发者文档（-583行）
  - ✅ 高频工具加场景引导和优先级提示
  - ✅ 统一格式，平均48字/工具
  - ✅ 第23轮：10个易混淆工具加交叉引用（search↔search_directory, get_headers↔get_sheet_headers, preview↔assess, compare_files↔compare_sheets）
  - ✅ 第23轮：5个工具加场景引导（list_sheets/create_file/find_last_row等）
- **后续方向**：
  - MCP调用准确性监控：每轮验证时记录AI是否第一次就选对工具
  - 常见选错场景：该用query却用get_range、该用update_query却用update_range、该用describe_table却用get_headers
  - 发现选错 → 针对性加描述提示或互斥说明
  - 低频但易混淆的工具加优先级排序提示
  - 工具描述A/B测试：对比优化前后AI选工具准确率
- **已知问题（第24轮发现）**：
  - excel_search 没有 search_type 参数，搜索模式通过 use_regex 布尔值控制
  - excel_compare_sheets 用 file1_path/file2_path 而非 file_path（与多数工具不一致）
- **第25轮已修复**：
  - ✅ excel_query/excel_update_query 参数名统一问题 → update_expression（消除混淆）
- **验收**：AI选工具错误率降低，高频场景不再选错
- **状态**：OPEN（持续迭代，不关闭）

## DONE（已完成）

### REQ-000 SQL查询引擎
- **来源**：初始评估
- **描述**：支持中文SQL查询游戏配置表
- **状态**：DONE ✅（第3-13轮）

### REQ-000 双行表头自动识别
- **来源**：游戏场景适配
- **描述**：自动检测第1行中文描述+第2行英文字段名
- **状态**：DONE ✅（第3轮）

### REQ-001 [P0→P2→DONE] 游戏领域函数
- **来源**：策划视角评价 — DPM/DPS是高频需求
- **状态**：DONE ✅（通过README策划教程中DPM数学表达式示例满足）

### REQ-002 [P1→DONE] 增量更新（WHERE条件批量修改）
- **状态**：DONE ✅（第15轮，excel_update_query第41个工具，14个测试通过）

### REQ-003 [P1→DONE] JOIN支持（跨表关联查询）
- **状态**：DONE ✅（第16轮，INNER/LEFT JOIN，15个测试通过）

### REQ-004 [P2→DONE] 查询结果导出（JSON/CSV）
- **状态**：DONE ✅（第14轮，JSON/CSV/TABLE三种格式，4个测试通过）

### REQ-005 [P2→DONE] excel_describe_table支持中文列名查询
- **状态**：DONE ✅（第3轮，双行表头自动识别+中英文对照）

### REQ-007 [P1→DONE] README文档同步更新
- **状态**：DONE ✅（第17轮，README已同步：JOIN/UPDATE文档、badge 1271测试/41工具）

### REQ-011 [P0→DONE] 安全加固
- **状态**：DONE ✅（第18-19轮，SecurityValidator路径穿越/大小/公式验证+文件锁+临时文件清理+33个测试）

### REQ-009 [P1→DONE] UPDATE事务保护
- **来源**：自我进化评价 — 写入失败可能损坏文件
- **描述**：excel_update_query写入前创建临时备份，失败自动回滚
- **状态**：DONE ✅（第17轮，shutil.copy2备份+回滚+1个测试验证）
