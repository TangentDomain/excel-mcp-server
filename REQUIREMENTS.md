# ExcelMCP 需求池

> 来源：客观评价 → 需求 → 实现 → 验证 → 关闭
> 每轮从 OPEN 中按优先级挑选，完成后移到 DONE

## 进化线路图

| 线路 | REQ | 优先级 | 关注点 | 当前状态 |
|------|-----|--------|--------|----------|
| ⚡ 性能优化 | REQ-015 | P1 | openpyxl I/O瓶颈、大表查询、写入批量优化 | OPEN |
| 🎯 工具描述 | REQ-006 | P1 | AI选工具准确率、参数命名一致性 | 持续迭代 |
| 🔧 工程治理 | REQ-010 | P1 | 代码复杂度/重复/依赖健康/架构 | 持续迭代 |
| 📈 可观测性 | REQ-013 | P2 | 结构化日志/工具频率/错误分类 | 部分完成 |
| 🌍 兼容性 | REQ-012 | P1 | 多客户端/跨平台/CI矩阵 | 大部分完成 |
| 🏗️ 基建 | REQ-008 | P2 | git worktree隔离 | OPEN |

## OPEN（待实现）

### REQ-015 [P1] 性能优化
- **来源**：CEO要求 — 性能优化要作为独立线持续跟进
- **描述**：openpyxl I/O是最大瓶颈（benchmark显示50行读1.6s vs pandas缓存30ms，差距50倍）
- **关注点**：
  - **读取优化** — excel_query/DESCRIBE等只读场景已用pandas缓存，但首次加载仍依赖openpyxl
  - **写入优化** — update_range/update_query逐行写入，大表批量修改很慢
  - **大表支持** — 1000+行配置表的查询/搜索/修改体验
  - **缓存策略** — LRU(max=10)是否足够？缓存失效策略是否合理？
  - **benchmark持续** — 每轮跑benchmark对比，检测性能退化
- **性能基线（第29轮benchmark）**：
  - SQL SELECT 50行: ~30ms | SQL GROUP BY 100行: ~80ms
  - openpyxl读取50行: ~1610ms | 读取100行: ~5742ms
  - 搜索精确/模糊/正则: 14-19ms
  - 写入20行×10列: ~23ms
- **已完成**：
  - ✅ 第6轮：DataFrame LRU缓存(max=10)，重复查询30-100倍提速
  - ✅ 第29轮：独立benchmark脚本，15项指标+历史对比
  - ✅ 第30轮：缓存跨调用共享（单例引擎，修复缓存从未复用的bug）
  - ✅ 第32轮：_load_worksheets批量双行表头检测（2N+1→N+1次文件打开）
  - ✅ 第32轮：excel_describe_table单次遍历所有列（N×M→M行I/O）
  - ✅ 第32轮：DESCRIBE内存优化（类型推断限制前100个值）
  - ✅ 第33轮：python-calamine替代openpyxl读取路径（get_range 1.6s→0.7ms，2300x提速）
  - ✅ 第47轮：智能追加优化（insert_mode目标行>末尾时跳过O(n)行移动+公式遍历）
- **后续方向**：
  - 写入场景：openpyxl的write_only模式或批量写入优化
  - 缓存预热/预加载策略
  - benchmark脚本适配calamine（性能基线已大幅变化）
- **验收**：get_range<50ms ✅（实际0.7ms），大表(1000行)操作流畅
- **状态**：OPEN

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
  - ✅ 第28轮：CI优化 — test extra替代dev（减少~20个不必要的包）+ pip缓存
  - ✅ 第28轮：CI全矩阵通过（含Windows 3.13），之前main失败为临时性问题
  - ✅ 第28轮：mypy overrides清理fastmcp残留
  - ✅ 第34轮：excel_writer.py 10个方法添加workbook.close()（修复文件句柄泄漏）
  - ✅ 第34轮：_create_temp_workbook移除重复代码块（公式计算不再加载文件2次）
  - ✅ 第34轮：eval()替换为_safe_eval_expr() AST白名单验证
- **验收**：至少在2个不同MCP客户端中验证通过，CI跑3.10和3.13
- **状态**：OPEN（CI矩阵已就绪且通过，多客户端验证待实际测试）

### REQ-013 [P2] 可观测性
- **来源**：运维需要 — 发布后需要监控工具运行状态
- **描述**：建立性能基线和工具使用追踪
- **关注点**：
  - **性能基线** — 每轮自动跑benchmark（大表查询/批量写入/搜索），对比历史数据检测退化
  - **工具使用频率** — ✅ excel_server_stats工具 + ToolCallTracker装饰器，实时统计
  - **结构化日志** — 统一日志格式，方便排查问题
  - **错误分类统计** — 按错误类型统计（文件不存在/格式错误/权限问题等）
- **已完成**：
  - ✅ 第29轮：独立benchmark脚本 `scripts/benchmark.py`（15项性能指标，SQL/读/写/搜索）
  - ✅ 第29轮：支持 `--quick` 快速模式（~30秒）和完整模式（含大表）
  - ✅ 第29轮：支持 `--compare` 与历史结果对比，检测性能退化（>30%变慢告警）
  - ✅ 第29轮：JSON报告输出，按类别汇总（sql/read/search/write）
  - ✅ 第35轮：ToolCallTracker + @_track_call装饰器 — 自动追踪42个工具调用次数/耗时/错误率
  - ✅ 第35轮：excel_server_stats工具 — 运行时统计查询（第42个工具）
  - ✅ 第36轮：JsonFormatter结构化JSON日志 — EXCEL_MCP_JSON_LOG=1激活，单行JSON输出
- **验收**：benchmark自动运行并输出对比报告，日志格式统一
- **状态**：OPEN（性能基线+追踪器+JSON日志已完成，错误分类统计待后续实现）

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
  - ✅ 第27轮：错误响应格式标准化（formatter.py归一化error→message，_format_error_result改用message）
  - ✅ 第30轮：SQL引擎缓存跨调用共享（模块级单例_get_engine()，修复缓存从未跨调用复用的隐性bug）
  - ✅ 第30轮：清理4个重复docstring（convert/restore/import/merge，server.py -12行）
  - ✅ 第34轮：excel_writer.py 10个方法添加workbook.close()（修复文件句柄泄漏）
  - ✅ 第34轮：_create_temp_workbook移除重复代码块（公式计算不再加载文件2次）
  - ✅ 第34轮：eval()替换为_safe_eval_expr() AST白名单验证（拒绝__import__等危险调用）
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

### REQ-016 [P0] SQL引擎增强
- **来源**：用户实测反馈 — 测试报告发现3个不支持功能
- **描述**：SQL查询引擎新增能力
- **已完成**（v1.0.16, 第46轮）：
  1. ✅ 子查询：WHERE col IN/NOT IN (SELECT ...)、标量子查询（WHERE col > (SELECT AVG...)）
  2. ✅ CASE WHEN表达式：CASE WHEN 条件 THEN 值 ELSE 默认 END
  3. ✅ COALESCE/IFNULL：空值替换
  4. ✅ EXISTS：WHERE EXISTS (SELECT ...)（含关联子查询）
  5. ✅ LEFT JOIN NULL处理bug修复
  6. ✅ 字符串函数(9个)：UPPER、LOWER、TRIM、LENGTH、CONCAT、REPLACE、SUBSTRING、LEFT、RIGHT
  7. ✅ CTE (WITH ... AS ...)：支持单CTE和多CTE链式引用
  8. ✅ HAVING聚合别名解析修复（COUNT(*) → cnt）
  9. ✅ FROM子查询不支持时清晰错误提示
- **未完成**（留待后续）：
  10. ❌ UNION/UNION ALL：合并查询结果（需要跨DataFrame concat逻辑）
  11. ❌ 窗口函数：ROW_NUMBER、RANK（复杂度高，游戏场景少见）
  12. ❌ RIGHT/FULL/CROSS JOIN（游戏场景极少使用）
  13. ❌ 跨文件JOIN：类似数据库跨库查询，`SELECT * FROM 技能表@file1.xlsx s JOIN 掉落表@file2.xlsx d ON s.技能ID = d.技能ID`
- **测试**：16个新测试（test_sql_enhanced.py），779全通过
- **验收标准**：每项至少2个测试用例 ✅ | 更新文件头支持列表 ✅ | 不支持项目有替代提示 ✅
- **状态**：IN_PROGRESS（核心功能已完成，UNION/窗口函数/扩展JOIN留待后续）

### REQ-017 [P1→DONE] Streamable HTTP + SSE传输模式
- **来源**：竞品分析 — haris-musa支持三重传输，我们仅stdio
- **描述**：暴露FastMCP原生支持的三种传输模式（stdio/sse/streamable-http）
- **验收标准**：命令行参数选择传输模式，不影响现有stdio功能
- **状态**：DONE ✅（main()新增--stdio/--sse/--streamable-http/--mount-path参数）

### REQ-018 [P1] Upsert（INSERT ON DUPLICATE KEY UPDATE）
- **来源**：数据库能力发散 — 类比MySQL的UPSERT
- **描述**：写入时自动判断ID是否存在，存在则更新，不存在则插入。策划合并配置高频操作。
- **语法参考**：`UPSERT INTO 技能表 VALUES (1001, '火球术', ...) ON DUPLICATE KEY UPDATE 伤害=VALUES(伤害)`
- **验收标准**：单条upsert + 批量upsert，至少3个测试
- **状态**：OPEN

### REQ-019 [P1] 批量INSERT
- **来源**：数据库能力发散 — 类比 `INSERT INTO ... VALUES (...), (...), (...)`
- **描述**：一次插入多行数据，当前只能逐行写入。策划批量导入几十条配置时效率提升显著。
- **验收标准**：批量INSERT + 与现有逐行写入兼容，至少3个测试
- **状态**：OPEN

### REQ-020 [P2] View（命名查询/保存的SQL）
- **来源**：数据库能力发散 — 类比 `CREATE VIEW v AS SELECT ...`
- **描述**：保存常用SQL查询为命名视图，后续直接调用视图名获取结果。避免每次重写复杂SQL。
- **语法参考**：`CREATE VIEW 高伤技能 AS SELECT * FROM 技能表 WHERE 伤害 > 200`，之后 `SELECT * FROM 高伤技能`
- **验收标准**：创建/调用/删除视图，视图跟随文件生命周期，至少3个测试
- **状态**：OPEN

### REQ-021 [P2] 写入校验（约束体系）
- **来源**：数据库能力发散 — 类比 FK/Unique/Check/Enum 约束
- **描述**：写入数据时自动校验：
  1. **FK约束**：引用其他表的ID是否存在（如怪物.掉落装备ID → 装备表.装备ID）
  2. **Unique约束**：ID列不允许重复
  3. **Check约束**：数值范围校验（伤害不能为负数、冷却不能为0）
  4. **Enum约束**：品质只能填普通/稀有/史诗/传说
  5. **Default值**：新增行自动填充默认值
- **验收标准**：每种约束至少2个测试，违规时返回清晰错误信息
- **状态**：OPEN

### REQ-022 [P2] Auto Increment（自增ID）
- **来源**：数据库能力发散 — 类比 AUTO_INCREMENT / SERIAL
- **描述**：新增行时自动分配下一个可用ID，策划不用手动维护ID连续性。
- **验收标准**：单行自增 + 批量自增 + 指定起始值，至少3个测试
- **状态**：OPEN
