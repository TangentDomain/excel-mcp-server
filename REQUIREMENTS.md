# ExcelMCP 需求池

> 来源：客观评价 → 需求 → 实现 → 验证 → 关闭
> 每轮从 OPEN 中按优先级挑选，完成后移到 DONE

## 进化线路图

| 线路 | REQ | 优先级 | 关注点 | 当前状态 |
|------|-----|--------|--------|----------|
| ⚡ 性能优化 | REQ-015 | P1 | openpyxl I/O瓶颈、大表查询、写入批量优化 | OPEN |
| 🎯 工具描述 | REQ-006 | P1 | AI选工具准确率、参数命名一致性 | 持续迭代 |
| 🔧 工程治理 | REQ-010 | P1 | 代码复杂度/重复/依赖健康/架构 | 持续迭代 |
| 📈 可观测性 | REQ-013 | P2 | 结构化日志/工具频率/错误分类 | DONE ✅ |
| 🌍 兼容性 | REQ-012 | P1 | 多客户端/跨平台/CI矩阵 | 大部分完成 |
| 🏗️ 基建 | REQ-008 | P2 | git worktree隔离 | DONE ✅ |
| 🤖 AI体验优化 | REQ-025 | P1 | 返回值统一/错误结构化/大结果截断/重复工具合并 | 持续迭代 |
| 📚 文档与门面 | REQ-026 | P1 | README优化/GitHub门面/使用示例/对比文档 | 持续迭代 |

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
  - 读取(50行): ~0.7ms (calamine) | 读取(100行): ~1ms (calamine)
  - 搜索精确/模糊/正则: ~1ms (calamine)
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
  - ✅ 第47轮：智能追加优化+excel_writer.py/advanced_sql_query.py/excel_operations.py内联import提升
  - ✅ 第48轮：excel_operations.py 32处冗余内联import全面清理（-29行），excel_writer.py 5处（-5行）
  - ✅ 第64轮：DRY消除+分发表统一+ORDER BY重构（净-31行）
  - ✅ 第65轮：COALESCE向量化（combine_first替代5处逐行循环）+CASE WHEN DRY（25行→4行）+_get_row_value数字字面量bugfix
- **验收**：每轮评价中包含工程治理评估，发现问题立即修复或建需求
- **状态**：OPEN（持续迭代，不关闭）

### REQ-008 [P2] 定时任务使用git worktree隔离测试
- **来源**：CEO建议 — 避免测试影响主工作目录
- **描述**：定时任务在 worktree 中进行开发和测试，测试通过后再合并回 develop/main
- **好处**：多个子代理可并行开发、测试不会污染主工作目录、失败可快速丢弃
- **验收**：定时任务自动创建worktree→开发→测试→合并→清理worktree
- **状态**：DONE ✅（cron prompt已内置worktree工作流，每轮自动创建feature branch + worktree）

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

### REQ-013 [P2→DONE] 可观测性
- **状态**：DONE ✅（第29/35/36/63轮，benchmark+tracker+JSON日志+错误分类统计，27个测试）

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
- **状态**：DONE ✅（v1.0.16, 第46轮，9项核心功能全部实现，16个新测试通过）

### REQ-027 [P2] SQL引擎增强（剩余项）
- **来源**：REQ-016未完成项拆分
- **描述**：
  1. ~~UNION/UNION ALL：合并查询结果~~ ✅（第55轮实现）
  2. 窗口函数：ROW_NUMBER、RANK、DENSE_RANK（复杂度高，游戏场景少见）
  3. ~~RIGHT/FULL/CROSS JOIN（游戏场景极少使用）~~ ✅（第58轮实现）
  4. 跨文件JOIN：`SELECT * FROM 技能表@file1.xlsx s JOIN 掉落表@file2.xlsx d ON s.技能ID = d.技能ID`
- **验收标准**：每项至少2个测试，更新文件头支持列表
- **状态**：OPEN（仅剩跨文件JOIN）
- **已完成**：
  - ✅ 第55轮：UNION/UNION ALL（递归提取SELECT+concat+去重+ORDER BY+LIMIT，13个测试）
  - ✅ 第58轮：RIGHT/FULL/CROSS JOIN（17个测试，SQL功能37→38）
- **v1.0.19修复**：EXISTS关联子查询无表限定符列引用re.sub参数顺序bug

### REQ-017 [P1→DONE] Streamable HTTP + SSE传输模式
- **来源**：竞品分析 — haris-musa支持三重传输，我们仅stdio
- **描述**：暴露FastMCP原生支持的三种传输模式（stdio/sse/streamable-http）
- **验收标准**：命令行参数选择传输模式，不影响现有stdio功能
- **状态**：DONE ✅（main()新增--stdio/--sse/--streamable-http/--mount-path参数）

### REQ-018 [P1→DONE] Upsert（INSERT ON DUPLICATE KEY UPDATE）
- **来源**：数据库能力发散 — 类比MySQL的UPSERT
- **描述**：写入时自动判断ID是否存在，存在则更新，不存在则插入。策划合并配置高频操作。
- **验收标准**：单条upsert + 批量upsert，至少3个测试 ✅（13个测试）
- **实现**：ExcelManager.upsert_row核心层，按键列查找→update/insert双路径，支持双行表头
- **状态**：DONE ✅（第54轮，excel_upsert_row第45个工具，13个测试通过）

### REQ-019 [P1→DONE] 批量INSERT
- **来源**：数据库能力发散 — 类比 `INSERT INTO ... VALUES (...), (...), (...)`
- **描述**：一次插入多行数据，当前只能逐行写入。策划批量导入几十条配置时效率提升显著。
- **验收标准**：批量INSERT + 与现有逐行写入兼容，至少3个测试 ✅（6个测试）
- **实现**：ExcelManager.batch_insert_rows核心层，按列名映射批量写入，未知列自动忽略
- **状态**：DONE ✅（第54轮，excel_batch_insert_rows第46个工具，6个测试通过）

### REQ-023 [P2→DONE] 复制Sheet
- **来源**：能力盘点 — 策划经常复制表做变体（如副本版怪物表、活动版装备表）
- **描述**：复制指定Sheet（含数据和格式）到同文件，支持重命名和位置指定
- **验收标准**：同文件复制 + 自定义名称 + 名称冲突自动递增 + 重命名参数，至少3个测试 ✅（6个测试）
- **实现**：openpyxl copy_worksheet + move_sheet调整位置 + _normalize_sheet_name规范化
- **状态**：DONE ✅（第53轮，excel_copy_sheet第43个工具，6个测试通过）

### REQ-024 [P2→DONE] 重命名列
- **来源**：能力盘点 — 改列名是常见操作（统一命名规范、适配跨表JOIN字段名）
- **描述**：重命名指定Sheet的列名（修改表头单元格值）
- **验收标准**：单列重命名 + 列名不存在时报错（提示实际列名） + header_row支持 + 至少3个测试 ✅（7个测试）
- **实现**：遍历表头行匹配old_header精确值 + 修改单元格 + 验证持久化
- **状态**：DONE ✅（第53轮，excel_rename_column第44个工具，7个测试通过）

### REQ-025 [P1] AI体验优化线（持续迭代，不关闭）
- **来源**：产品定位复盘 — MCP工具的用户是AI不是人类，需要优化AI使用体验
- **关注点**：
  1. **MCP instructions优化**：FastMCP实例的instructions字段，AI首次连接时看到的"自我介绍"，决定AI对MCP整体能力的理解
  2. **工具docstring优化**：每个工具的描述是AI选择工具的唯一依据，直接影响选工具准确率（REQ-006的执行层）
  3. **返回值结构统一**：所有工具返回统一的JSON结构（success/error/data/meta），降低AI解析成本
  4. **错误信息结构化**：SQL报错返回`{error, suggestion, available_columns, original_sql}`，AI能直接用suggestion重试
  5. **大结果自动截断**：查询结果超过阈值（如200行）时自动截断+提示"建议加WHERE/LIMIT"，保护AI上下文
  6. **合并重复工具**：~~preview_operation→assess_data_impact（第79轮合并，detailed参数）~~，get_headers/get_sheet_headers保持独立（不同用途）
- **已完成**：
  - ✅ 第51轮：MCP instructions全面更新（6项过时信息修正，SQL功能29→35项）
  - ✅ 第51轮：大结果自动截断（>500行截断为前500行，returned_rows/truncated字段）
- **验收标准**：每个子项独立验收，MCP验证中AI选工具准确率作为核心指标
- **状态**：OPEN（持续迭代，不关闭）

### REQ-026 [P1] 文档与门面优化线（持续迭代，不关闭）
- **来源**：竞品分析 — haris-musa的README和GitHub门面是获取用户的关键，我们差距明显
- **关注点**：
  1. **README优化**：30秒上手教程前置、游戏场景示例突出、安装badge醒目、中英文同步
  2. **GitHub门面**：About描述、Topics、Homepage链接持续更新，保持与版本同步
  3. **使用示例**：游戏配置表典型场景的SQL示例和工具调用示例（给配置MCP的开发者看）
  4. **对比文档**：与haris-musa等竞品的功能对比表，突出SQL引擎+性能+游戏垂直差异化
  5. **Changelog**：版本更新日志，让用户看到活跃度
- **验收标准**：每次发布新版本时检查README和门面是否同步
- **状态**：OPEN（持续迭代，不关闭）
