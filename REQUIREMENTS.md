# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/ROADMAP)，已归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求

### REQ-034 [P0] 代码整洁度清理
- **问题1**：根目录16个临时脚本（2301行），CI不引用，应该删除。保留mcp_verification.py和test_mcp_real.py，其余14个删除
- **问题2**：安全测试重复3个文件（test_security.py + test_security_features.py + test_req010_r67.py = 79个测试测同一件事），合并为1个
- **问题3**：Streaming测试6个文件101个测试，功能重叠（req015_streaming_read vs streaming_read_verify），合并
- **问题4**：test_req010_r67.py带轮次编号，是临时需求测试，应合并到对应功能文件后删除
- **验收**：根目录无临时脚本（保留.gitignore认可的文件）；无带轮次编号的测试文件；全量测试通过
- **来源**：CEO指示 + 项目健康度分析（2026-03-28）

### REQ-033 [P1] 游戏场景极限探索（每10轮执行1次）
- **性质**：子代理扮演游戏策划，用MCP工具完成真实的游戏配置表工作
- **核心**：每次自由发挥，创造性思考，做不同的事。在复杂操作中发现产品和SQL引擎的bug
- **怎么做**：
  1. 先构思一个游戏场景（每次不同，不重复）
  2. 用create_file建表，用get_range/insert/delete/update等操作数据，用SQL做复杂查询
  3. 操作要有挑战性：跨sheet关联、复杂条件、大数据量、边界值、嵌套操作
  4. 在过程中观察：哪里报错了、结果不符合预期、SQL执行有歧义、工具行为不一致
  5. 分析发现的问题，能修就修，不能修就记入REQUIREMENTS
- **重点**：是"用产品发现问题"，不是"跑测试验证功能"。像真实策划用MCP配表一样，越复杂越好
- **验收**：每次产出场景描述+操作过程+发现的问题（bug/体验问题/SQL边界）
- **来源**：CEO指示（2026-03-27）

### REQ-025 [P1] AI体验优化线（持续迭代，不关闭） ✅ 第131轮完成阶段性优化
- **关注点**：instructions优化（已完成）、docstring优化（持续，第131轮大幅提升）、返回值统一（进行中）、错误信息结构化、大结果截断（已完成）、合并重复工具（preview/assess已完成，get_headers待合并）
- **第131轮进展**：docstring质量评分提升200%（2个excellent → 6个excellent），重点优化excel_search_directory、excel_get_range、excel_update_range、excel_assess_data_impact

### REQ-026 [P1] 文档与门面优化线（持续迭代，不关闭）✅ 第165轮完成
- **关注点**：README 30秒上手教程、GitHub门面、使用示例、竞品对比、Changelog
- **完成**：完善30秒上手教程，创建examples/完整游戏场景示例，更新README引用，新增故障排除章节，同步版本信息

### REQ-028 [P1] FROM子查询支持 ✅
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过

### REQ-015 [P1] 性能优化（写入）
- **描述**：openpyxl write_only模式，减少写入内存和时间
- **完成**：v1.6.0，所有修改操作支持streaming参数，copy-modify-write方案

### REQ-012 [P1] 兼容性验证 ✅
- **描述**：多客户端实际测试（Cursor、Claude Desktop等）
- **完成**：2026-03-27，第128轮，100%兼容性通过（Cursor、Claude Desktop、VSCode MCP、流式写入）

### REQ-006 [P1] 工具描述持续优化（持续迭代，不关闭）- ✅ 第108轮完成
- **描述**：持续优化工具描述的一致性和完整性
- **完成**：中英文README文档同步，44个工具游戏场景描述完整，统一返回格式

### REQ-031 [P2] CI Node.js 20弃用警告 ✅
- **来源**：GitHub Actions 通知（2026-03-27）
- **问题**：actions/checkout@v4 和 actions/setup-python@v5 运行在 Node.js 20，2026年9月16日将被移除
- **修复**：actions/checkout@v5 + actions/setup-python@v6 + FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true
- **完成时间**：2026-03-27，第150轮

### REQ-030 [P0] SQL引擎边界情况（3项Bug） - ✅ v1.6.8
- **来源**：MCP真实验证后发现（2026-03-27）
- **Bug 1** ✅：`MAX(表达式)` 不支持多列表达式 → 新增表达式求值递归处理
- **Bug 2** ✅：SELECT子句中标量子查询不支持 → 新增Subquery处理分支
- **Bug 3** ✅：LEFT JOIN IS NULL → 经验证已正常工作，无需修复
- **完成时间**：2026-03-27，第120轮，v1.6.8发布

### REQ-032 [P0] MCP真实验证发现的新bug - ✅ v1.6.25
- **Bug 1** ✅：`excel_list_sheets`获取工作表列表失败，返回0个工作表
  - **现象**：实际文件有3个工作表，但API返回空列表
  - **错误信息**：`'<=' not supported between instances of 'int' and 'NoneType'`
  - **修复**：添加`_safe_float_comparison`函数处理None值，SQL WHERE条件比较不再crash
- **Bug 2** ✅：`excel_delete_rows`参数不匹配
  - **修复**：新增`condition`参数，支持SQL WHERE条件删除行（自动查询匹配行号，从后往前删除避免偏移）
- **Bug 3** ✅：`excel_batch_insert_rows`参数不匹配
  - **修复**：新增`insert_position`和`condition`参数，支持指定行号或SQL条件定位插入
- **完成时间**：2026-03-27，第146轮，v1.6.25发布

### REQ-029 [P0] JOIN表别名列引用 + describe_table流式写入后崩溃 - ✅ v1.6.7
- **Bug 1**：JOIN后SQL表别名不生效，`r.名称`不被识别，pandas JOIN后列名加`_x`/`_y`后缀，SQL别名未映射
- **Bug 2**：`batch_insert_rows`/`delete_rows` streaming写入后，openpyxl `read_only=True`模式 `ws.max_row`返回None，导致`describe_table`崩溃
- **验收**：`SELECT r.名称 FROM 角色 r JOIN 技能 s ON r.职业 = s.职业限制` 返回正确列名；streaming写入后 describe_table 正常返回
- **来源**：主会话MCP真实调用验证（2026-03-27）
- **修复**：增强`_expression_to_column_reference`别名映射（5层回退），添加`_join_column_mapping`记录JOIN列映射；`describe_table`添加try/except处理`max_row=None`
- **完成时间**：2026-03-27，第117轮，v1.6.6发布

### REQ-010 [P1] 工程治理（持续迭代，不关闭）- ✅ 第106轮完成
- **描述**：代码质量、测试覆盖、文档完整性、项目结构、安全性优化
- **完成**：工程健康评估85/100，新增安全验证工具，发布v1.5.4