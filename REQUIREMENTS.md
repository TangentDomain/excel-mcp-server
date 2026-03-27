# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/ROADMAP)，已归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求

### REQ-025 [P1] AI体验优化线（持续迭代，不关闭）
- **关注点**：instructions优化（已完成）、docstring优化（持续）、返回值统一（进行中）、错误信息结构化、大结果截断（已完成）、合并重复工具（preview/assess已完成，get_headers待合并）

### REQ-026 [P1] 文档与门面优化线（持续迭代，不关闭）✅ 第111轮完成
- **关注点**：README 30秒上手教程、GitHub门面、使用示例、竞品对比、Changelog
- **完成**：完善30秒上手教程，创建examples/完整游戏场景示例，更新README引用

### REQ-028 [P1] FROM子查询支持 ✅
- **描述**：支持 `SELECT * FROM (SELECT ...) AS t WHERE ...`
- **验收**：基础FROM子查询 + WHERE过滤 + JOIN结果子查询 + 嵌套子查询拒绝 + 空结果 + DISTINCT + 无别名，12个测试全通过

### REQ-015 [P1] 性能优化（写入）
- **描述**：openpyxl write_only模式，减少写入内存和时间
- **完成**：v1.6.0，所有修改操作支持streaming参数，copy-modify-write方案

### REQ-012 [P1] 兼容性验证
- **描述**：多客户端实际测试（Cursor、Claude Desktop等）

### REQ-006 [P1] 工具描述持续优化（持续迭代，不关闭）- ✅ 第108轮完成
- **描述**：持续优化工具描述的一致性和完整性
- **完成**：中英文README文档同步，44个工具游戏场景描述完整，统一返回格式

<<<<<<< HEAD
### REQ-029 [P0] JOIN表别名列引用 + describe_table流式写入后崩溃 - ✅ v1.6.6
=======
### REQ-031 [P2] CI Node.js 20弃用警告
- **来源**：GitHub Actions 通知（2026-03-27）
- **问题**：actions/checkout@v4 和 actions/setup-python@v5 运行在 Node.js 20，2026年9月16日将被移除
- **修复**：升级 actions 版本或加 `FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true`
- **截止**：2026-09-16

### REQ-030 [P0] SQL引擎边界情况（3项阻断性Bug）
- **来源**：本轮MCP真实验证后发现（2026-03-27），影响核心功能可用性
- **Bug 1**：`MAX(表达式)` 不支持 `MAX(a + b + c)` 形式，只支持 `MAX(列名)`，聚合函数内多列表达式计算失败
- **Bug 2**：标量子查询列名解析错误——`WHERE sub.职业` 被解析为原始表列名而非子查询别名，关联条件失效
- **Bug 3**：`LEFT JOIN + IS NULL` 匹配产生的NULL行被过滤掉（sqlglot转换问题），建议用 `NOT IN` 子查询替代
- **验收**：`SELECT MAX(攻击力+加成) FROM 表` 返回正确聚合值；标量子查询别名正确解析；LEFT JOIN IS NULL 返回无匹配行
- **优先级**：提升至P0，因为影响用户正常使用聚合函数和子查询

### REQ-029 [P0] JOIN表别名列引用 + describe_table流式写入后崩溃 - ✅ v1.6.7
>>>>>>> develop
- **Bug 1**：JOIN后SQL表别名不生效，`r.名称`不被识别，pandas JOIN后列名加`_x`/`_y`后缀，SQL别名未映射
- **Bug 2**：`batch_insert_rows`/`delete_rows` streaming写入后，openpyxl `read_only=True`模式 `ws.max_row`返回None，导致`describe_table`崩溃
- **验收**：`SELECT r.名称 FROM 角色 r JOIN 技能 s ON r.职业 = s.职业限制` 返回正确列名；streaming写入后 describe_table 正常返回
- **来源**：主会话MCP真实调用验证（2026-03-27）
- **修复**：增强`_expression_to_column_reference`别名映射（5层回退），添加`_join_column_mapping`记录JOIN列映射；`describe_table`添加try/except处理`max_row=None`
- **完成时间**：2026-03-27，第117轮，v1.6.6发布

### REQ-010 [P1] 工程治理（持续迭代，不关闭）- ✅ 第106轮完成
- **描述**：代码质量、测试覆盖、文档完整性、项目结构、安全性优化
- **完成**：工程健康评估85/100，新增安全验证工具，发布v1.5.4