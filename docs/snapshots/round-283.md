{
  "REQUIREMENTS": {
    "REQ-028": {
      "title": "excel_update_range insert_mode 默认值改为 false",
      "status": "DONE",
      "priority": "P0",
      "description": "excel_update_range 的 insert_mode 默认为 true，导致写入已有数据文件时会物理插入新行，顶走原始数据，造成文件损坏。CEO 在 MapEvent.xlsx 上踩坑 3 次，SVN 还原 3 次，浪费 30 分钟。",
      "acceptance_criteria": [
        "insert_mode 默认值从 true 改为 false（覆盖模式）",
        "docstring 必须写清楚每个参数的含义和默认值，不能只提一个参数",
        "系统 prompt 中的描述和代码实际默认值必须一致",
        "工具描述中明确说明 insert_mode 的两种行为差异，让 LLM 知道什么时候该用哪种",
        "811/811 测试全通过",
        "更新 README 中 excel_update_range 的说明",
        "新增专项测试（test/test_insert_mode.py）：",
        "  - 覆盖模式默认行为：不传 insert_mode 时，写入已有数据行应覆盖而非插入",
        "  - 覆盖模式精确验证：写入后目标单元格值正确，相邻行/列数据不变",
        "  - 插入模式显式开启：insert_mode=true 时，写入后原数据正确下移",
        "  - 插入模式验证：插入后行数增加，原数据行索引+1，内容不变",
        "  - 多列写入覆盖：一次写入多列数据，验证同行其他列不受影响",
        "  - 多行写入覆盖：一次写入多行数据，验证非目标行不受影响",
        "  - 边界场景：写入空文件、写入末尾行、写入单单元格",
        "  - 使用 MapEvent-damaged.xlsx 等真实配置表结构作为测试 fixture",
        "  - 所有测试必须断言具体单元格值，不能只检查行数"
      ],
      "constraints": [
        "这是破坏性行为变更，默认值改了但功能不变",
        "insert_mode=true 仍然保留，只是不再默认"
      ],
      "notes": "CEO 2026-03-30 生产事故。核心教训：写入已有数据的默认行为必须是覆盖，不是插入。插入是新功能，应该显式开启。\n\n事故根因分析：\n1. insert_mode 默认 True → 违反直觉\n2. docstring 只提了 preserve_formulas，完全没提 insert_mode → LLM 不知道这个参数\n3. 系统 prompt 写的'默认覆盖'和代码实际默认值 True 矛盾 → 文档骗人\n\n经验教训：docstring 必须覆盖所有参数，LLM 读 docstring 不读源码，描述和默认值不一致等于坑。"
    },
    "REQ-029": {
      "title": "工程强化：约束可机器验证（Codex 仓库分析反哺）",
      "status": "DONE",
      "priority": "P1",
      "description": "参考 OpenAI Codex 仓库工程体系分析，将'靠 LLM 自觉遵守'的规则升级为'靠脚本验证'的规则。核心心法：规则没有 enforcement 等于不存在。",
      "acceptance_criteria": [
        "【契约验证脚本】创建 scripts/lint_docstring_contract.py：",
        "  - 遍历所有 public 函数（非 _ 开头），解析函数签名获取参数列表",
        "  - 解析 docstring Args 段，提取已记录的参数名和默认值",
        "  - 对比：函数签名有但 docstring 没有的参数 → ERROR",
        "  - 对比：docstring 描述的默认值与代码实际默认值不一致 → ERROR",
        "  - 输出格式：函数名 | 参数名 | 问题类型（缺失/默认值不匹配/类型缺失）",
        "  - exit code：有 ERROR 则 1，否则 0",
        "【CI 集成】将 lint_docstring_contract.py 加入 CI workflow：",
        "  - 作为独立 job 或现有 lint job 的步骤",
        "  - 失败则阻断合并（不是 warning）",
        "【Conventional Commits】commit message 强制格式：",
        "  - 格式：[REQ-XXX] type: 简述，如 [REQ-028] fix: insert_mode 默认值改为 false",
        "  - type 范围：feat/fix/refactor/docs/test/chore/perf",
        "  - 写入 docs/RULES.md 作为代码规范",
        "【Pre-commit 关键文件保护】创建 scripts/pre_commit_check.sh：",
        "  - 检查 docs/RULES.md 存在且非空",
        "  - 检查 README.md 首行语言为中文",
        "  - 检查 README.en.md 首行语言为英文",
        "  - grep 冲突标记 <<<<<< → 有则 exit 1",
        "  - grep 敏感信息（AK/LTAI/ft522/admin/secret）→ 有则 exit 1",
        "  - 加入 .git/hooks/ 或 CI",
        "【变更检测精准测试】（长期）CI 分层策略：",
        "  - 检测 src/ 改动 → 跑全量测试",
        "  - 仅改 test/ → 跑对应测试文件",
        "  - 仅改 docs/ → 跳过测试",
        "  - 预估节省 30-40% CI 时间"
      ],
      "constraints": [
        "REQ-028 完成后再开始此需求",
        "脚本用 Python 标准库（ast模块），不引入新依赖",
        "CI 变更不能破坏现有 workflow",
        "不影响现有测试流程"
      ],
      "notes": "来源：2026-03-30 深度分析 openai/codex 仓库工程体系。核心发现：每条规则都有对应的强制执行机制（AGENTS.md→justfile→CI→tests），形成完整 enforcement chain。我们的元迭代体系已有这个意识（心跳检查、REQUIREMENTS.md），但可以更系统。\n\n优先级：契约验证脚本 > Pre-commit > Conventional Commits > CI 分层。"
    },
    "REQ-030": {
      "title": "API参数命名与常见术语对齐",
      "status": "DONE",
      "priority": "P2",
      "acceptance_criteria": [
        "create_chart 的 chart_type 支持 'column' 作为 'bar' 的别名",
        "create_pivot_table 的 agg_func 支持 'mean' 作为 'average' 的别名",
        "docstring 中同时标注常用别名",
        "全量测试通过"
      ],
      "constraints": [
        "向后兼容：原有参数值继续有效",
        "新增别名不影响现有功能"
      ],
      "notes": "来源：监工第3轮报告。'column' 和 'mean' 是 pandas/openpyxl 用户的常用术语，不支持会导致用户困惑。\n\n完成情况（2026-04-01）：\n✅ create_chart 的 chart_type 已支持 'column' 作为 'bar' 的别名（chart_map 中两者都映射到 BarChart）\n✅ docstring 已说明 'column' 作为 'bar' 的别名\n✅ 测试 test_create_column_chart 已存在并通过\n✅ create_pivot_table 函数已实现，支持 'mean' 作为 'average' 的别名\n✅ 全量测试通过，包含所有聚合函数（sum/count/average/mean/max/min/std/var）及其别名\n✅ docstring 中明确标注了别名关系"
    },
    "REQ-032": {
      "title": "性能优化：大型Excel文件处理提速（2GB+）",
      "status": "IN-PROGRESS",
      "priority": "P1",
      "type": "perf",
      "description": "处理大型Excel文件（2GB+）时遇到性能瓶颈，MCP调用响应缓慢。需要优化内存使用和数据处理速度，提升用户体验。",
      "acceptance_criteria": [
        "性能基准测试：创建 scripts/performance-benchmark.py",
        "测试文件：生成 2GB+ 的测试文件，包含百万级数据行",
        "测量当前性能：read_data_from_excel、write_data_to_excel 耗时",
        "内存使用优化：减少 pandas DataFrame 内存占用",
        "流式处理：实现大文件分块读取和写入",
        "缓存机制：为重复查询添加智能缓存",
        "并发处理：对批量操作使用多线程优化",
        "性能目标：提升 3-5倍处理速度，内存占用降低 50%",
        "MCP验证：使用真实的2GB+文件验证优化效果",
        "全量测试通过，确保不影响小文件处理"
      ],
      "constraints": [
        "向后兼容：不能影响现有API接口",
        "稳定性：优化后必须通过所有测试",
        "内存安全：不能因优化导致内存泄漏"
      ],
      "notes": "来源：用户反馈和监控数据。大型配置文件（如2GB+的游戏配置表）处理耗时过长，影响工作效率。\n\n优先级：P1，因为直接影响用户体验。\n\n完成情况（2026-04-01 第二轮）：\n✅ 性能基准测试脚本 scripts/performance-benchmark.py 已创建\n  - 支持 calamine/openpyxl/pandas 多引擎对比\n  - 支持 write_only 写入性能测试\n  - 支持 SQL 查询性能测试\n  - 支持 dtype 优化内存对比\n  - 自动生成 JSON 报告\n  - 基准数据：calamine ~100K行/s, openpyxl read_only ~80K行/s\n✅ 文件大小检测：ExcelReader 自动检测文件大小，>50MB 标记为大文件\n✅ 大文件优化读取：>50MB 文件的范围查询自动切换到 openpyxl read_only + iter_rows 按需加载\n✅ 文件大小元数据：list_sheets 和 get_range 返回 file_size_mb 和 is_large_file\n✅ DataFrame dtype 优化：_optimize_dtypes() 方法\n  - int64→int8/int16/int32/uint8/uint16/uint32 自动降级\n  - float64→float32 降级（精度足够）\n  - 低基数 object→category 转换（基数比<0.3）\n  - 自动应用于 _load_excel_data() 加载流程\n✅ 缓存增强：\n  - DataFrame 缓存容量 10→20\n  - 查询结果缓存容量 5→15\n  - 新增 _estimate_cache_memory_mb() 内存估算\n  - 新增 evict_cache_by_memory() 内存感知淘汰\n✅ 文件大小限制 100MB→2GB\n✅ 新增 14 个测试（test_performance_optimization.py），全部通过\n✅ 全量测试 839 passed，无回归\n⏳ 待做：2GB+ 真实文件MCP验证、多线程批量操作优化",
      "attempts": 0,
      "last_failure": ""
    },
    "REQ-053": {
      "title": "ORDER BY 浮点/混合类型列返回0行",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "ORDER BY 数值列 DESC/ASC 返回0行。ORDER BY ID DESC（整数列）正常，但 ORDER BY 值 DESC（含NULL、超大数1.5E10、负数、零值、小数的混合列）返回0行。可能原因是dtype解析时混合类型导致排序失败。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID, 值 FROM Sheet ORDER BY 值 DESC')\nprint(len(r['data']), 'rows')  # 预期20行，实际0行\n\"\n```\n修复后必须跑验证代码，输出20 rows才能标DONE。"
    },
    "REQ-054": {
      "title": "嵌套子查询只返回1行",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "WHERE 值 > (SELECT AVG(值) FROM Sheet WHERE 分类 = 'A') 嵌套子查询只返回1行，预期返回9行。硬编码值 WHERE 值 > 279.95 返回正确10行，说明子查询结果传递有问题。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID, 名称, 值 FROM Sheet WHERE 值 > (SELECT AVG(值) FROM Sheet WHERE 分类 = \\\"A\\\")')\nprint(len(r['data']), 'rows')  # 预期9行\n\"\n```\n修复后必须跑验证代码，输出9 rows才能标DONE。"
    },
    "REQ-055": {
      "title": "支持 EXCEPT / INTERSECT 集合操作",
      "type": "feature",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "EXCEPT 和 INTERSECT 返回0行，未实现。UNION 已支持。底层可用 pandas merge(how='inner'/'outer') 或 set 操作实现。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID FROM Sheet WHERE 分类 = \\\"A\\\" EXCEPT SELECT ID FROM Sheet WHERE 值 > 300')\nprint(len(r['data']), 'rows')  # 预期12行（A类ID减去值>300的ID）\n```\n预期 EXCEPT: {1,2,4,6,8,10,12,16}，预期 INTERSECT: {14,18,20}。"
    },
    "REQ-056": {
      "title": "支持 CTE (WITH AS) 公共表表达式",
      "type": "feature",
      "priority": "P3",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "WITH high_val AS (SELECT ... FROM ...) SELECT ... FROM high_val 返回0行，未实现。需要多步解析：先执行CTE子句存中间结果，再在主查询中引用。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','WITH high_val AS (SELECT ID, 名称, 值 FROM Sheet WHERE 值 > 300) SELECT * FROM high_val WHERE 分类 = \\\"B\\\"')\nprint(len(r['data']), 'rows')  # 预期5行\n```\n"
    },
    "REQ-057": {
      "title": "支持窗口函数 (ROW_NUMBER/RANK/SUM OVER)",
      "type": "feature",
      "priority": "P3",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "ROW_NUMBER() OVER (ORDER BY ...)、RANK() OVER (PARTITION BY ... ORDER BY ...)、SUM(值) OVER (PARTITION BY ...) 均返回0行。底层可用 pandas rolling/groupby.shift/transform 模拟。使用频率低，P3优先级。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID, 分类, 值, RANK() OVER (PARTITION BY 分类 ORDER BY 值 DESC) as rk FROM Sheet')\nprint(len(r['data']), 'rows')  # 预期20行\n```\n"
    },
    "REQ-058": {
      "title": "COALESCE 对 NULL 值不生效",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "COALESCE(值, 0) 对ID=2的空值行返回空字符串''而非0。可能是因为空单元格被读为空字符串而非NULL。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID, COALESCE(值, 0) as safe_val FROM Sheet WHERE ID = 2')\nprint(r['data'][1])  # 预期 [2, '空值测试', 0]，实际 [2, '空值测试', '']\n```\n修复后第二行第三个值必须是0。"
    },
    "REQ-059": {
      "title": "HAVING 不过滤 TOTAL 行",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "HAVING COUNT(*) > 5 应只返回A类（11行），但实际返回了A、B、TOTAL三行。HAVING 应该只过滤分组行，不包含 TOTAL 汇总行。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT 分类, COUNT(*) as cnt FROM Sheet GROUP BY 分类 HAVING COUNT(*) > 5')\nfor row in r['data']: print(row)  # 预期只有 ['A', 11]，不应包含B和TOTAL\n```\n修复后应只有1行数据（不含表头）。"
    },
    "REQ-060": {
      "title": "子查询 IN 返回全量数据",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "WHERE 分类 IN (SELECT DISTINCT 分类 FROM Sheet WHERE 值 > 500) 应只返回B类（值>500的行），但实际返回了全部20行。子查询结果没有正确传递给IN条件。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,'src')\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query('/tmp/extreme-test.xlsx','SELECT ID, 名称 FROM Sheet WHERE 分类 IN (SELECT DISTINCT 分类 FROM Sheet WHERE 值 > 500)')\nprint(len(r['data'])-1, 'rows')  # 预期9行（所有B类），实际返回了20行\n```\n修复后应返回9行B类数据。"
    }
  }
}
