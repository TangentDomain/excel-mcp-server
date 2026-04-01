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
      "notes": "来源：监工第3轮报告。'column' 和 'mean' 是 pandas/openpyxl 用户的常用术语，不支持会导致用户困惑。\n\n完成情况（2026-04-01）：\n✅ create_chart 的 chart_type 已支持 'column' 作为 'bar' 的别名（chart_map 中两者都映射到 BarChart）\n✅ docstring 已说明 'column' 作为 'bar' 的别名\n✅ 测试 test_create_column_chart 已存在并通过\n❌ create_pivot_table 函数不存在，无法实现 mean/average 别名功能\n\n注：pivot_table 功能未被实现，此部分验收标准不适用。如需此功能，应创建新需求先实现基础功能。"
    }
  }
}
