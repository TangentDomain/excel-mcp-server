{
  "REQUIREMENTS": {
    "REQ-035": {
      "title": "配置化：硬编码常量提取为配置项",
      "type": "refactor",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "attempts": 0,
      "last_failure": "",
      "description": "多个硬编码值应提取为可配置常量：max_files=100, query_cache_ttl=300, target_mb=512.0, MAX_RESULT_ROWS=500等。",
      "notes": "分布在server.py和advanced_sql_query.py中"
    },
    "REQ-044": {
      "title": "find_last_row列名查找与check_duplicate_ids一致化",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 1,
      "last_failure": "",
      "description": "find_last_row中使用column_index_from_string直接解释列参数，与check_duplicate_ids修复后的行为不一致。当用户传入列名(如'ID')时会被错误解释为列字母。应抽取公共列解析方法，统一先查表头再回退列字母的逻辑。",
      "notes": "第257轮修复：先查表头匹配列名，找不到再回退列字母解释"
    },
    "REQ-045": {
      "title": "batch_insert_rows insert_position模块导入错误",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "边缘案例测试",
      "attempts": 1,
      "last_failure": "",
      "description": "batch_insert_rows指定insert_position时报错：No module named 'excel_mcp_server_fastmcp.api.excel...'。模块路径可能有误。",
      "notes": "第257轮修复：ExcelWriter导入路径从api.excel_writer改为core.excel_writer"
    },
    "REQ-046": {
      "title": "delete_rows condition数值类型比较问题",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 1,
      "last_failure": "",
      "description": "delete_rows使用condition参数（如'Score < 60'）时，对数值列返回0行删除。疑似条件解析将数值列作为字符串处理，导致比较失败。",
      "notes": "第257轮修复：df.query前用pd.to_numeric(errors='ignore')转换数值列"
    },
    "REQ-036": {
      "title": "边缘案例自动化测试：每轮自动搜索并验证奇怪场景",
      "type": "feature",
      "priority": "P1",
      "status": "OPEN",
      "source": "CEO",
      "attempts": 1,
      "last_failure": "",
      "description": "每轮执行时，自动搜索一些稀奇古怪的Excel使用场景（如超长公式、特殊字符sheet名、合并单元格+筛选、数据透视表嵌套、条件格式+VBA、超大文件性能等），用uvx安装的MCP工具实际调用测试，记录是否崩溃/返回错误/正常处理。",
      "acceptance_criteria": [
        "每轮至少测试1个新边缘案例",
        "测试结果记录到docs/EDGE-CASE-TESTS.md",
        "格式：日期、案例描述、操作步骤、预期结果、实际结果、是否通过",
        "崩溃或错误自动创建REQ",
        "优先从Stack Overflow/GitHub Issues搜索真实用户遇到的奇怪问题"
      ],
      "notes": "第243轮测试10个案例6通过4失败(REQ-038/039/040)；第245轮测试10个案例9通过1失败(REQ-041)；第247轮测试5个案例全通过(REQ-042修复)；第248轮测试16个案例15通过1信息；第249轮测试16个案例13通过3信息(含server.py修复)；第250轮测试15个案例12通过2信息1失败(||拼接不支持)；第252轮测试33个案例33全通过(核心API稳定性验证)；第253轮测试25个案例11通过11信息3失败(streaming写入不可见)；第254轮测试20个案例20全通过(含check_duplicate_ids列名查找bug修复+发布v1.7.6)；第255轮测试30个案例25通过3信息2失败(T132/T133 ROUND/ABS不支持+T141嵌套聚合计算列丢失)；第256轮测试20个案例19通过1失败(T168 evaluate_formula独立数学表达式不支持)；第258轮测试30个案例(T181-T210)30全通过(SQL GROUP BY/HAVING/LIKE/空表查询/跨sheet公式/100行批量插入/特殊字符sheet名/错误场景容错)"
    }
  }
}
