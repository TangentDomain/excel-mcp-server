{
  "REQUIREMENTS": {
    "REQ-034": {
      "title": "代码质量：路径验证逻辑抽取为装饰器",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "description": "server.py中_validate_path检查模式重复出现10+次，抽取为装饰器减少重复代码。",
      "notes": "涉及行768、800、836、862、1019、1055、1170、1486、1532、1578"
    },
    "REQ-035": {
      "title": "配置化：硬编码常量提取为配置项",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "description": "多个硬编码值应提取为可配置常量：max_files=100, query_cache_ttl=300, target_mb=512.0, MAX_RESULT_ROWS=500等。",
      "notes": "分布在server.py和advanced_sql_query.py中"
    },
    "REQ-036": {
      "title": "边缘案例自动化测试：每轮自动搜索并验证奇怪场景",
      "priority": "P1",
      "source": "CEO",
      "description": "每轮执行时，自动搜索一些稀奇古怪的Excel使用场景（如超长公式、特殊字符sheet名、合并单元格+筛选、数据透视表嵌套、条件格式+VBA、超大文件性能等），用uvx安装的MCP工具实际调用测试，记录是否崩溃/返回错误/正常处理。",
      "acceptance_criteria": [
        "每轮至少测试1个新边缘案例",
        "测试结果记录到docs/EDGE-CASE-TESTS.md",
        "格式：日期、案例描述、操作步骤、预期结果、实际结果、是否通过",
        "崩溃或错误自动创建REQ",
        "优先从Stack Overflow/GitHub Issues搜索真实用户遇到的奇怪问题"
      ],
      "notes": "第243轮测试10个案例，6通过4失败，发现3个BUG+1个功能缺失，已创建REQ-038/039/040"
    },
    "REQ-037": {
      "title": "线程安全：formula_cache并发访问保护",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "description": "formula_cache.py中的缓存操作在并发MCP调用场景下缺乏线程安全保护，可能导致缓存数据竞争。",
      "notes": "涉及utils/formula_cache.py，需要添加threading.Lock保护缓存读写"
    },
    "REQ-038": {
      "title": "BUG：工作表名称非法字符静默替换 + 超长名称静默截断",
      "priority": "P1",
      "status": "OPEN",
      "source": "边缘案例测试",
      "description": "_normalize_sheet_name()将方括号[]等非法字符静默替换为下划线（Data [2024]→Data _2024_），超长名称静默截断为25+...字符。用户不知情地创建了与预期不同的工作表名，后续引用会失败。",
      "notes": "core/excel_manager.py:215-224，应拒绝非法名称并返回明确错误信息，或至少警告名称已修改"
    },
    "REQ-039": {
      "title": "功能缺失：list_sheets不区分隐藏工作表",
      "priority": "P2",
      "status": "OPEN",
      "source": "边缘案例测试",
      "description": "excel_list_sheets将visible/hidden/veryHidden工作表一视同仁列出，用户无法区分。应增加sheet_state字段标识可见性。",
      "notes": "涉及server.py的excel_list_sheets工具"
    },
    "REQ-040": {
      "title": "信息不准确：稀疏工作表file_info维度被格式化膨胀",
      "priority": "P2",
      "status": "OPEN",
      "source": "边缘案例测试",
      "description": "当工作表在远端单元格（如Z100）仅有格式化而无数据时，excel_get_file_info返回total_rows=100、total_cols=26，与实际数据范围不符。",
      "notes": "应区分data_range和formatted_range，或标注实际数据维度"
    }
  }
}
