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
      "notes": "第243轮测试10个案例6通过4失败(REQ-038/039/040)；第245轮测试10个案例9通过1失败(REQ-041)；第247轮测试5个案例全通过(REQ-042修复)；第248轮测试16个案例15通过1信息；第249轮测试16个案例13通过3信息(含server.py修复)；第250轮测试15个案例12通过2信息1失败(||拼接不支持)"
    },
    "REQ-037": {
      "title": "线程安全：formula_cache并发访问保护",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "description": "formula_cache.py中的缓存操作在并发MCP调用场景下缺乏线程安全保护，可能导致缓存数据竞争。",
      "notes": "涉及utils/formula_cache.py，需要添加threading.Lock保护缓存读写"
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
