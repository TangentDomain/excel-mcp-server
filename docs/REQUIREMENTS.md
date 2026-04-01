{
  "REQUIREMENTS": {
    "REQ-033": {
      "title": "性能优化：iterrows替换为itertuples",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "description": "advanced_sql_query.py中多处使用df.iterrows()遍历DataFrame，性能较差。替换为itertuples()或向量化操作。",
      "notes": "5处iterrows全部替换：advanced_sql_query.py结果序列化+条件过滤、server.py UPDATE/DELETE行匹配+透视表写入"
    },
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
      "notes": "用uvx excel-mcp-server-fastmcp启动服务器进行测试，测试完要kill进程"
    },
    "REQ-037": {
      "title": "线程安全：formula_cache并发访问保护",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "description": "formula_cache.py中的缓存操作在并发MCP调用场景下缺乏线程安全保护，可能导致缓存数据竞争。",
      "notes": "涉及utils/formula_cache.py，需要添加threading.Lock保护缓存读写"
    }
  }
}
