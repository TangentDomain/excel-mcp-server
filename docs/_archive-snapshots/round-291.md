{
  "REQUIREMENTS": {
    "REQ-053": {
      "title": "ORDER BY 浮点/混合类型列返回0行",
      "type": "fix",
      "priority": "P1",
      "status": "PAUSED",
      "attempts": 4,
      "last_failure": "连续4次attempt未完成，ORDER BY混合类型排序涉及pandas底层限制，需更深层方案",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "ORDER BY 数值列 DESC/ASC 返回0行。ORDER BY ID DESC（整数列）正常，但 ORDER BY 值 DESC（含NULL、超大数1.5E10、负数、零值、小数的混合列）返回0行。可能原因是dtype解析时混合类型导致排序失败。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT ID, 值 FROM Sheet ORDER BY 值 DESC\")\nprint(len(r[\"data\"]), \"rows\")  # 预期20行，实际0行\n\"\n```\n修复后必须跑验证代码，输出20 rows才能标DONE。"
    },



  }
}
    "REQ-061": {
      "title": "GROUP BY 聚合逻辑 bug修复",
      "type": "fix",
      "priority": "P0",
      "status": "OPEN",
      "attempts": 0,
      "source": "FEEDBACK.md",
      "created": "2026-04-05",
      "description": "GROUP BY 聚合逻辑存在bug，需要修复_apply_group_by_aggregation方法中的聚合逻辑错误。CEO明确要求优先修复，不能绕过。",
      "notes": "验证规则：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/test.xlsx\", \"SELECT ID, value FROM data GROUP BY ID\")\nprint(len(r[\"data\"]), \"rows\")  # 修复后输出预期结果\n\"\n```\n修复后必须跑验证代码，输出正确结果才能标DONE。"
    },
    "REQ-062": {
      "title": "DOCSTRING 文档完整性修复",
      "type": "feature",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 0,
      "source": "FEEDBACK.md",
      "created": "2026-04-05",
      "description": "批量修复所有公共函数的docstring，确保包含Args/Parameters和Returns段，建立自动化docstring检查机制，设定90%以上合规率目标。当前506个函数中发现541个文档问题，合规率仅-6.9%。",
      "notes": "目标：所有公共函数都应有完整的Args/Parameters和Returns文档段，文档覆盖率达到90%以上"
    }
