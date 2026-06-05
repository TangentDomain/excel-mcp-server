{
  "REQUIREMENTS": {
    "REQ-065": {
      "title": "修复主会话sed替换引入的函数签名回归",
      "type": "fix",
      "priority": "P0",
      "status": "IN-PROGRESS",
      "attempts": 1,
      "last_failure": null,
      "source": "FEEDBACK-013",
      "created": "2026-04-06",
      "description": "主会话用 sed -i 批量修复12处 `df -> pd.DataFrame:` 语法错误时，引入了函数签名回归：1) _apply_order_by(self, parsed_sql, df) 参数名被错误替换；2) _apply_join_clause(self, joins, left_df) 同理；3) _evaluate_case_expression 中 row 变量未定义。需要逐个检查被sed替换过的函数签名，对照git diff恢复正确签名。",
      "notes": "测试用例：\n- SELECT name FROM employees ORDER BY salary DESC\n- SELECT name, CASE WHEN salary > 30000 THEN '高' ELSE '低' END FROM employees\n- SELECT e.name FROM employees e JOIN orders o ON e.name = o.customer\n修复后必须跑测试用例验证。"
    },
    "REQ-064": {
      "title": "修复execute_sql_query方法签名和参数丢失",
      "type": "fix",
      "priority": "P0",
      "status": "DONE",
      "attempts": 1,
      "last_failure": null,
      "source": "FEEDBACK-012",
      "created": "2026-04-06",
      "completed": "2026-04-06",
      "description": "REQ-062 docstring修复时引入的回归：execute_sql_query方法签名被错误简化，丢失了sheet_name/limit/include_headers/output_format等参数，且方法体中使用了未定义的sheet_name变量，导致所有SQL查询功能不可用。需要恢复到REQ-062之前的正常版本。",
      "notes": "验证代码：\n\
修复后必须跑验证代码，输出FIXED才能标DONE。验证通过：✅",
      "commit": "ac03e2c"
    },
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
    }
  }
}
