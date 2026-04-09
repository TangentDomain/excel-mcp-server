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
    "REQ-054": {
      "title": "嵌套子查询只返回1行",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "WHERE 值 > (SELECT AVG(值) FROM Sheet WHERE 分类 = \"A\") 嵌套子查询只返回1行，预期返回9行。硬编码值 WHERE 值 > 279.95 返回正确10行，说明子查询结果传递有问题。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT ID, 名称, 值 FROM Sheet WHERE 值 > (SELECT AVG(值) FROM Sheet WHERE 分类 = \\\"A\\\")\")\nprint(len(r[\"data\"]), \"rows\")  # 预期9行\n\"\n```\n修复后必须跑验证代码，输出9 rows才能标DONE。"
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
      "description": "EXCEPT 和 INTERSECT 返回0行，未实现。UNION 已支持。底层可用 pandas merge(how=\"inner\"/\"outer\") 或 set 操作实现。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT ID FROM Sheet WHERE 分类 = \\\"A\\\" EXCEPT SELECT ID FROM Sheet WHERE 值 > 300\")\nprint(len(r[\"data\"]), \"rows\")  # 预期12行（A类ID减去值>300的ID）\n```\n预期 EXCEPT: {1,2,4,6,8,10,12,16}，预期 INTERSECT: {14,18,20}。"
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
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"WITH high_val AS (SELECT ID, 名称, 值 FROM Sheet WHERE 值 > 300) SELECT * FROM high_val WHERE 分类 = \\\"B\\\"\")\nprint(len(r[\"data\"]), \"rows\")  # 预期5行\n```\n"
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
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT ID, 分类, 值, RANK() OVER (PARTITION BY 分类 ORDER BY 值 DESC) as rk FROM Sheet\")\nprint(len(r[\"data\"]), \"rows\")  # 预期20行\n```\n"
    },
      "status": "OPEN",
      "attempts": 1,
      "last_failure": "",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "description": "HAVING COUNT(*) > 5 应只返回A类（11行），但实际返回了A、B、TOTAL三行。HAVING 应该只过滤分组行，不包含 TOTAL 汇总行。",
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT 分类, COUNT(*) as cnt FROM Sheet GROUP BY 分类 HAVING COUNT(*) > 5\")\nfor row in r[\"data\"]: print(row)  # 预期只有 [\"A\", 11]，不应包含B和TOTAL\n```\n修复后应只有1行数据（不含表头）。"
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
      "notes": "验证代码：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/extreme-test.xlsx\",\"SELECT ID, 名称 FROM Sheet WHERE 分类 IN (SELECT DISTINCT 分类 FROM Sheet WHERE 值 > 500)\")\nprint(len(r[\"data\"])-1, \"rows\")  # 预期9行（所有B类），实际返回了20行\n```\n修复后应返回9行B类数据。"
    "REQ-061": {
      "title": "GROUP BY 聚合逻辑 bug",
      "type": "fix",
      "priority": "P0",
      "status": "IN-PROGRESS",
      "attempts": 3,
      "last_failure": "第286轮超时未完成",
      "source": "FEEDBACK.md OPEN-#1",
      "created": "2026-04-04",
      "description": "GROUP BY 聚合逻辑导致TOTAL行数据错误，影响所有聚合查询。必须在_apply_group_by_aggregation方法中修复聚合计算逻辑。",
      "notes": "验证规则：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/test.xlsx\",\"SELECT 分类, SUM(值) as total_val FROM Sheet GROUP BY 分类 ORDER BY 分类\")\nfor row in r[\"data\"]: print(row)  # 第二行应该是 [\"TOTAL\", \"实际聚合值\"]\n\"\n```\n修复后TOTAL行的聚合值必须正确。"
>>>>>>> develop
    }
  }
}
