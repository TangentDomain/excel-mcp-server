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
    "REQ-066": {
      "title": "数据验证写入失败（67%失败率）",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 2,
      "source": "董事长全面测试报告（107次调用，25.2%错误率）",
      "created": "2026-04-06",
      "description": "数据验证（write_data_validation）写入失败率高达67%。需要排查验证规则的创建逻辑，确保各种验证类型（整数、小数、列表、文本长度等）都能正确写入。",
      "notes": "董事长测试报告发现，P1优先级。需要覆盖所有验证类型进行测试。"
    },
    "REQ-067": {
      "title": "插入列后实际未插入但报成功",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "description": "调用insert_columns插入列后，返回成功但实际列未插入。可能是openpyxl操作后未正确保存，或insert逻辑有误。",
      "notes": "需要验证insert_columns的完整流程：创建→插入→保存→验证。"
    },
    "REQ-068": {
      "title": "公式计算报'不支持的文件格式'",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "description": "apply_formula执行公式计算时报'不支持的文件格式'错误。需要排查公式应用时的文件格式检测逻辑。",
      "notes": "可能与文件扩展名或openpyxl/calamine引擎选择有关。"
    },
    "REQ-069": {
      "title": "写入覆盖功能异常（OperationResult无get属性）",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "description": "写入覆盖操作报OperationResult对象无get属性错误。OperationResult可能是dataclass或Pydantic model，代码中用dict.get()方式访问导致AttributeError。",
      "notes": "检查所有使用OperationResult的地方，统一用属性访问而非dict方法。"
    },
    "REQ-070": {
      "title": "双行表头识别不一致（describe_table vs get_headers）",
      "type": "fix",
      "priority": "P1",
      "status": "OPEN",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "description": "describe_table和get_headers对双行表头的识别结果不一致。需要统一表头解析逻辑，确保同一文件两种API返回一致的表头信息。",
      "notes": "需要定义明确的表头识别规则，特别是合并单元格和双行表头的场景。"
    },
    "REQ-071": {
      "title": "修复Conventional Commits提交格式违规",
      "type": "fix",
      "priority": "P2",
      "status": "OPEN",
      "attempts": 2,
      "source": "FEEDBACK.md #1",
      "created": "2026-04-06",
      "description": "提交 `4d230c9` 违反规范，缺少type前缀，格式为 `[REQ-065] DONE + 新增REQ-066~070` 而非 `[REQ-065] type: DONE + 新增REQ-066~070`。需要修正提交信息，添加正确的type前缀（如feat:或fix:），确保符合CONVENTIONAL_COMMITS.md规范。",
      "notes": "使用 `git commit --amend --no-edit` 修正提交信息，添加正确的type前缀。type必须是feat/fix/refactor/docs/test/chore/perf之一。"
    }
  }
}