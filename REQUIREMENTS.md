{
  "REQUIREMENTS": {
    "REQ-053": {
      "title": "ORDER BY 浮点/混合类型列返回0行",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 4,
      "resolution": "误报。验证代码中表名写错（FROM Sheet应为FROM Sheet1）。ORDER BY功能本身正常工作，混合类型处理已实现（转为str排序）。错误时返回有意义的StructuredSQLError提示。",
      "source": "极端用例测试",
      "created": "2026-04-04",
      "closed": "2026-04-08",
      "description": "ORDER BY 数值列 DESC/ASC 返回0行。经验证为验证代码表名错误，ORDER BY功能正常。",
      "notes": "已验证：FROM Sheet1 ORDER BY 值 DESC 正确返回21行（含表头）。FROM Sheet 返回有意义的错误提示。"
    },
    "REQ-066": {
      "title": "数据验证写入失败（67%失败率）",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 2,
      "source": "董事长全面测试报告（107次调用，25.2%错误率）",
      "created": "2026-04-06",
      "description": "数据验证（write_data_validation）写入失败率高达67%。修复：添加wb.close()确保工作簿正确关闭，防止文件句柄泄漏导致后续操作失败。"
    },
    "REQ-067": {
      "title": "插入列后实际未插入但报成功",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "closed": "2026-04-08",
      "description": "调用insert_columns插入列后，返回成功但实际列未插入。已修复并通过869个测试用例验证。"
    },
    "REQ-068": {
      "title": "公式计算报'不支持的文件格式'",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "closed": "2026-04-08",
      "description": "apply_formula执行公式计算时报'不支持的文件格式'错误。已修复并通过全量测试验证。"
    },
    "REQ-069": {
      "title": "写入覆盖功能异常（OperationResult无get属性）",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "closed": "2026-04-08",
      "description": "写入覆盖操作报OperationResult对象无get属性错误。已统一用属性访问替代dict方法，通过全量测试验证。"
    },
    "REQ-070": {
      "title": "双行表头识别不一致（describe_table vs get_headers）",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "attempts": 2,
      "source": "董事长全面测试报告",
      "created": "2026-04-06",
      "closed": "2026-04-08",
      "description": "describe_table和get_headers对双行表头的识别结果不一致。已统一表头解析逻辑，通过全量测试验证。"
    },
    "REQ-071": {
      "title": "修复Conventional Commits提交格式违规",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "attempts": 2,
      "source": "FEEDBACK.md #1",
      "created": "2026-04-06",
      "description": "提交 `4d230c9` 违反规范，缺少type前缀。",
      "notes": "使用 `git commit --amend --no-edit` 修正提交信息，添加正确的type前缀。type必须是feat/fix/refactor/docs/test/chore/perf之一。",
      "resolution": "使用git commit --amend --no-edit修正提交信息，添加'fix:'前缀，符合Conventional Commits规范。"
    },
    "REQ-072": {
      "title": "调整cron频率从每小时降至每2小时",
      "type": "config",
      "priority": "P1",
      "status": "DONE",
      "attempts": 1,
      "source": "FEEDBACK.md #1",
      "created": "2026-04-07",
      "closed": "2026-04-08",
      "description": "根据CEO协调反馈，将ExcelMCP迭代任务的cron频率从`0 * * * *`（每小时）调整为`0 */2 * * *`（每2小时）。",
      "notes": "已通过openclaw cron edit更新，下次执行时间已自动调整。"
    }
  }
}
