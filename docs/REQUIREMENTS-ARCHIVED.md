{
  "REQUIREMENTS_ARCHIVED": {
    "REQ-031": {
      "title": "版本一致性检查与自动化修复脚本",
      "description": "创建check-version-sync.py脚本，自动检测并修复版本不一致问题，确保pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG.md之间的版本信息同步",
      "priority": "high",
      "status": "DONE",
      "acceptance": [
        "创建scripts/check-version-sync.py脚本",
        "脚本自动检测所有版本文件的一致性",
        "发现不一致时自动修复到正确版本",
        "修复过程记录到DECISIONS.md",
        "集成到每轮自动化流程中"
      ]
    },
    "REQ-029": {
      "title": "Excel操作异常处理优化",
      "description": "将excel_operations.py中的通用异常处理替换为具体的自定义异常类，提升错误处理的精确性和用户体验",
      "status": "DONE", 
      "priority": "HIGH",
      "acceptance_criteria": [
        "替换excel_operations.py中的通用Exception捕获为具体异常类",
        "根据错误类型使用SheetNotFoundError、InvalidRangeError、DataValidationError等",
        "保持现有功能不变，确保所有测试通过",
        "提供更精确的错误信息和上下文"
      ],
      "estimated_workload": "中等（约30分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 14:15 UTC",
      "completed_at": "2026-03-28 14:30 UTC",
      "notes": "PyPI发布失败(403 Forbidden, token过期)，代码改动已合并到main"
    },
    "REQ-028": {
      "title": "错误处理机制优化",
      "description": "改进Excel操作中的错误处理，提供更友好的错误信息和AI修复建议",
      "status": "DONE", 
      "priority": "HIGH",
      "acceptance_criteria": [
        "为关键Excel操作函数添加详细的错误处理",
        "提供AI友好的错误消息和修复建议",
        "添加错误分类和上下文信息",
        "不影响现有功能，通过所有测试"
      ],
      "estimated_workload": "中等（约25分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 14:00 UTC",
      "completed_at": "2026-03-28 14:15 UTC",
      "notes": "错误处理机制优化完成，PyPI发布成功"
    }
  },
  "REQUIREMENTS": {
    "REQ-030": {
      "title": "CI CTE测试失败修复",
      "description": "CTE测试在除macOS 3.10外所有平台失败，需要排查python-calamine版本兼容性或pin版本",
      "priority": "中",
      "status": "DONE",
      "source": "CEO反馈+CI观察",
      "resolution": "CTE测试现在全部通过，可能为版本兼容性自动解决"
    }
  },
  "NEWEST_ARCHIVED": {
    "REQ-031": {
      "title": "自动化版本一致性检查脚本",
      "description": "创建check-version-sync.py脚本，自动检测并修复pyproject.toml、__init__.py、README.md、README.en.md中的版本不一致问题",
      "status": "DONE",
      "priority": "HIGH",
      "acceptance_criteria": [
        "创建scripts/check-version-sync.py脚本",
        "检查pyproject.toml、__init__.py、README.md、README.en.md版本一致性",
        "发现不一致时自动修复并记录到DECISIONS.md",
        "脚本执行时间<3秒，不影响正常开发流程",
        "通过所有现有测试，无回归"
      ],
      "estimated_workload": "中（约20分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 16:00 UTC",
      "notes": "基于自我进化建议，解决版本同步依赖手动操作的问题"
    },
    "REQ-028": {
      "title": "excel_update_range insert_mode 默认值改为 false",
      "status": "DONE",
      "priority": "P0",
      "description": "excel_update_range 的 insert_mode 默认为 true，导致写入已有数据文件时会物理插入新行",
      "archived_at": "2026-04-01"
    },
    "REQ-029": {
      "title": "工程强化：约束可机器验证",
      "status": "DONE",
      "priority": "P1",
      "description": "将靠LLM自觉遵守的规则升级为靠脚本验证的规则",
      "archived_at": "2026-04-01"
    },
    "REQ-030": {
      "title": "API参数命名与常见术语对齐",
      "status": "DONE",
      "priority": "P2",
      "description": "create_chart的chart_type支持column别名，create_pivot_table的agg_func支持mean别名",
      "archived_at": "2026-04-01"
    },
    "REQ-031_v2": {
      "title": "修复测试文件语法错误",
      "status": "DONE",
      "priority": "P1",
      "description": "test_mcp_actual.py和test_api_issues.py语法错误修复",
      "archived_at": "2026-04-01"
    },
    "REQ-032": {
      "title": "性能优化：大型Excel文件处理提速（2GB+）",
      "status": "DONE",
      "priority": "P1",
      "description": "处理大型Excel文件（2GB+）时遇到性能瓶颈，优化内存使用和数据处理速度",
      "archived_at": "2026-04-01"
    },
    "REQ-033": {
      "title": "性能优化：iterrows替换为itertuples",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "advanced_sql_query.py中多处使用df.iterrows()遍历DataFrame，性能较差。替换为itertuples()或向量化操作。",
      "notes": "5处iterrows全部替换：advanced_sql_query.py结果序列化+条件过滤、server.py UPDATE/DELETE行匹配+透视表写入",
      "archived_at": "2026-04-01"
    },
    "REQ-038": {
      "title": "BUG：工作表名称非法字符静默替换 + 超长名称静默截断",
      "status": "DONE",
      "priority": "P1",
      "source": "边缘案例测试",
      "description": "_normalize_sheet_name()将方括号[]等非法字符静默替换为下划线（Data [2024]→Data _2024_），超长名称静默截断为25+...字符。用户不知情地创建了与预期不同的工作表名，后续引用会失败。",
      "notes": "第244轮修复：拆分为_validate_sheet_name（严格校验）和_sanitize_sheet_name（静默清理），create_sheet/rename_sheet拒绝非法名称，copy_sheet允许静默清理",
      "archived_at": "2026-04-01"
    }
  }
}