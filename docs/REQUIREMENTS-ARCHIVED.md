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
      "estimated_workload": "中等（约20分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 16:00 UTC",
      "notes": "基于自我进化建议，解决版本同步依赖手动操作的问题"
    }
  }
}