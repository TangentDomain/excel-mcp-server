{
  "REQUIREMENTS": {
    "REQ-031": {
      "title": "自动化版本一致性检查脚本",
      "description": "创建check-version-sync.py脚本，自动检测并修复pyproject.toml、__init__.py、README.md、README.en.md中的版本不一致问题",
      "status": "OPEN",
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