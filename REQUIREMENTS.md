# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求

### REQ-037 [P1] 仓库文件治理

**状态**: OPEN
**优先级**: P1

**问题**:
仓库根目录和scripts/、docs/积累了大量元迭代内部文件、临时产物、备份脚本，用户看到的仓库杂乱，影响专业形象。

**治理原则**:
- **用户看到的根目录**只有：README.md、README.en.md、LICENSE、CHANGELOG.md、CONTRIBUTING.md、pyproject.toml、.github/、src/、tests/、examples/
- **元迭代内部文件**（.cron-prompt.md、.cron-focus.md、docs/NOW.md、docs/DECISIONS.md、docs/ROADMAP.md、docs/RULES.md等）全部.gitignore排除，不进入用户仓库
- **临时产物**（*.json报告、__pycache__、.backup文件）加入.gitignore
- **scripts/** 只保留对用户有用的1-2个（check-version-sync.py、health-monitor.py），其余删除或.gitignore
- **docs/** 只保留对外文档（README-*.md、testing-guidelines.md），内部文件.gitignore

**具体清理清单**:

根目录删除：
- AGENTS.md、CLAUDE.md（OpenClaw/Claude配置，不属于用户仓库）
- DEPLOYMENT.md、deploy.bat、Makefile（MCP项目不需要）
- CONTRIBUTING.zh-CN.md（有英文版就够了）
- 项目说明.md（跟README重复）
- REQUIREMENTS-ARCHIVED.md、ARCHIVED.md（移到docs/或删除）

根目录.gitignore：
- .cron-prompt.md、.cron-focus.md
- docs-sync-report.json、health-monitor-result.json、star-stats.json、req035_done.txt

scripts/清理：
- 删除所有*-old.py、*-old-*.py备份文件
- 删除docs-sync-check.py、README-monitoring.md、run-monitor.bat、monitor-config.json（元迭代内部）
- 删除test_check_version_sync.py、test_compatibility.py、validate_mcp_tools.py、mcp_test.py、quick-mcp-verify.py（验证脚本，应该在tests/）
- 删除__pycache__目录
- 保留：check-version-sync.py、health-monitor.py、github-stats.py、star-thanks.py

docs/清理：
- 删除DECISIONS-ARCHIVED-*-R123.md、*-TOP10.md、*-temp.md（临时归档）
- 删除NOW.md.backup、SELF-EVALUATION-198.md
- 删除FEISHU_SUMMARY.md（内部推送记录）
- 删除readme-redesign.md（设计稿，用完）
- 删除游戏开发Excel配置表比较指南.md（跟README-gaming重复）
- .gitignore排除：NOW.md、DECISIONS.md、ROADMAP.md、RULES.md、REQUIREMENTS.md、ARCHIVED.md、DECISIONS-ARCHIVED.md、FEISHU_SUMMARY.md

**注意**:
- 清理前先commit当前状态，避免丢失
- 删除.gitignore排除的文件用 `git rm --cached`
- 清理后验证CI通过
- 子代理后续轮次生成的内部文件必须被.gitignore覆盖

**验收标准**:
- 用户clone仓库后，根目录整洁，只有必要的项目文件
- scripts/ ≤ 5个文件
- docs/ ≤ 8个文件（对外文档）
- CI通过

暂无活跃需求

## 已完成需求

参见 [ARCHIVED.md](ARCHIVED.md) 获取已完成的需求记录。