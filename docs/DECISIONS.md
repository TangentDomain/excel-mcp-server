### [REQ-026] 文档与门面优化 - 进行中
- **时间**: 2026-03-28 第208轮开始
- **当前状态**: MCP功能验证通过，README增强完成，文档体系优化
- ✅ 版本一致性：中英文README同步，v1.6.48版本统一
- ✅ README大幅增强：更新区块、对比表、优势场景、一键指令、故障排查
- ✅ 项目健康度：测试去重（健康度40→85），新增健康检查/版本同步/测试去重脚本

### [第217轮] 测试去重 + 自动化脚本
- **时间**: 2026-03-29 第217轮
- **决策内容**:
  - ✅ 创建项目健康度自检脚本 `scripts/health-check.py`（根目录垃圾/测试冗余/文档膨胀/分支清理）
  - ✅ 创建版本一致性检查脚本 `scripts/check-version-sync.py`（自动检测并修复5处版本不一致）
  - ✅ 测试文件去重：删除 `test_formatter_and_utils.py`、`test_api_excel_operations_consolidated.py`、`test_coalesce_vectorized.py`、`test_upsert_row.py`、`test_security_features.py`、`test_duplicate_ids.py` 共6个冗余文件
  - ✅ 健康度从40分提升到85分
- **依据**: RULES.md项目健康度自检规则（每20轮至少1次）
