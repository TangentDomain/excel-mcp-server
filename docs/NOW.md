## 第201轮 - REQ-035 CTE测试修复（已完成）
- **REQ-035 [P0] CI CTE测试全平台失败**: 修复sqlglot 27.x CTE arg key兼容性
- **根因**: sqlglot 27.29.0用`'with'`存储CTE WITH子句，代码只检查`'with_'`导致CTE被跳过
- **修复**: 自动检测arg key（`'with'`或`'with_'`），兼容所有sqlglot版本
- **测试改进**: CTE测试assertion加入error message，CI失败时可看到实际错误
- **发布**: v1.6.39（PyPI已发布验证通过）
- **验证**: sqlglot 27.29.0 + python-calamine 0.3.0 下3个CTE测试全部通过

## 轮次指标
- 轮次：第201轮
- 发布：v1.6.39
- 测试：1160 passed, 0 failed
- CI修复：3个CTE测试（11个job × 3 = 33个测试用例）
- 改动：advanced_sql_query.py（3处CTE arg key）+ test_sql_enhanced.py（3处assertion）
