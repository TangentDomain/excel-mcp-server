# 第120轮 - REQ-030 修复完成 (v1.6.8) ✅

## 状态
版本：v1.6.8 | 工具：44 | 测试：1107

## 本轮完成
- **Bug 1修复**：`MAX(a+b)`等聚合函数内多列表达式计算 — 新增`_is_expression`和`_evaluate_expression`方法，支持四则运算和字面量嵌套
- **Bug 2修复**：SELECT子句中的标量子查询 — `_apply_select_expressions`和`_apply_group_by_aggregation`均新增Subquery处理
- **Bug 3验证**：LEFT JOIN + IS NULL经验证已正常工作，无需修复
- **全量测试**：1107 passed
- **PyPI发布**：v1.6.8 已发布

## 修复摘要
| Bug | 问题 | 修复方案 |
|-----|------|----------|
| 1 | `MAX(攻击力+防御力)`失败 | 表达式求值递归处理Add/Sub/Mul/Div/Literal |
| 2 | SELECT中的标量子查询不支持 | 新增Subquery分支，支持SELECT/WHERE/HAVING |
| 3 | LEFT JOIN IS NULL | 已验证正常，无需修复 |

## 下轮待办
- [ ] MCP真实验证（确认修复）
- [ ] README中英文同步检查
