# DECISIONS.md - 决策记录

## D014: REQ-029 JOIN表别名映射修复 (2026-03-27, R134)
**需求**: REQ-029 BUG FIX
**问题**: JOIN查询中使用表限定符(r.名称)引用左表列时，_apply_select_expressions无法解析
**根因**: qualified列名查找失败后直接回退到无限定符列名，没有尝试JOIN映射
**决策**: 在qualified查找失败时，先调用_expression_to_column_reference进行完整映射
**方案**:
1. _apply_select_expressions增加映射回退逻辑
2. _expression_to_column_reference增强_x/_y后缀处理
3. 创建MCP真实验证脚本，验证12项核心功能
**验证**: 1156 passed, JOIN别名映射测试通过, MCP真实验证完成, PyPI v1.6.18发布