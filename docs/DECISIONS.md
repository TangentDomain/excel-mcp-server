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

## D015: README同步与版本更新 (2026-03-27, R135)
**需求**: README检查（中英文同步）
**问题**: 中文和英文README存在不一致，版本信息过时，测试数量不准确
**根因**: 持续迭代过程中未及时同步文档，版本管理分散
**决策**: 统一更新两个README文件，同步版本号和测试数量
**方案**:
1. 更新测试覆盖数量：1099 → 1168个测试函数
2. 同步版本号：pyproject.toml和__init__.py更新到1.6.18
3. 确保中英文README完全同步
**验证**: 版本一致性检查通过，test count验证完成