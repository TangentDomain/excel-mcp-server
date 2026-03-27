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

## D016: REQ-006 核心工具描述优化 (2026-03-27, R136)
**需求**: REQ-006 工程治理 - 工具描述持续优化
**问题**: 部分工具描述冗长但实用性不足，AI难以快速提取关键信息
**根因**: 描述偏向"功能罗列"，缺少"使用建议"和"工具间配合指南"
**决策**: 优化核心工具描述，以"AI可用性"为导向重构描述结构
**方案**:
1. excel_describe_table: 去掉冗余场景，增加"使用建议"和"配合使用"
2. excel_upsert_row: 精简参数说明，突出"关键参数"
3. excel_batch_insert_rows: 增加实用技巧，明确新增vs更新工具选择
4. excel_get_headers: 添加"实用技巧"section，增强AI处理能力
**验证**: 工具导入测试通过，PyPI v1.6.19发布