# 第135轮 - README同步完成 ✅

---

## 状态
版本：v1.6.18 | 工具：44 | 测试：1168

## 本轮完成
- **REQ-029 BUG FIX**：修复JOIN表别名映射bug
  - Bug 1: SELECT中使用表限定符(r.名称)时无法正确解析左表列引用 → 已修复
  - Bug 2: describe_table流式写入后max_row=None崩溃 → 之前已修复，验证确认
- MCP真实验证完成：12项核心功能测试通过
- 全量测试1156 passed
- PyPI v1.6.18 已发布

### 关键改动
- `_apply_select_expressions`: qualified列名查找失败时回退到`_expression_to_column_reference`映射
- `_expression_to_column_reference`: 增强_x/_y后缀处理，新增table_part_x/y直接匹配
- MCP真实验证脚本创建，覆盖JOIN别名、describe_table稳定性等核心功能

## 下轮待办
- [ ] REQ-006 工程治理（持续迭代）
- [ ] REQ-010 文档与门面优化