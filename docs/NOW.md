# 第138轮 - REQ-029 JOIN表别名 + describe_table崩溃修复 ✅

---

## 状态
版本：v1.6.20 | 工具：44 | 测试：1168

## 本轮完成
- **REQ-029 P0 阻断性Bug修复**（2个bug）
  - **Bug 1 - JOIN表别名映射**：修复了JOIN查询中使用表限定符(`r.名称`/`s.名称`)引用列时，`_apply_select_expressions`无法正确映射pandas merge后的`_x`/`_y`后缀问题
    - 在`_expression_to_column_reference`中增加多格式回退查找逻辑
  - **Bug 2 - describe_table崩溃**：修复了streaming写入后openpyxl read_only模式`max_row=None`导致的崩溃
    - 优化`excel_describe_table`中行数统计逻辑，增加iter_rows回退机制
- **PyPI发布**：v1.6.20 → https://pypi.org/project/excel-mcp-server-fastmcp/1.6.20/

## 下轮待办
- [ ] REQ-010 文档与门面优化
- [ ] REQ-006 工程治理（持续迭代）
