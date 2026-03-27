# 第117轮 - REQ-029 JOIN别名崩溃修复 + v1.6.6发布 ✅

## 状态
版本：v1.6.6 | 工具：44 | 测试：1107

## 本轮完成
- **REQ-029 JOIN表别名映射 + describe_table崩溃修复**：
  - Bug1修复：JOIN后SQL表别名`r.名称`不生效问题，增强_expression_to_column_reference的5层回退映射机制
  - Bug2修复：streaming写入后openpyxl read_only模式max_row=None导致的describe_table崩溃
  - 修复JOIN列映射：新增_join_column_mapping记录JOIN列映射，支持pandas后缀格式`_x`/`_y`
  - describe_table异常处理：添加try/except处理max_row=None情况，优先使用max_row，失败时用iter_rows统计
  - MCP验证：JOIN别名映射正确，streaming写入后describe_table正常统计行数
  - 全量测试通过：1107个测试全部通过
- **v1.6.6发布**：PyPI + GitHub推送完成
- **清理**：worktree已清理，测试文件已清理

## 待办
- [ ] REQ-029 继续其他工具返回值统一
- [ ] MCP真实验证（下一轮需做）

## 决策
- **决策**：JOIN表别名映射采用5层回退机制，确保别名正确解析
- **原因**：pandas JOIN后会添加_x/_y后缀，用户期望使用原始别名r.名称
- **方案**：_expression_to_column_reference新增5层映射逻辑，直接匹配→pandas后缀→JOIN映射→表别名映射→原始列名
- **决策**：describe_table采用max_row优先、iter_rows备用的统计策略
- **原因**：streaming写入后openpyxl read_only模式的max_row可能为None
- **方案**：优先使用max_row统计，访问失败时用iter_rows遍历统计，确保结果准确