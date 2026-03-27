# 第114轮 - REQ-029 JOIN别名 + describe_table崩溃修复 ✅ 完成

## 状态
版本：v1.6.3 | 工具：44 | 测试：1107

## 本轮完成
- REQ-029 Bug 1修复：JOIN表别名解析增强
  - FROM表别名优先使用Table.alias属性，备用TableAlias/Alias节点遍历
  - SELECT重复别名自动添加表前缀避免覆盖
  - JOIN右表别名解析同样优先Table.alias
  - CROSS JOIN列冲突用临时重命名保护
- REQ-029 Bug 2修复：describe_table max_row=None健壮处理
  - read_only模式下max_row=None时逐行扫描+连续空行检测停止
  - 修复原有死代码bug（continue后不可达的break）
- 全量测试1107 passed
- 合并develop→main，发布PyPI v1.6.3，tag+push完成

## 待办
- [ ] MCP验证（至少8项游戏场景）
- [ ] 更新DECISIONS.md记录决策
