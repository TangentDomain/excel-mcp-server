# 第142轮 - REQ-029 JOIN _x/_y后缀bug修复

---

## 状态
版本：v1.6.23 | 工具：44 | 测试：1159

## 本轮完成
- **REQ-029 Bug 1修复**：JOIN ON不同列名时pandas merge产生_x/_y后缀
  - 根因：elif条件`left_on_col == right_on_col`限制过严，遗漏左ON列在右表存在的场景
  - 修复：移除该限制，确保左ON列在右表存在时始终重命名为`alias.col`
  - 3个回归测试：无后缀/别名列可引用/WHERE正常
- **REQ-029 Bug 2确认**：describe_table崩溃已在v1.6.20修复（D018），无需再改
- **PyPI**：v1.6.23已发布

## 下轮待办
- [ ] 每5轮MCP真实验证（下次第145轮）
- [ ] REQ-025 docstring继续优化剩余函数
- [ ] REQ-010 文档与门面优化
