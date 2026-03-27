# 第115轮 - REQ-029 增强修复 + v1.6.4 发布 ✅ 完成

## 状态
版本：v1.6.4 | 工具：44 | 测试：1107

## 本轮完成
- REQ-029 增强修复（基于第114轮基础）：
  - Bug 1 增强：`_expression_to_column_reference` 5层回退别名映射
    - 直接别名格式 `r.名称` → 列中查找
    - pandas后缀格式 `r_名称` → 列中查找
    - `_join_column_mapping` 映射 → JOIN时记录的重命名关系
    - 表别名解析 `_table_aliases` → 原始表名+列名
    - 原始列名兜底
  - Bug 2 增强：`describe_table` max_row=None 异常处理
    - try/except包裹max_row访问，失败时回退到iter_rows统计
- 全量测试1107 passed（0 failed）
- 合并feature/req-029→develop→main
- 发布PyPI v1.6.4，tag+push完成
- REQ-029 标记为 ✅ 完成

## 待办
- [ ] MCP验证（至少8项游戏场景）
- [ ] 更新DECISIONS.md记录决策
