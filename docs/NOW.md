# NOW.md - 第254轮

## 当前状态
- **轮次**: 第254轮
- **时间**: 2026-04-02

## 完成工作
- REQ-036: 边缘案例测试T111-T130（20个API工具测试）
  - 覆盖工具：get_range, check_duplicate_ids, query(WHERE/GROUP BY/ORDER BY), copy_sheet, compare_sheets, compare_files, rename_column, describe_table, search(regex), update_query, assess_data_impact, preview_operation, search_directory, upsert_row(insert/update), get_operation_history, server_stats
  - 20 PASS / 0 INFO / 0 FAIL
  - 发现并修复BUG：check_duplicate_ids传入列名时被错误解释为列字母（column_index_from_string('ID')=238）
  - 修复方案：先在表头行搜索列名，未找到再回退列字母解释
- v1.7.6 已发布 PyPI

## 关键指标
- **版本**: v1.7.6
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: ca9a025 (develop), d7380b0 (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
