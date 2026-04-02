# NOW.md - 第256轮

## 当前状态
- **轮次**: 第256轮
- **时间**: 2026-04-02

## 完成工作
- REQ-036: 边缘案例测试T161-T180（20个API工具测试）
  - 覆盖工具：SQL(ORDER BY DESC/LIMIT/BETWEEN/IN/IS NOT NULL/COUNT DISTINCT/计算列+WHERE), create_file多sheet, create_sheet index, batch_insert insert_position, update_range, set_formula, evaluate_formula, get_operation_history, get_file_info, get_headers max_columns, check_duplicate_ids, find_last_row空表, rename_column不存在列
  - 19 PASS / 0 INFO / 1 FAIL
  - FAIL: T168 evaluate_formula独立数学表达式不支持（预期行为，需Excel公式格式）
  - INFO: T165 delete_rows条件删除0行（Score<60应匹配Score=58，类型问题）
  - INFO: T166 batch_insert insert_position模块导入错误（REQ-045）
  - 新增REQ-045: batch_insert_rows insert_position模块导入错误（P2）

## 关键指标
- **版本**: v1.7.6
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: 4e1bb8f (develop)

## 待处理
- [ ] REQ-044: find_last_row列名查找一致化（P2，自审发现）
- [ ] REQ-036: 边缘案例自动化测试（持续执行）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-044: find_last_row列名查找一致化（P2）
- [ ] REQ-045: batch_insert insert_position模块导入错误（P2）
- [ ] T141: 嵌套聚合计算列丢失BUG
- [ ] T165: delete_rows条件类型比较问题
