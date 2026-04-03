# NOW.md - 第264轮

## 当前状态
- **轮次**: 第264轮
- **时间**: 2026-04-03

## 完成工作
- REQ-036: 边缘案例测试T356-T375（20个案例，17通过3信息0失败）
  - server_stats/operation_history/search_directory/describe_table全部正常
  - clear_validation/clear_conditional_format正常
  - upsert_row更新/插入正常（需用列名非列字母）
  - convert_format xlsx→json正常
  - copy_sheet/rename_column/format_cells/data_validation正常
  - SQL CASE WHEN正常
  - BUG修复：DataValidationError 3参数调用→2参数（validators.py）
  - BUG修复：write_only_override降级路径ExcelWriter缺失导入（server.py）
  - 发布v1.7.12

## 关键指标
- **版本**: v1.7.12
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: 205fbe2 (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行，P1）
- [ ] REQ-047: Sheet验证重复代码重构（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-049: Docstring合规率提升（P2）
