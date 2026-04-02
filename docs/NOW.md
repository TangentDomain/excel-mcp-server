# NOW.md - 第255轮

## 当前状态
- **轮次**: 第255轮
- **时间**: 2026-04-02

## 完成工作
- REQ-036: 边缘案例测试T131-T160（30个API工具测试）
  - 覆盖工具：SQL函数(TRIM/ROUND/ABS/SUBSTR/REPLACE/CAST/MIN/MAX/COUNT/SUM/AVG), LIKE(下划线)/NOT LIKE, 多聚合GROUP BY, 嵌套聚合, 空范围, 多条件AND WHERE, 空表查询, get_headers合并/双行模式, search大小写不敏感, SQL(!=/<>), batch_insert dict, 日期比较, 单单元格get_range, find_last_row非流式, 除零保护, (OR+AND)组合, rename_column后查询, 多列排序混合, get_file_info精度
  - 25 PASS / 3 INFO / 2 FAIL
  - 发现BUG：嵌套聚合表达式`SUM(Score)/COUNT(*)`计算列丢失（T141）
  - 功能缺失：ROUND/ABS数学函数不支持（T132/T133）
  - get_headers双行模式误判（T146）

## 关键指标
- **版本**: v1.7.6
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: ef73ad8 (develop)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] T141: 嵌套聚合计算列丢失BUG
