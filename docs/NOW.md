# NOW.md - 第258轮

## 当前状态
- **轮次**: 第258轮
- **时间**: 2026-04-02

## 完成工作
- REQ-036: 边缘案例测试T231-T255
  - 发现并修复format_cells_user_friendly BUG：调用不存在的update_range_range方法→改为format_cells
  - 验证数据验证(set_data_validation/clear_validation)正常工作
  - 验证条件格式(add/clear_conditional_format)基本正常
  - 验证图表(create_chart bar/line)正常工作
  - 验证user_friendly API(get/update_range_user_friendly)正常工作
  - 验证SQL高级查询(CASE WHEN/IN/LIKE/COUNT DISTINCT/FROM子查询/HAVING/BETWEEN)全部正常
  - 验证备份操作(create_backup/list_backups)正常工作
  - v1.7.9发布

## 关键指标
- **版本**: v1.7.9
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: 008cdba (develop)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行，P1）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
