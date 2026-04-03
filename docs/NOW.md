# NOW.md - 第266轮

## 当前状态
- **轮次**: 第266轮
- **时间**: 2026-04-03

## 完成工作
- REQ-036: 边缘案例测试T396-T415（20个案例，15通过4失败1错误0实际BUG）
  - 发现并修复BUG：excel_list_charts的chart.chart_type→chart.type，chart.data.srcDataSource不存在，
    chart.title.text需要从RichText提取纯文本，chart.dLbls空值保护
  - 发现并修复BUG：excel_clear_validation用dv.formula1匹配范围(应使用dv.sqref)
  - 修正3处"工作表不存在"错误码OPERATION_FAILED→SHEET_NOT_FOUND
  - 发布v1.7.14

## 关键指标
- **版本**: v1.7.14
- **测试**: 851 passed + MCP冒烟通过
- **Commit**: 61a405b (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行，P1）
- [ ] REQ-047: Sheet验证重复代码重构（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-049: Docstring合规率提升（P2）
