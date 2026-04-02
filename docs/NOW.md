# NOW.md - 第261轮

## 当前状态
- **轮次**: 第261轮
- **时间**: 2026-04-02

## 完成工作
- REQ-036: 边缘案例测试T296-T315（20个案例）
  - 换行符/Tab字符读写正常
  - SQL NULL/HAVING/DISTINCT/ORDER BY DESC LIMIT正常
  - 超长文本(32768字符)写入成功
  - 下划线开头/全数字Sheet名创建成功
  - 科学计数法公式设置成功
  - 修复T315: Sheet自身对比空Sheet NoneType错误（excel_compare.py）
  - T313 evaluate_formula已知限制（INFO）

## 关键指标
- **版本**: v1.7.10
- **测试**: 851 passed, MCP冒烟通过
- **Commit**: ba0c5f9 (develop)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行，P1）
- [ ] REQ-047: Sheet验证重复代码重构（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-048: 删除最后一个Sheet保护（P2）
