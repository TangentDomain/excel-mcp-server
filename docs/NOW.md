# NOW.md - 第244轮

## 当前状态
- **轮次**: 第244轮
- **时间**: 2026-04-01

## 完成工作
- 文档维护检查通过（删除scripts/test_api_issues.py）
- CI检查通过（green）
- REQ-038: 工作表名称非法字符拒绝+超长截断改为报错
  - 拆分_normalize_sheet_name为_validate_sheet_name+_sanitize_sheet_name
  - create_sheet/rename_sheet拒绝非法名称，返回明确错误信息
  - copy_sheet自动生成名称允许静默清理
  - 更新测试适配新行为
- v1.6.59发布到PyPI

## 关键指标
- **版本**: v1.6.59 (已发布PyPI)
- **测试**: 851 passed + MCP冒烟测试通过
- **Commit**: 951e065 (develop), c088acb (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（每轮执行1个新案例）
- [ ] REQ-034: 路径验证逻辑抽取为装饰器（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-037: formula_cache并发访问保护（P2）
- [ ] REQ-039: list_sheets不区分隐藏工作表（P2）
- [ ] REQ-040: 稀疏sheet维度信息不准确（P2）
