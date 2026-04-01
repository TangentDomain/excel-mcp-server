# NOW.md - 第245轮

## 当前状态
- **轮次**: 第245轮
- **时间**: 2026-04-01

## 完成工作
- 文档维护检查通过（REQ-038移至ARCHIVED）
- CI检查通过（green）
- REQ-036: 边缘案例测试（第2轮）
  - 测试10个新案例：9通过1失败
  - 发现REQ-041：SQL含空格列名返回列头字符串而非实际值
- REQ-041: SQL列名映射修复
  - 新增_preprocess_quoted_identifiers方法
  - SQL解析前将"Original Name"替换为`cleaned_name`
  - 用户现在可用SELECT "Player Name"正确查询含空格列
- v1.6.60发布到PyPI

## 关键指标
- **版本**: v1.6.60 (已发布PyPI)
- **测试**: 851 passed + MCP冒烟测试通过
- **Commit**: 7734ad0 (develop), 359c0a6 (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（每轮执行1个新案例）
- [ ] REQ-034: 路径验证逻辑抽取为装饰器（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-037: formula_cache并发访问保护（P2）
- [ ] REQ-039: list_sheets不区分隐藏工作表（P2）
- [ ] REQ-040: 稀疏sheet维度信息不准确（P2）
