# NOW.md - 第247轮

## 当前状态
- **轮次**: 第247轮
- **时间**: 2026-04-02

## 完成工作
- 文档维护检查通过（REQ-039/041归档，版本一致）
- CI检查通过（green）
- REQ-036: 边缘案例测试21 - SQL双引号标识符与字符串字面量冲突
- REQ-042: 修复_preprocess_quoted_identifiers两个BUG
  - AST方法精确替换：只替换列引用位置（SELECT/ORDER BY/GROUP BY/HAVING），WHERE值位置保持不变
  - 新增_col_map_cache解决缓存命中时列名映射丢失
- v1.7.1发布到PyPI

## 关键指标
- **版本**: v1.7.1 (已发布PyPI)
- **测试**: 851 passed + MCP冒烟测试通过
- **Commit**: bd729a1 (develop), cd1783a (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（每轮执行1个新案例）
- [ ] REQ-034: 路径验证逻辑抽取为装饰器（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-037: formula_cache并发访问保护（P2）
- [ ] REQ-040: 稀疏sheet维度信息不准确（P2）
