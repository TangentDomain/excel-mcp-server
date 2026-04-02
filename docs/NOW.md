# NOW.md - 第251轮

## 当前状态
- **轮次**: 第251轮
- **时间**: 2026-04-02

## 完成工作
- 文档维护检查通过
- REQ-043: 安全回归修复 - 为所有接受file_path的MCP工具添加路径遍历保护
  - 10个函数添加@_validate_file_path装饰器
  - 2个函数(excel_merge_files/excel_merge_multiple_files)添加内联_validate_path调用
  - excel_merge_files补充output_path验证
- v1.7.4发布到PyPI

## 关键指标
- **版本**: v1.7.4 (已发布PyPI)
- **测试**: 851 passed + MCP冒烟测试通过
- **Commit**: 25cfd24 (develop), dd93f7e (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行）
- [ ] REQ-034: 路径验证逻辑抽取为装饰器（P2）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
- [ ] REQ-037: formula_cache并发访问保护（P2）
- [ ] REQ-040: 稀疏sheet维度信息不准确（P2）
