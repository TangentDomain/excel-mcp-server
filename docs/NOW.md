# NOW.md - 第252轮

## 当前状态
- **轮次**: 第252轮
- **时间**: 2026-04-02

## 完成工作
- 文档维护检查通过
- REQ-040: get_file_info区分数据范围和格式化范围
  - total_rows/total_cols改为仅反映实际数据维度
  - 格式化范围不同时额外报告formatted_rows/formatted_cols
  - 移除read_only模式以支持准确计算
- REQ-037: 标记DONE（formula_cache已有threading.RLock保护）
- REQ-036: 边缘案例测试33个全通过(T53-T85)
- v1.7.5发布到PyPI

## 关键指标
- **版本**: v1.7.5 (已发布PyPI)
- **测试**: 851 passed + MCP冒烟测试通过
- **Commit**: 8a8abb2 (develop), c82ebe5 (main)

## 待处理
- [ ] REQ-036: 边缘案例自动化测试（持续执行）
- [ ] REQ-035: 硬编码常量提取为配置项（P2）
