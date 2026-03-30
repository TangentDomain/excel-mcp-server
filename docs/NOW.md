# NOW.md - 第227轮

## 当前状态
- **轮次**: 第227轮（REQ-027 格式回归修复 + REQ-028 测试修复）
- **时间**: 2026-03-30 13:30 UTC

## 完成工作
- ✅ REQ-027 格式回归修复：20个 test_server.py 测试全部通过
  - _wrap: data+meta 字段展平到顶层，保持向后兼容
  - _strip_defaults: 有语义空列表不再被移除
  - excel_operations.get_all_headers: 修复从 data 中取 sheets 的路径
- ✅ REQ-028 测试修复：8个 test_insert_mode.py 测试全部通过
  - 列映射修正（C列=类型，D列=消耗MP）
  - 适配 get_range 返回 {coordinate, value} 格式
  - 工作表名补全

## 关键指标
- **版本**: v1.6.53
- **API测试**: 78/78 (server+insert_mode) ✅, 821/826 (全量)
- **PyPI**: TestPyPI token过期，需更新
- **合并**: main + develop 已推送，tag v1.6.53

## 待处理
- [ ] TestPyPI token 更新
- [ ] 5个其他测试失败（test_core/test_integration_comprehensive，非本轮引入）
- [ ] REQ-029 工程强化（P1，需 REQ-028 完成后开始）
