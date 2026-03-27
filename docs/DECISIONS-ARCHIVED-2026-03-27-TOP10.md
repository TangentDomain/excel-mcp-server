# DECISIONS.md - 决策记录

## D002: 全局_wrap自动补充message + 顶层字段向后兼容 (2026-03-27, R118)
**需求**: REQ-025 AI体验优化 - 返回值统一
**决策**: _wrap成功时自动补充message，新格式与旧格式并存
**原因**: Operations层返回的结果可能缺少message导致格式不统一；直接移除顶层字段会破坏测试和用户代码
**方案**: (1)_wrap成功时若无message默认"操作成功" (2)顶层保留旧字段(sheets/query_info/validation_info)，同时新增data/meta
**影响**: 所有通过_wrap包装的工具统一为{success, message, data, meta}，旧字段继续可用

## D001: get_headers返回值统一 (2026-03-27, R116)
