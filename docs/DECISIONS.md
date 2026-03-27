# DECISIONS.md - 决策记录

## D001: get_headers返回值统一 (2026-03-27, R116)
**需求**: REQ-025 AI体验优化 - 返回值统一
**决策**: get_headers同时返回新格式(data+meta)和旧格式(顶层字段)
**原因**: 现有测试和调用方依赖顶层headers/field_names/descriptions字段，直接移除会破坏兼容性
**方案**: 新增data结构化字段和meta元信息，保留顶层向后兼容字段
**影响**: API同时支持新旧两种访问方式，后续版本逐步deprecate旧字段