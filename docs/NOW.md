# 第116轮 - REQ-025 返回值统一 + v1.6.5发布 ✅

## 状态
版本：v1.6.5 | 工具：44 | 测试：1107

## 本轮完成
- **REQ-025 返回值统一（get_headers）**：
  - 改进get_headers返回格式：新增结构化data字段（field_names/descriptions/dual_rows）
  - 新增meta元信息字段（sheet_name/header_row/header_count/dual_row_mode）
  - 保留顶层向后兼容字段（headers/field_names/descriptions/header_count/sheet_name/header_row）
  - 改进get_all_headers返回格式，新增data结构化字段
  - 更新README添加详细使用示例和返回格式说明
  - 更新test_server.py适配新data格式
  - 全量测试1107个全部通过
- **v1.6.5发布**：PyPI + GitHub推送完成
- **清理**：4个worktree已清理，/tmp测试文件已清理

## 待办
- [ ] REQ-025 继续其他工具返回值统一
- [ ] MCP真实验证（每5轮1次，当前未做）

## 决策
- **决策**：get_headers返回值同时包含新格式(data+meta)和旧格式(顶层字段)，确保向后兼容
- **原因**：现有测试和调用方依赖顶层字段，直接移除会导致大面积失败
- **方案**：新字段(data/meta)供新调用方使用，旧字段保持兼容，后续版本逐步deprecate