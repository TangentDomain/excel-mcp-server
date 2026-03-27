# 第146轮 - REQ-032 P0 bug修复轮

---

## 状态
版本：v1.6.25 | 工具：44 | 测试：1159

## 本轮完成
- **REQ-032 P0 bug修复** ✅
  - Bug 1：SQL WHERE条件比较None值TypeError → 添加`_safe_float_comparison`函数
  - Bug 2：`excel_delete_rows`新增condition参数，支持SQL条件删除
  - Bug 3：`excel_batch_insert_rows`新增insert_position/condition参数，支持定位插入
  - 全量测试1159通过，PyPI v1.6.25已发布

## 修复详情
- `_safe_float_comparison`：模块级函数，处理None值和类型转换异常
- `excel_delete_rows(condition=...)`：自动查询匹配行号→从后往前删除避免偏移
- `excel_batch_insert_rows(insert_position=..., condition=...)`：先插入空行→逐行写入数据
- 新增`ExcelOperations.batch_insert_rows_at`方法

## 下轮待办
- [ ] MCP真实验证（验证Bug修复效果）
- [ ] REQ-006 工具描述持续优化
