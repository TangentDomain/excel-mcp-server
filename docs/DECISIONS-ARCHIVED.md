## D023: REQ-032 SQL比较None值安全处理 (2026-03-27, R146)
**需求**: REQ-032 P0 bug修复
**问题**: SQL WHERE条件比较时，单元格值为None导致`'<=' not supported between instances of 'int' and 'NoneType'` TypeError
**根因**: `_COMPARISON_OPS`分发表中GT/GTE/LT/LTE lambda直接调用`float(l)`和`float(r)`，未处理None值
**决策**: 添加模块级`_safe_float_comparison`函数，None值时返回False
**方案**:
1. 在`advanced_sql_query.py`类定义前添加`_safe_float_comparison(left, right, op)`函数
2. `_COMPARISON_OPS`中的GT/GTE/LT/LTE改用该函数
3. 同时修复`excel_delete_rows`和`excel_batch_insert_rows`参数不匹配问题

