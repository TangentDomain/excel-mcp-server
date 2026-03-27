## D025: REQ-015 copy_sheet streaming支持 (2026-03-27, R152)
**需求**: REQ-015 write_only覆盖修改操作 - copy_sheet streaming
**问题**: excel_copy_sheet工具缺少streaming参数，大文件复制性能差
**根因**: 早期streaming改造集中在update/insert/delete操作，copy_sheet遗漏
**决策**: 为copy_sheet添加streaming参数，使用calamine+write_only实现流式复制
**方案**:
1. server.py: excel_copy_sheet新增streaming参数（默认True）
2. ExcelOperations.copy_sheet: 透传streaming参数
3. ExcelManager.copy_sheet: 新增_copy_sheet_streaming方法
4. 使用calamine读取所有工作表数据 → write_only重建（保留其他工作表）
5. 保留源工作表列宽、自动降级到openpyxl、名称冲突自动编号
**验证**: 5个新测试+1159个已有测试全部通过（共1164），PyPI v1.6.27发布