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

## D026: REQ-026 文档与门面优化 (2026-03-27, R153)
**需求**: REQ-026 文档与门面优化 - 项目文档持续优化
**问题**: 项目文档版本信息不一致，CHANGELOG更新滞后
**根因**: 每轮发布后文档更新流程不够自动化，版本同步存在延迟
**决策**: 建立标准化的文档更新流程，确保版本信息实时同步
**方案**:
1. README.md/README.en.md: 统一更新版本号到最新发布版本
2. CHANGELOG.md: 每轮发布后自动添加新版本条目
3. 文档门面优化：保持中英文文档版本完全一致
4. 建立文档版本一致性检查机制
**验证**: 文档版本信息同步完成，用户体验提升