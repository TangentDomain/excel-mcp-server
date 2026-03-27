# REQ-015 MCP验证 - 修改操作流式写入

## 场景1: 大批量数据插入测试
excel_batch_insert_rows
- 文件: tests/test_data/large_game_config.xlsx
- 工作表: items
- 数据: 1000行装备数据
- 期望: 流式模式生效，内存占用低

## 场景2: 条件批量更新
excel_upsert_row  
- 文件: tests/test_data/character_data.xlsx
- 工作表: characters
- 条件: id="hero_001"
- 更新: level=100, exp=50000
- 期望: 流式模式生效

## 场景3: 批量删除行
excel_delete_rows
- 文件: tests/test_data/npc_data.xlsx  
- 工作表: npcs
- 范围: 第50-200行
- 期望: 流式模式生效，保留其他数据

## 场景4: 删除列操作
excel_delete_columns
- 文件: tests/test_data/item_properties.xlsx
- 工作表: properties  
- 列: 第3-5列（属性列）
- 期望: 流式模式生效

## 场景5: 覆盖范围更新
excel_update_range
- 文件: tests/test_data/skill_matrix.xlsx
- 工作表: skills
- 范围: B2:D100
- 数据: 新技能数值
- 期望: 流式覆盖模式生效

## 场景6: 大文件修改内存测试
excel_batch_insert_rows
- 文件: 10MB+ 配置表
- 数据: 5000行
- 监控: 内存使用情况
- 期望: 内存增长不超过100MB

## 场景7: 错误处理降级
excel_insert_rows
- 文件: 无calamine环境
- 期望: 自动降级到openpyxl，功能正常

## 场景8: 混合操作性能对比
- 传统模式: openpyxl全量加载
- 流式模式: calamine读取+write_only写入
- 对比: 大文件修改耗时和内存占用