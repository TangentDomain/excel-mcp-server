# DECISIONS.md - 决策记录

## D008: REQ-029 MCP真实验证通过 (2026-03-28, R124+)
**需求**: REQ-029 JOIN表别名 + describe_table崩溃验证
**决策**: 两个P0 bug经验证均已修复，无需额外代码改动
**Bug 1 (JOIN表别名)**: ✅ `SELECT r.名称, s.名称 as 技能名称 FROM 角色 r JOIN 技能 s ON ...` 正确返回别名列名
**Bug 2 (streaming describe)**: ✅ streaming写入后describe_table正确识别列名（原测试脚本创建Excel方式有误导致误报）
**教训**: MCP真实验证时测试文件创建方式必须正确（ws.append(["A","B","C"]) 一行多列），错误创建方式（逐列ws.append(["A"])）会导致误判
**影响**: REQ-029标记为已验证完成，无需代码修改

## D004: 发现REQ-030 SQL引擎Bug (2026-03-27, R119)
**需求**: MCP真实验证 - 每5轮必做
**决策**: 进行MCP真实验证，发现3个SQL引擎边界问题
**原因**: 真实环境测试能暴露单元测试未覆盖的边界情况，影响用户体验
**发现**: 
- Bug 1: MAX/SUM聚合函数不支持多列表达式计算
- Bug 2: 标量子查询别名解析错误
- Bug 3: LEFT JOIN IS NULL结果过滤问题
**优先级**: 提升至P0，影响核心功能可用性
**影响**: 需要修复SQL引擎，增强聚合函数和子查询处理能力

## D005: REQ-030 修复方案 (2026-03-27, R120)
**需求**: REQ-030 SQL引擎Bug修复
**决策**: Bug 1和Bug 2已修复，Bug 3经验证无需修复
**Bug 1方案**: 新增`_is_expression`和`_evaluate_expression`方法，递归处理Add/Sub/Mul/Div/Literal表达式树
**Bug 2方案**: 在`_apply_select_expressions`和`_apply_group_by_aggregation`中新增Subquery处理分支，支持标量子查询
**Bug 3结论**: LEFT JOIN生成的NaN在pandas层正确保留，IS NULL/IS NOT NULL均能正确判断，无需修改
**影响**: 聚合函数现在支持`MAX(攻击力+防御力)`等表达式，标量子查询可在SELECT/WHERE/HAVING中使用

## D006: find_last_row降级路径处理dimension=None (2026-03-27, R123)
**需求**: REQ-015 流式写入后读取工具验证
**决策**: find_last_row在max_row/max_column为None时使用iter_rows降级遍历
**原因**: StreamingWriter使用write_only模式，不写<dimension>元数据，read_only模式下max_row返回None导致TypeError崩溃
**方案**: 检测None后改用sheet.iter_rows逐行遍历（read_only模式下仍为流式读取，内存高效），同时覆盖无column和有column两种路径
**影响**: 修复后流式写入的文件可正常被find_last_row、describe_table等依赖max_row的工具处理

## D007: check_duplicate_ids流式写入兼容 (2026-03-27, R123)
**需求**: REQ-015 流式写入后读取工具验证
**决策**: check_duplicate_ids中max_row/max_column为None时跳过边界检查
**原因**: 流式写入后read_only模式打开文件时dimension为None，header_row > None / col_idx > None 导致TypeError
**方案**: 两个边界检查都加上 `is not None` 前置条件
**影响**: 12项读取工具在流式写入后全部验证通过
