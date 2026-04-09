# ExcelMCP 项目重构需求

## 背景
质量抽检发现ExcelMCP项目存在结构问题，影响可维护性和标准化程度。虽然当前功能正常，但项目结构需要重构以符合Python包标准。

## 需求详情

### REQ-EXCEL-001: Python包标准化重构
- **类型**: 项目重构
- **优先级**: P1
- **状态**: DONE
- **创建时间**: 2026-04-08
- **来源**: 自迭代质量抽检反馈

### 具体要求
1. **目录结构重组**
   - 创建标准的Python包结构: `src/excel_mcp/` 包目录
   - 将现有Python模块移入包内
   - 建立正确的`__init__.py`文件结构

2. **包管理文件完善**
   - 完善pyproject.toml配置
   - 添加依赖版本锁定
   - 定义正确的包入口点

3. **项目结构规范化**
   - 添加setup.py作为备选安装方式
   - 创建requirements.txt锁定生产依赖
   - 添加development requirements for dev依赖

4. **模块重构**
   - 保持现有API接口不变
   - 重构内部模块组织
   - 确保所有测试继续通过

### 验收标准
- [ ] 项目符合Python包标准布局
- [ ] 可通过`pip install .`正确安装
- [ ] 所有现有功能保持不变
- [ ] 测试套件全部通过
- [ ] 文档同步更新

### 影响评估
- **风险**: 低（重构不改变API）
- **工作量**: 中等（需要重新组织结构）
- **收益**: 提升可维护性和标准化程度
REQ-EXCEL-002 (P2) - 添加自动化测试配置
=======================
来源: 质量抽检反馈 #6
需求: 在package.json中添加test脚本，建立GitHub Actions CI/CD，为核心API功能添加单元测试验证
验证规则: 
- ✅ npm test 命令可执行
- ✅ GitHub Actions配置文件存在
- ✅ 核心API功能有测试覆盖
优先级: P2 (体验改进)

---

## 用户反馈需求（2026-04-09 测试报告）

> 来源：用户提供的ExcelMCP测试报告（complex_test.xlsx + 问题报告）
> 测试日期：2026-04-09 | 测试SQL查询25+条 | 工具函数15+个 | 发现问题12个

### REQ-EXCEL-003 (P0): IN/NOT IN 操作符内部错误
- **状态**: DONE
- **来源**: L1 用户直接反馈（Bug #1）
- **问题**: `AdvancedSQLQueryEngine._in_to_pandas() got an unexpected keyword argument 'negate'`
- **影响**: 所有使用IN和NOT IN的查询全部失败
- **复现**: `SELECT 技能名称, 伤害 FROM 技能配置 WHERE 技能ID IN (1, 3, 5, 7, 9)`
- **验收**: IN和NOT IN查询正常执行并返回正确结果

### REQ-EXCEL-004 (P0): EXISTS 子查询返回错误结果
- **状态**: DONE
- **来源**: L1 用户直接反馈（Bug #2）
- **问题**: EXISTS子查询返回所有行而非过滤后的结果
- **复现**: `SELECT ... WHERE EXISTS (SELECT 1 FROM s2 WHERE s2.技能ID = 技能配置.技能ID AND s2.伤害 > 200)`
- **验收**: EXISTS子查询只返回满足条件的行
- **修复时间**: 2026-04-09

### REQ-EXCEL-005 (P1): SELECT 子句不支持计算表达式
- **状态**: OPEN
- **来源**: L1 用户直接反馈（限制 #1）
- **问题**: `SELECT 技能名称, (伤害 * 1.2) as 预期伤害` 报错"不支持的表达式"
- **验收**: SELECT中支持算术运算(* / + -)

### REQ-EXCEL-006 (P1): WHERE 子句不支持算术表达式
- **状态**: OPEN
- **来源**: L1 用户直接反馈（限制 #2）
- **问题**: `WHERE 力量 + 敏捷 + 智力 > 180` 报错"不支持的表达式类型"
- **验收**: WHERE中支持算术运算比较

### REQ-EXCEL-007 (P1): ORDER BY 不能使用 SELECT 别名
- **状态**: OPEN
- **来源**: L1 用户直接反馈（限制 #3）
- **问题**: ORDER BY用别名报"列不存在"
- **验收**: ORDER BY支持SELECT中定义的别名

### REQ-EXCEL-008 (P1): JOIN 只支持等值连接
- **状态**: OPEN
- **来源**: L1 用户直接反馈（限制 #4）
- **问题**: `ON s.等级限制 <= e.等级限制` 报"请使用等值连接"
- **验收**: JOIN ON支持非等值比较(<= >= < >)

### REQ-EXCEL-009 (P1): CTE 不支持复杂表达式
- **状态**: OPEN
- **来源**: L1 用户直接反馈（限制 #5）
- **问题**: WITH子句中使用算术表达式报错
- **验收**: CTE中支持计算表达式

### REQ-EXCEL-010 (P2): batch_update_ranges 参数格式不一致
- **状态**: DONE
- **来源**: L1 用户直接反馈（工具问题 #1）
- **问题**: updates参数需要sheet_name但文档不清晰，range含sheet名时报错
- **验收**: 参数格式统一，文档明确

### REQ-EXCEL-011 (P2): set_data_validation / format_cells 缺少 sheet_name 提示
- **状态**: DONE
- **来源**: L1 用户直接反馈（工具问题 #2/#3）
- **问题**: 必需参数未在错误提示中明确说明
- **验收**: 缺少必需参数时给出明确提示

### REQ-EXCEL-012 (P2): write_only_override 返回类型错误
- **状态**: DONE
- **来源**: L1 用户直接反馈（工具问题 #4）
- **问题**: 返回OperationResult对象而非dict
- **验收**: 返回类型一致且符合文档说明

### REQ-EXCEL-013 (P2): add_conditional_format 不支持的格式类型
- **状态**: DONE
- **来源**: L1 用户直接反馈（工具问题 #5）
- **问题**: 文档说支持highlight类型但实际只支持cellValue和formula
- **验收**: 要么实现highlight类型，要么文档与实际一致

---

## 窗口函数扩展（SQL差距分析 2026-04-09）

### REQ-EXCEL-014 (P0): LAG / LEAD 窗口函数
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: 缺少LAG/LEAD，无法取前N行/后N行的值（如对比前后等级属性变化）
- **验收**: `LAG(col, N) OVER (PARTITION BY ... ORDER BY ...)` 和 `LEAD(col, N) OVER (...)` 可用
- **替代方案**: 自JOIN偏移行（复杂度高）

### REQ-EXCEL-015 (P0): GROUP_CONCAT 聚合函数
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: 缺少GROUP_CONCAT，无法在分组内拼接字符串（如同组内拼接所有技能名）
- **验收**: `SELECT dept, GROUP_CONCAT(name) FROM t GROUP BY dept` 可用
- **替代方案**: 应用层聚合

### REQ-EXCEL-016 (P0): FIRST_VALUE / LAST_VALUE 窗口函数
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: 缺少FIRST_VALUE/LAST_VALUE，无法取分组内首/末行值
- **验收**: `FIRST_VALUE(col) OVER (PARTITION BY ... ORDER BY ...)` 可用
- **替代方案**: 子查询+LIMIT 1（复杂度高）

### REQ-EXCEL-017 (P1): NTILE 窗口函数
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: 缺少NTILE，无法将分组均分为N桶
- **验收**: `NTILE(N) OVER (PARTITION BY ... ORDER BY ...)` 可用

### REQ-EXCEL-018 (P1): PERCENT_RANK / CUME_DIST 窗口函数
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: 缺少百分比排名和累积分布函数
- **验收**: `PERCENT_RANK() OVER (...)` 和 `CUME_DIST() OVER (...)` 可用

### REQ-EXCEL-019 (P1): COUNT() OVER () 窗口聚合
- **状态**: OPEN
- **来源**: SQL差距分析报告
- **问题**: COUNT()作为窗口函数不可用（仅作为普通聚合可用）
- **验收**: `COUNT(*) OVER (PARTITION BY ...)` 返回分组内总数

