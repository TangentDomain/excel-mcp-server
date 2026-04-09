# ExcelMCP 问题报告

测试日期：2026-04-09
测试文件：`C:/Users/Administrator/complex_test.xlsx`

---

## 严重 Bug

### Bug 1: IN 操作导致内部错误

**问题描述**：使用 `IN` 操作符时，SQL 引擎抛出内部异常。

**错误信息**：
```
SQL执行错误: AdvancedSQLQueryEngine._in_to_pandas() got an unexpected keyword argument 'negate'
```

**复现方式**：
```python
excel_query(
    file_path="complex_test.xlsx",
    query_expression="SELECT 技能名称, 伤害, 冷却时间 FROM 技能配置 WHERE 技能ID IN (1, 3, 5, 7, 9)"
)
```

**影响范围**：所有使用 `IN` 和 `NOT IN` 的查询

---

### Bug 2: EXISTS 子查询返回错误结果

**问题描述**：EXISTS 子查询返回了所有行，而不是按条件过滤的结果。

**复现方式**：
```python
excel_query(
    file_path="complex_test.xlsx",
    query_expression="""SELECT 技能名称, 技能类型, 伤害 FROM 技能配置
    WHERE EXISTS (SELECT 1 FROM 技能配置 s2 WHERE s2.技能ID = 技能配置.技能ID AND s2.伤害 > 200)"""
)
```

**预期结果**：只返回伤害 > 200 的技能
**实际结果**：返回了所有技能（包括伤害 <= 200 的）

---

## SQL 功能限制

### 限制 1: SELECT 子句中不支持计算表达式

**问题描述**：在 SELECT 子句中使用算术表达式会报错。

**错误信息**：
```
处理SELECT表达式失败: 不支持的表达式: (伤害 * 1.2)
```

**复现方式**：
```sql
SELECT 技能名称, 伤害, (伤害 * 1.2) as 预期伤害
FROM 技能配置
WHERE 伤害 > 0
```

**影响范围**：所有需要在 SELECT 中进行计算的场景

---

### 限制 2: WHERE 子句中不支持算术表达式

**问题描述**：在 WHERE 子句中使用算术表达式进行比较会报错。

**错误信息**：
```
不支持的表达式类型: <class 'sqlglot.expressions.Add'>
```

**复现方式**：
```sql
SELECT 角色名, 力量, 敏捷, 智力
FROM 角色属性
WHERE 力量 + 敏捷 + 智力 > 180
```

**变通方法**：使用 FROM 子查询先计算，再过滤
```sql
SELECT * FROM (
    SELECT 角色名, 力量, 敏捷, 智力, (力量 + 敏捷 + 智力) as 总属性
    FROM 角色属性
) WHERE 总属性 > 180
```

---

### 限制 3: ORDER BY 不能使用 SELECT 别名

**问题描述**：在 ORDER BY 中使用 SELECT 定义的别名会报错"列不存在"。

**错误信息**：
```
排序列 '总战力' 不存在.可用列: ['怪物ID', '怪物名称', ...]
```

**复现方式**：
```sql
SELECT 怪物名称, 血量, 攻击力, 防御力, (血量 + 攻击力 + 防御力) as 总战力
FROM 怪物表
WHERE 怪物类型 = 'Boss'
ORDER BY 总战力 DESC
```

**变通方法**：在 ORDER BY 中重复表达式
```sql
SELECT 怪物名称, 血量, 攻击力, 防御力
FROM 怪物表
WHERE 怪物类型 = 'Boss'
ORDER BY 血量 + 攻击力 + 防御力 DESC
```

---

### 限制 4: JOIN 只支持等值连接

**问题描述**：JOIN ON 条件不支持非等值比较（如 <=, >=, <, >）。

**错误信息**：
```
JOIN ON条件格式不支持,请使用等值连接: ON a.id = b.id
```

**复现方式**：
```sql
SELECT s.技能名称, e.装备名称
FROM 技能配置 s
INNER JOIN 装备表 e
ON s.等级限制 <= e.等级限制
WHERE s.伤害 > 150
```

**影响范围**：需要基于范围进行关联的场景

---

### 限制 5: CTE 中不支持复杂表达式

**问题描述**：在 WITH 子句（CTE）中使用算术表达式会报错。

**错误信息**：
```
CTE '高级角色' 执行失败: 处理SELECT表达式失败: 不支持的表达式: (力量 + 敏捷 + 智力)
```

**复现方式**：
```sql
WITH 高级角色 AS (
    SELECT 角色名, (力量 + 敏捷 + 智力) as 总属性
    FROM 角色属性
    WHERE 等级 >= 25
)
SELECT * FROM 高级角色 WHERE 总属性 > 150
```

---

## 工具参数问题

### 问题 1: batch_update_ranges 参数格式不一致

**问题描述**：`updates` 参数中需要包含 `sheet_name`，但文档不清晰。

**错误调用**：
```python
excel_batch_update_ranges(
    file_path="complex_test.xlsx",
    updates=[
        {"range": "装备表!D2", "data": [["普通"]]},
        {"range": "装备表!D4", "data": [["史诗"]]}
    ]
)
```
**错误**：`装备表!D2 is not a valid coordinate or range`

**正确调用**：
```python
excel_batch_update_ranges(
    file_path="complex_test.xlsx",
    updates=[
        {"range": "D2", "sheet_name": "装备表", "data": [["普通"]]},
        {"range": "D4", "sheet_name": "装备表", "data": [["史诗"]]}
    ]
)
```

---

### 问题 2: set_data_validation 缺少 sheet_name 参数

**问题描述**：工具要求 `sheet_name` 参数，但在一些调用模式中未明确提示。

**正确调用**：
```python
excel_set_data_validation(
    file_path="complex_test.xlsx",
    sheet_name="技能配置",  # 必需参数
    range_address="技能配置!C2:C20",
    validation_type="list",
    criteria="攻击,防御,辅助"
)
```

---

### 问题 3: format_cells 缺少 sheet_name 参数

**问题描述**：工具要求 `sheet_name` 参数，但文档不清晰。

**正确调用**：
```python
excel_format_cells(
    file_path="complex_test.xlsx",
    sheet_name="技能配置",  # 必需参数
    range="A1:H1",
    formatting={"bold": True, "font_size": 14}
)
```

---

### 问题 4: write_only_override 返回类型错误

**问题描述**：工具返回 `OperationResult` 对象而不是字典，可能导致类型检查错误。

**复现方式**：
```python
result = excel_write_only_override(
    file_path="complex_test.xlsx",
    sheet_name="测试",
    range_spec="测试!A1:C4",
    data=[["测试ID", "测试名称", "测试值"], [1, "A", 100]]
)
# result 是 OperationResult 类型，不是 dict
```

---

### 问题 5: add_conditional_format 不支持的格式类型

**问题描述**：文档说明支持 `highlight` 类型，但实际使用时报错。

**错误信息**：
```
不支持的格式类型
```

**复现方式**：
```python
excel_add_conditional_format(
    file_path="complex_test.xlsx",
    sheet_name="技能配置",
    range_address="技能配置!D2:D20",
    format_type="highlight",  # 不支持
    criteria="伤害>200",
    format_style="lightRed"
)
```

**支持的类型**：根据错误提示，仅支持 `cellValue` 和 `formula`

---

## 测试数据

测试使用的 Excel 文件包含以下工作表：

| 工作表 | 行数 | 说明 |
|--------|------|------|
| 技能配置 | 17行 | 包含技能ID、名称、类型、伤害、冷却时间、法力消耗、等级限制 |
| 装备表 | 12行 | 包含装备ID、名称、部位、稀有度、攻击力、防御力、价格等 |
| 怪物表 | 13行 | 包含怪物ID、名称、等级、血量、攻击力、防御力等 |
| 角色属性 | 10行 | 包含角色ID、名称、职业、等级、属性值等 |

---

## 建议

1. **修复 IN/NOT IN 操作符的内部错误**：这是高优先级 Bug，影响常用查询
2. **改进 EXISTS 子查询逻辑**：确保返回正确的过滤结果
3. **统一工具参数格式**：明确哪些工具需要 `sheet_name` 参数
4. **改进错误消息**：当 ORDER BY 使用别名时，给出更明确的提示
5. **完善文档**：更新工具参数说明，特别是 `batch_update_ranges` 的格式
6. **考虑支持更多 SQL 特性**：SELECT 计算表达式、非等值 JOIN 等

---

## 测试统计

| 项目 | 数量 |
|------|------|
| 测试的 SQL 查询 | 25+ |
| 测试的工具函数 | 15+ |
| 发现的问题 | 12 |
| 严重 Bug | 2 |
| SQL 限制 | 5 |
| 工具参数问题 | 5 |
