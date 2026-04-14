# Excel MCP Server - 开发指南

## 📋 项目概述

游戏开发专用的 Excel 配置表 MCP 服务器，支持高级 SQL 查询、批量操作、跨文件 JOIN 等功能。

## 🧪 测试方法论

### 核心原则：**先 Python API 测试，再 MCP 工具验证**

在开发新功能时，遵循以下测试流程：

1. **实现阶段**：直接修改 `src/` 代码
2. **验证阶段**：通过 Python API 直接测试（无需重启 MCP）
3. **集成阶段**：确认功能正常后，再通过 MCP 工具验证

### 模拟 MCP 工具测试

在开发过程中，使用 Python API 直接调用模拟 MCP 工具行为：

```python
from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query
)

# 模拟 excel_query 工具
result = execute_advanced_sql_query(
    file_path,
    "SELECT ID, Name, ROUND(Price, 0) as Price FROM 装备"
)

# 模拟 excel_update_query 工具
result = execute_advanced_update_query(
    file_path,
    "UPDATE 装备 SET Price = ROUND(Price * 1.1, 2) WHERE Rarity = 'Epic'"
)
```

### 完整测试模板

```python
"""
功能测试模板 - 复制此模板进行新功能测试
"""
from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query
)

file_path = '/path/to/test.xlsx'

print("=" * 70)
print("🔧 新功能测试")
print("=" * 70)

tests = [
    ("测试1: 基础功能", "SELECT ..."),
    ("测试2: WHERE 条件", "SELECT ... WHERE ..."),
    ("测试3: ORDER BY", "SELECT ... ORDER BY ..."),
    ("测试4: UPDATE", "UPDATE ..."),
]

passed = 0
for name, sql in tests:
    is_update = sql.strip().upper().startswith('UPDATE')

    if is_update:
        result = execute_advanced_update_query(file_path, sql)
    else:
        result = execute_advanced_sql_query(file_path, sql)

    if result['success']:
        print(f"✅ {name}")
        passed += 1
    else:
        print(f"❌ {name}")
        print(f"   错误: {result.get('message', '')[:60]}")

print(f"\n🎉 结果: {passed}/{len(tests)} 测试通过")
```

### 测试数据准备

创建真实场景的测试数据：

```python
import pandas as pd
from openpyxl import Workbook
import random

# 创建游戏装备配置表
wb = Workbook()
ws = wb.active
ws.title = "装备"

# 表头
ws.append(["ID", "Name", "BaseAtk", "AtkBonus", "Price", "Rarity"])

# 生成测试数据
for i in range(1, 51):
    ws.append([
        i,
        f"Item-{i}",
        random.randint(10, 100),
        round(random.uniform(5.5, 45.8), 2),
        round(random.uniform(50.5, 9999.99), 2),
        random.choice(["Common", "Rare", "Epic", "Legendary"])
    ])

wb.save('/tmp/test_data.xlsx')
```

## 🌐 MCP 工具直接测试

### 测试场景设计

通过 Claude Code 直接调用 MCP 工具进行真实场景测试：

**基础功能测试**:
```sql
-- 简单查询
SELECT * FROM 装备配置 WHERE Rarity = 'Legendary'

-- 排序和限制
SELECT Name, Price FROM 装备配置 ORDER BY Price DESC LIMIT 5

-- 聚合统计
SELECT Rarity, COUNT(*) as Count, AVG(Price) as AvgPrice
FROM 装备配置 GROUP BY Rarity
```

**高级功能测试**:
```sql
-- CTE 子查询
WITH HighValueItems AS (
    SELECT * FROM 装备配置 WHERE Price > 1000
)
SELECT * FROM HighValueItems WHERE BaseAtk > 50

-- 窗口函数
SELECT Name, Price,
       RANK() OVER (ORDER BY Price DESC) as PriceRank
FROM 装备配置

-- 多表 JOIN
SELECT e.Name, m.Name, d.DropRate
FROM 装备配置 e
JOIN 掉落配置 d ON e.ID = d.ItemID
JOIN 怪物配置 m ON d.MonsterID = m.ID
```

**批量操作测试**:
```sql
-- 批量更新
UPDATE 装备配置 SET Price = ROUND(Price * 1.1, 2)
WHERE Rarity IN ('Epic', 'Legendary')

-- 批量插入
INSERT INTO 装备配置 (ID, Name, BaseAtk, Price, Rarity)
VALUES (100, 'Excalibur', 150, 9999.99, 'Legendary')
```

### MCP 连接诊断

**检查 MCP 服务器状态**:
```bash
# 使用诊断脚本
./verify_mcp.sh

# 或手动测试
echo '{"jsonrpc":"2.0","id":1,"method":"ping"}' | \
  uv run excel-mcp-server-fastmcp --stdio 2>&1 | head -5
```

**常见连接问题**:

1. **路径问题**
   - 症状: `Cannot find module`
   - 解决: `.mcp.json` 使用相对路径 `"."`

2. **导入错误**
   - 症状: `cannot import name 'main'`
   - 解决: 检查 `__init__.py` 是否导出 main 函数

3. **环境变量**
   - 调试: 添加 `PYTHONUNBUFFERED=1` 和 `MCP_DEBUG=1`

### 极限测试场景

验证服务器在高负载下的表现：

```sql
-- 复杂嵌套查询
WITH RankedItems AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY Rarity ORDER BY Price DESC) as rn
    FROM 装备配置
),
Filtered AS (
    SELECT * FROM RankedItems WHERE rn <= 3
)
SELECT e.Name, e.Price,
       (SELECT COUNT(*) FROM 掉落配置 d WHERE d.ItemID = e.ID) as DropCount
FROM Filtered e
WHERE e.Price > (SELECT AVG(Price) * 2 FROM 装备配置)
```

**性能指标**:
- 简单查询: < 10ms
- 复杂 JOIN: < 50ms
- 窗口函数: < 30ms
- 批量更新: < 100ms

## 🔧 开发工作流

### 1. 功能开发

```bash
# 1. 修改源代码
vim src/excel_mcp_server_fastmcp/api/advanced_sql_query.py

# 2. 运行回归测试
python -m pytest tests/ --ignore=tests/test_pivot_table.py -q

# 3. 使用 Python API 验证新功能
uv run python test_new_feature.py
```

### 2. 调试技巧

```python
# 启用详细日志
import logging
logging.basicConfig(level=logging.DEBUG)

# 查看返回数据结构
result = execute_advanced_sql_query(file_path, sql)
print(f"Success: {result['success']}")
print(f"Data: {result['data'][:3]}")  # 查看前3行
print(f"Message: {result.get('message', '')}")
```

### 3. 常见问题排查

**问题：isinstance() arg 2 must be a type**
- **原因**：使用了 `frozenset` 而非 `tuple`
- **解决**：`_FUNCS = frozenset({exp.Round})` → `_FUNCS = (exp.Round,)`

**问题：MCP 工具返回旧代码错误**
- **原因**：Python 字节码缓存
- **解决**：清除缓存后重启
```bash
find src -type d -name "__pycache__" -exec rm -rf {} +
find src -name "*.pyc" -delete
```

**问题：双行表头导致聚合查询错误**
- **原因**：中英文表头别名解析冲突
- **解决**：使用单行表头测试，或检查别名映射

## 📦 MCP 工具映射

| MCP 工具 | Python API | 用途 |
|---------|-----------|------|
| excel_query | execute_advanced_sql_query | SQL 查询 |
| excel_update_query | execute_advanced_update_query | UPDATE 语句 |
| excel_insert_query | execute_advanced_insert_query | INSERT 语句 |
| excel_delete_query | execute_advanced_delete_query | DELETE 语句 |

## 🎯 测试检查清单

开发新功能时，确保覆盖以下场景：

- [ ] SELECT 子句
- [ ] WHERE 条件
- [ ] ORDER BY 排序
- [ ] GROUP BY 聚合
- [ ] UPDATE 赋值
- [ ] 数学表达式嵌套
- [ ] CASE WHEN 表达式
- [ ] 边界情况（NULL、空值、负数）
- [ ] 窗口函数（如适用）
- [ ] 跨文件操作（如适用）

## 🚀 部署前检查

```bash
# 1. 完整回归测试
python -m pytest tests/ --ignore=tests/test_pivot_table.py

# 2. 清理缓存
find src -type d -name "__pycache__" -exec rm -rf {} +

# 3. 提交代码
git add src/
git commit -m "feat: add new feature"

# 4. 用户重启 Claude Code 后生效
```

## 📝 代码规范

- 函数添加：在 `_COMPLEX_EXPR_TYPES` 中注册新类型
- 分发表定义：使用 `tuple` 而非 `frozenset`
- 向量化优先：使用 `pd.Series` 向量操作
- 异常回退：提供逐行处理的 fallback

## 🎯 SQL 核心原则

**只支持 SQL 标准支持的功能，SQL 都不支持的就不用支持了**

### 为什么？

1. **与主流数据库保持一致** - MySQL, PostgreSQL, SQL Server, Oracle 都遵循 SQL 标准
2. **避免自创特殊行为** - 不实现 SQL 标准之外的特殊逻辑
3. **降低维护成本** - 符合标准的行为更容易维护和文档化
4. **用户可预期** - 用户已有的SQL知识可以直接应用

### 应用规范

**功能开发**:
- ✅ **支持**: 符合 SQL 标准的功能（如窗口函数、JOIN、聚合）
- ❌ **不支持**: SQL 标准明确不支持的功能（如 WHERE 引用窗口函数别名）
- 💡 **错误提示**: 对不支持的功能返回友好错误提示，说明原因和解决方案

**测试用例设计**:
- ✅ **测试用例必须符合 SQL 标准**: 不测试 SQL 不支持的写法
- ❌ **不测试边缘 bug**: SQL 标准限制不是 bug
- ✅ **可以测试错误提示质量**: 标记为 `expected_error: True`

### 典型 SQL 标准限制

**WHERE 子句限制**:
- ❌ **不能引用**: 窗口函数别名（如 `RANK() OVER(...)`）、SELECT 子句别名
- ✅ **可以引用**: 原始列名、字面量

**原因**: SQL 执行顺序为 `FROM → WHERE → 窗口函数 → SELECT`

**SQL 执行顺序**:
```
1. FROM     → 确定数据源
2. JOIN     → 表关联
3. WHERE    → 筛选行
4. GROUP BY → 分组
5. HAVING   → 分组筛选
6. 窗口函数 → 计算窗口函数
7. SELECT   → 选择列
8. ORDER BY → 排序
9. LIMIT    → 限制结果
```

### 例子

**✅ 正确的测试**:
```sql
-- 符合 SQL 标准
SELECT * FROM (SELECT *, RANK() as 排名 FROM 表) t WHERE 排名 <= 3
```

**❌ 错误的测试（不要这样设计）**:
```sql
-- SQL 标准不支持，不要期望通过
SELECT *, RANK() as 排名 FROM 表 WHERE 排名 <= 3
```

---

**记录日期**: 2026-04-13
**Why**: 测试场景4 `WHERE 窗口函数别名` 失败，用户强调"只支持sql支持的，sql都不支持的就不用支持了"

## 🔄 版本更新

更新 `__init__.py` 中的版本号：

```python
__version__ = "1.9.3"  # 递增版本号
```

---

**最后更新**: 2026-04-14
**维护者**: tangjian
