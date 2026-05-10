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

---

## 🛡️ 不变量驱动开发（IDD）体系

### 定位正确性

本项目是「SQL-over-Excel」引擎，正确性天然二值——同一条 SQL 在 ExcelMCP 和 SQLite 上的结果必须一致（浮点容差 0.01）。
因此可以直接进入 IDD 循环，无需投影层。

### 真值来源

| 真值 | 位置 | 查阅方式 |
|------|------|----------|
| SQL 标准行为 | SQLite 3.x | `calibrator` 导入 Excel → 跑同 SQL → 对比结果 |
| 文件完整性 | Excel 文件本身 | 操作前后用 openpyxl 读取验证 |
| API 契约 | `advanced_sql_query.py` 公共 API | 返回值结构 `{success, data, message}` |

### 判据执行方式

```bash
# 全量不变量检查（CI 集成用）
python -m pytest tests/invariants/ -v --tb=short

# 快速烟雾测试（开发中用）
python -m pytest tests/invariants/ -v -k "smoke"

# SQLite 交叉校验
python -m pytest tests/test_p02_sqlite_cross_validation.py -v
```

执行时间：全量 < 60s，烟雾 < 10s。输出格式：pytest 标准输出，二值（PASS/FAIL）。

### 不变量清单

#### L1 外部真值（SQL 标准 + 文件系统）

| 编号 | 名称 | 检查内容 | 发现来源 |
|------|------|----------|----------|
| INV-1 | 结果结构一致性 | `result` 必须包含 `success` (bool)、`data` (list)、`message` (str) 三个键 | 初始设计 |
| INV-2 | SQL-SQLite 结果对齐 | 同一 SQL 在 ExcelMCP 和 SQLite 上的结果集一致（浮点容差 0.01，列顺序无关） | 初始设计 |
| INV-3 | 文件完整性守恒 | SELECT 不修改文件；UPDATE/INSERT/DELETE 只修改目标 sheet，其他 sheet 和文件属性不变 | 初始设计 |
| INV-4 | 行数守恒 | `SELECT COUNT(*)` 返回的行数 = 实际数据行数（不含表头） | 初始设计 |

#### L2 架构原则

| 编号 | 名称 | 检查内容 | 发现来源 |
|------|------|----------|----------|
| INV-5 | 失败安全 | `success=False` 时 `data` 为空列表，`message` 非空且不含堆栈信息 | 初始设计 |
| INV-6 | 错误可分类 | 所有错误消息能被 `ToolCallTracker.classify_error()` 归入已知类别 | 初始设计 |
| INV-7 | 幂等读取 | 同一 SELECT 连续执行两次，结果完全一致 | 初始设计 |
| INV-8 | LIMIT 约束 | `SELECT ... LIMIT N` 返回行数 ≤ N | 初始设计 |
| INV-9 | 聚合语义正确 | `COUNT(*)` ≥ `COUNT(col)`（NULL 不计）；`SUM` 忽略 NULL；空表 `COUNT` → 0，`SUM/AVG/MIN/MAX` → NULL | 初始设计 |

#### L3 具体不变量（对抗发现，持续生长）

| 编号 | 名称 | 检查内容 | 发现来源 |
|------|------|----------|----------|
| INV-10 | 窗口函数唯一性 | `ROW_NUMBER()` 在同一 PARTITION 内严格递增且无重复 | 初始设计 |
| INV-11 | 排名标准合规 | `RANK()` 在并列时跳号（1,1,3），`DENSE_RANK()` 不跳号（1,1,2） | 初始设计 |
| INV-12 | 空表安全 | 空表上执行任意 SELECT 返回空数据行但不报错；聚合返回 NULL/0 | 初始设计 |
| INV-13 | 特殊字符安全 | 列名/值含中文、emoji、单引号、反斜杠时查询不崩溃 | 初始设计 |
| INV-14 | 除零安全 | `1/0` 返回 NULL 而非 inf 或崩溃 | 初始设计 |
| INV-15 | LIKE 安全 | LIKE 模式含正则元字符（`[`、`]`）时不崩溃；超长模式被拒绝 | 初始设计 |
| INV-16 | UPDATE 后读回验证 | UPDATE 后 SELECT 读回，SET 表达式生效、非目标列/行不变、幂等 | Round 1 对抗 |
| INV-17 | INSERT 行数守恒 | INSERT N 行后 COUNT(*) 增加 N；读回验证列值；失败不改变文件 | Round 1 对抗 |
| INV-18 | DELETE 行数守恒 | DELETE 后 COUNT(*) 减少 affected_rows；被删行不再出现；失败安全 | Round 1 对抗 |
| INV-20 | 公式列守恒 | 含公式的文件 UPDATE 后，非目标列公式仍存在或被正确计算 | Round 1 对抗 |
| INV-22 | affected_rows 精确 | `affected_rows` == 实际变更行数（openpyxl 验证） | Round 1 对抗 |
| INV-23 | 无匹配写操作安全 | UPDATE/DELETE WHERE 无匹配 → affected_rows=0，文件完全不变 | Round 1 对抗 |
| INV-24 | NULL 写入语义 | 数值列写 NULL 后读回为 NULL/空；不影响其他列/行/行数 | Round 1 对抗 |
| INV-25 | DISTINCT 语义正确性 | DISTINCT 消除重复行；COUNT(DISTINCT) 排除 NULL（SQL 标准） | Round 3 对抗 |
| INV-26 | HAVING 子句正确性 | HAVING 在 GROUP BY 后正确过滤；HAVING vs WHERE 区别 | Round 3 对抗 |
| INV-27 | NULL 比较正确性 | IS NULL/IS NOT NULL 正确识别空值；UPDATE 后 IS NULL 可检测 | Round 3 对抗 |
| INV-28 | 子查询正确性 | IN(SELECT...)、NOT IN(SELECT...)、标量子查询正确执行 | Round 3 对抗 |
| INV-29 | OFFSET 边界正确性 | OFFSET 超总行数返回空；OFFSET 0 等同不使用 | Round 3 对抗 |
| INV-30 | NOT IN/NOT LIKE 语义 | NOT IN 排除指定值；NOT LIKE 排除匹配行 | Round 3 对抗 |
| INV-31 | 双行表头写操作 | 双行表头表的 UPDATE/INSERT/DELETE 正确工作 | Round 3 对抗 |
| INV-32 | _ROW_NUMBER_ 写操作 | UPDATE/DELETE WHERE _ROW_NUMBER_ 精确定位行 | Round 3 对抗 |
```
L1: SQL 标准（SQLite 参考实现）
  └─ L2: 结果可复现、失败安全、API 契约稳定
      └─ L3: INV-1~4 结果结构/对齐/完整性/行数
      └─ L3: INV-5~6 错误处理质量
      └─ L3: INV-7~9 读取幂等/LIMIT/聚合语义
      └─ L3: INV-10~15 窗口函数/空表/特殊字符/除零/LIKE
      └─ L3: INV-19 写操作 SQLite 对齐
      └─ L3: INV-20 公式列守恒
      └─ L3: INV-21 跨文件 JOIN 真值
      └─ L3: INV-22~24 affected_rows精确/无匹配安全/NULL写入语义
      └─ L3: INV-25~28 DISTINCT/HAVING/NULL比较/子查询
      └─ L3: INV-29~32 OFFSET边界/NOT操作/双行表头写操作/_ROW_NUMBER_写操作

### 对抗策略

当前已实施的对抗维度：

| 数据篡改 | 空表、单行表、全 NULL 列、超长字符串(5000char)、inf/nan | INV-4, INV-9, INV-12, INV-13, INV-14 |
| 类型混淆 | 数字列混入文本、空字符串 vs NULL | INV-9, INV-13 |
| 编码破坏 | 中文/日文/韩文/emoji 列名和值、单引号、反斜杠 | INV-13 |
| 边界值 | LIMIT 0、OFFSET 超范围、负数、极大值(1e15)、极小值(0.000000001) | INV-8, INV-12 |
| SQL 注入 | LIKE 模式含正则元字符、超长 LIKE 模式 | INV-15 |
| 窗口函数 | 空/单行表上跑所有窗口函数、无 ORDER BY 的窗口 | INV-10, INV-11, INV-12 |
| 交叉验证 | 同一 SQL 跑 ExcelMCP + SQLite 对比 | INV-2 |
| **写操作语义** | **UPDATE 读回验证、INSERT/DELETE 行数守恒、NULL 写入** | **INV-16, INV-17, INV-18, INV-24** |
| **公式保留** | **含公式 Excel 上 UPDATE 后公式列完整性** | **INV-20** |
| **affected_rows 精确性** | **WHERE 匹配/不匹配时 affected_rows 准确性** | **INV-22, INV-23** |

待升级维度（连续 3 轮零发现后启用）：
- **并发竞争**：多线程同时 UPDATE 同一文件
- **大文件压力**：10 万行 × 1000 列的聚合/JOIN 性能

### 收敛状态

- **当前轮次**：Round 3（SQL 功能边界）
- **连续零发现**：0 轮
- **不变量总数**：32 条（L1: 4, L2: 5, L3: 23）
- **测试总数**：154 passed, 3 skipped
- **收敛评级**：C（三轮对抗完成，连续 0 新 bug 发现）
- **下次对抗触发条件**：新功能合并后、或手动启动

```
- ❌ 不要在 UPDATE 后假设值已正确写入 → 用 SELECT 读回验证（INV-16）
- ❌ 不要在 INSERT/DELETE 后假设行数正确 → 用 COUNT(*) 验证（INV-17, INV-18）
- ❌ 不要在 UPDATE/INSERT/DELETE 后假设文件未变更 → 用 openpyxl 重新读取验证（INV-3）
- ❌ 不要假设 SELECT 返回的行数等于 Excel 可见行数 → 用 COUNT(*) 验证（INV-4）
- ❌ 不要手动构造 expected 结果 → 用 SQLite 交叉校验（INV-2）
- ❌ 不要在测试中容忍 inf/nan → 应返回 NULL（INV-14）
- ❌ 不要跳过空表测试 → 空表是最常见的边界 case（INV-12）
- ❌ 不要假设 affected_rows 准确 → 用 openpyxl 实际验证（INV-22）
- ❌ 不要假设无匹配的写操作是安全的 → 验证文件未变（INV-23）
```

### 不变量生长记录

#### Round 0 — 2026-05-10（初始设计）

**对抗策略**：基于现有测试覆盖分析，从 test_p02_sqlite_cross_validation.py、test_r42_edge_cases.py、test_r44_edge_cases.py 中提炼。

**发现**：
- 新 bug：0 个
- 新不变量：15 条（全部初始设计）
- 假阴性升级：0 条

**不变集状态**：77 passed, 3 skipped

#### Round 1 — 2026-05-10（写操作对抗）

**对抗策略**：Oracle + Explore 并行诊断，针对写入路径的系统性盲区设计对抗。

**诊断发现**（Oracle 诊断报告）：
- FN-1: UPDATE 静默写错值（高危假阴性）→ 设计 INV-16 消除
- FN-2: INSERT 写到错误行（高危假阴性）→ 设计 INV-17 消除
- FN-3: DELETE 删错行（高危假阴性）→ 设计 INV-18 消除
- FN-4: UPDATE 覆盖公式列（高危假阴性）→ 设计 INV-20 消除
- FN-8: affected_rows 多报/少报（中危假阴性）→ 设计 INV-22 消除
- G1~G5: 写操作无 SQLite 交叉校验（高严重度间隙）→ 留待 Round 2

**对抗结果**：
- 新 bug：0 个（全部 114 测试通过）
- 新不变量：7 条（INV-16/17/18/20/22/23/24）
- 假阴性消除：6 个高危假阴性被 INV-16~24 覆盖
- 剩余假阴性：G1~G5（写操作 SQLite 对齐）、跨文件 JOIN 真值

**不变集状态**：114 passed, 3 skipped

**技术发现**：pandas FutureWarning 提示 L9367 `df.at[idx, col_name] = new_val` 存在 dtype 不兼容风险，`except (ValueError, TypeError, OverflowError): pass` 静默跳过类型转换错误。非当前 bug 但属技术债。

#### Round 2 — 2026-05-11（SQLite 对齐 + 跨文件 JOIN）

**对抗策略**：针对 G1（写操作 SQLite 对齐）和 G2（跨文件 JOIN 真值）两个高优先级间隙设计对抗。

**诊断发现**：
- calibrator `cmd_query` 不调用 `conn.commit()`，DML 不持久化 → 测试中直接用 sqlite3 连接执行 DML
- 跨文件 JOIN 语法 `表名@'path'`，calibrator 支持多文件导入同一 db → 可直接对比

**对抗结果**：
- 新 bug：0 个（全部 125 测试通过）
- 新不变量：2 条（INV-19/21）
- 假阴性消除：G1（写操作 SQLite 对齐）、G2（跨文件 JOIN 真值）
- 剩余假阴性：并发竞争、大文件压力、跨文件 UPDATE 原子性

**不变集状态**：125 passed, 3 skipped

#### Round 3 — 2026-05-11（SQL 功能边界）

**对抗策略**：扫描 SQL 功能覆盖间隙，针对 DISTINCT/HAVING/NULL比较/子查询/OFFSET/双行表头写操作/_ROW_NUMBER_ 写操作设计对抗。

**对抗结果**：
- 新 bug：0 个（全部 154 测试通过）
- 新不变量：8 条（INV-25~32）
- 假阴性消除：SQL 功能边界覆盖间隙全部填补
- 剩余假阴性：并发竞争、大文件压力、跨文件 UPDATE 原子性

**不变集状态**：154 passed, 3 skipped

---

**最后更新**: 2026-05-11
**维护者**: tangjian
