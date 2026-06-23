---
name: excel
description: 游戏开发 Excel 配置表管理 — SQL 查询、批量操作、结构管理、格式化。通过 CLI 可执行文件调用，加载无状态，支持 self-update。
keywords: ["excel", "xlsx", "配置表", "SQL", "openpyxl", "游戏开发", "怪物表", "物品表", "技能表", "查询", "批量操作", "公式", "工作表", "表格"]
version: "2.0.0"
tags:
  - excel
  - sql
  - gamedev
---

# Excel 配置表 Skill

游戏开发专用 Excel 配置表管理。SQL-over-Excel 引擎，26 个工具，支持高级 SQL 查询、批量操作、跨文件 JOIN。

## 调用方式

**入口**：`bin/excel-cli.py`（本 skill 目录内）

```bash
python bin/excel-cli.py <command> [options]
# 首次运行自动：uv venv → uv pip install git+... → 执行命令
```
所有操作通过 CLI 命令调用，输出 JSON。skill 自动管理 venv + 安装。

**自更新**：

```bash
# 检查最新版本
excel-cli self-update --check

# 执行更新（自动 pip install git+...）
excel-cli self-update
```

**输出格式**：`{success: bool, data: Any, message: str, meta: dict}`，退出码 0=成功 1=失败。

## 场景决策（选对工具）

### 读数据
| 你的需求 | 用哪个 | 说明 |
|---------|--------|------|
| 按条件筛选/聚合/多表关联/跨文件 | `query --sql "SELECT …"` | SQL 引擎，能处理复杂条件 |
| 知道精确坐标，如 A1:C10 | `get-range --range "A1:C10"` | 直接读单元格范围 |
| 快速了解表结构（列名+类型+样本值） | `describe-table` | 看有哪些列、什么类型 |
| 只需要表头（中文+英文列名） | `get-headers` | 比 describe-table 更轻量 |
| 搜索某个值在哪 | `search --pattern "xxx"` | 返回行列坐标 |
| 跨多个 Excel 文件搜索 | `search-directory` | 指定目录搜全部文件 |
| 找最后一行的行号（追加数据前用） | `find-last-row --column A` | 定位空行位置 |

### 写数据
| 你的需求 | 用哪个 | 注意 |
|---------|--------|------|
| 按条件批量改（如所有等级>5的怪加血） | `update-query --sql "UPDATE … SET … WHERE …"` | 支持表达式，可用 `--dry-run` 预览 |
| 精确位置写入（知道 A1:C10 填什么） | `update-range --range "A1:C10" --data '[[…]]'` | **默认覆盖！追加要加 `--insert-mode`** |
| 按 ID 改单行（存在更新，不存在插入） | `upsert-row --key-column ID --key-value N --updates '{…}'` | 幂等安全，推荐用于单行 |
| 批量插入多行 | `insert-query --sql "INSERT INTO … VALUES …"` | SQL INSERT 语法 |
| 按条件删行 | `delete-query --sql "DELETE FROM … WHERE …"` | **必须带 WHERE** |
| 按行号删行 | `structure --operation delete_rows --index N --count M` | 直接指定行位置 |

### 结构 / 格式 / 其他
| 需求 | 用哪个 |
|------|--------|
| 新建文件 | `create-file` |
| 增/删/重命名/复制工作表 | `create-sheet / delete-sheet / rename-sheet / copy-sheet` |
| 插入/删除行或列 | `structure --operation insert_rows / delete_rows / insert_columns / delete_columns` |
| 改列名 | `rename-column` |
| 设置行高/列宽 | `set-layout --operation set_row_height / set_column_width` |
| 格式化单元格（字体/边框/合并/背景色） | `format-cells --range "…" --formatting '{…}'` |
| 写公式 | `set-formula --cell A1 --formula "=SUM(B1:B10)"` |
| 复杂逻辑/循环操作 | `run-python --code "…"`（Python 脚本，可用 query/update 变量） |
| 两表按 ID 对比差异 | `compare-sheets` |
| 备份/恢复 | `backup --operation create / list / restore` |

## 防错自查清单

| # | 自查问题 | 常见错误 | 正确做法 |
|---|---------|---------|---------|
| 1 | 追加还是覆盖？ | update-range 追加忘了 --insert-mode | 追加→`--insert-mode`+先 `find-last-row`；覆盖→默认 |
| 2 | 一行还是批量？ | 改单行用 update-query | 单行→`upsert-row`；批量/条件→`update-query` |
| 3 | 范围含工作表名？ | 多表文件 cell_range="A1:C10" 报错 | 用 `"Sheet名!A1:C10"` |
| 4 | SQL 类型对吗？ | SELECT 传给 update-query | 查→query / 改→update-query / 增→insert-query / 删→delete-query |
| 5 | 写入后验证了？ | 写完不验证 | 写入→query 验证→有备份可恢复 |

## 双行表头

当 Excel 有双行表头（第 1 行中文 + 第 2 行英文）时：
- **SQL 工具**（query/update-query/insert-query/delete-query）：中英文名都能用
- **describe-table 返回的列名**：第 2 行英文名
- **upsert-row 的 --key-column**：中英文都能用
- **建议**：直接用 describe-table 返回的英文名

## 完整命令参考

### 查询类
```bash
query --file F --sql "SELECT * FROM 表 WHERE 条件" [--no-headers] [--format table|json]
list-sheets --file F
get-headers --file F [--sheet S] [--header-row 1] [--max-columns N]
get-range --file F --range "A1:C10" [--sheet S] [--formatting]
describe-table --file F [--sheet S]
search --file F --pattern "文本" [--sheet S] [--case-sensitive] [--whole-word] [--regex]
search-directory --dir D --pattern "文本" [--extensions ".xlsx,.xls"]
find-last-row --file F --sheet S --column A
compare-sheets --file1 F1 --sheet1 S1 --file2 F2 --sheet2 S2 [--id-column 1]
```

### 写入类
```bash
update-query --file F --sql "UPDATE 表 SET 列=值 WHERE 条件" [--dry-run]
insert-query --file F --sql "INSERT INTO 表 (列) VALUES (值)" [--dry-run]
delete-query --file F --sql "DELETE FROM 表 WHERE 条件" [--dry-run]
update-range --file F --range "A1:B2" --data '[["a","b"],["c","d"]]' [--insert-mode] [--sheet S]
upsert-row --file F --sheet S --key-column ID --key-value 3 --updates '{"血量":900}'
set-formula --file F --sheet S --cell A1 --formula "=SUM(B1:B10)"
run-python --file F --code "query('SELECT * FROM 表')" [--sheet S] [--timeout 30]
```

### 结构操作
```bash
create-file --file F [--sheets '["表1","表2"]']
create-sheet --file F --name 新表 [--index 0]
delete-sheet --file F --name 表名
rename-sheet --file F --old-name 旧 --new-name 新
copy-sheet --file F --source 源表 [--new-name 副本]
structure --file F --sheet S --operation insert_rows --index 2 --count 3
rename-column --file F --sheet S --old-header 旧名 --new-header 新名
```

### 格式化
```bash
format-cells --file F --sheet S --range "A1:C10" [--formatting '{"bold":true}'] [--preset header]
set-layout --file F --sheet S --operation set_row_height --index 1 --value 30
```

### 备份
```bash
backup --file F --operation create
backup --file F --operation list
backup --file F --operation restore --backup-id 20260623_120000
```

### 自更新
```bash
self-update    # 从 GitHub 拉取最新版本
```

## SQL 已支持功能

基础: SELECT, DISTINCT, 别名(AS), 数学表达式(+-*/%)
条件: WHERE, LIKE, IN, BETWEEN, 子查询
聚合: COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
多表: INNER/LEFT/RIGHT/FULL JOIN, 跨文件 JOIN
窗口: ROW_NUMBER, RANK, DENSE_RANK
高级: CASE WHEN, CTE(WITH), EXISTS, UNION, UNION ALL
字符串: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING
排序: ORDER BY, LIMIT, OFFSET

SQL 真值来源：SQLite 3.x，calibrator 交叉校验保证对齐。

## 边界

- **不负责**：数据库管理（MySQL/PostgreSQL）、CSV 处理、Excel 透视表创建
- **只处理**：.xlsx / .xls 文件的配置表读写查询
- **SQL 限制**：不支持 WHERE 引用窗口函数别名、SELECT 别名在 WHERE 中
