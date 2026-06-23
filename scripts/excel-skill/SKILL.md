---
name: excel
description: 游戏开发 Excel 配置表管理 — SQL 查询、批量操作、结构管理、格式化。通过 CLI 可执行文件调用，加载无状态，支持 self-update。
keywords: ["excel", "配置表", "SQL", "openpyxl", "游戏开发", "xlsx", "查询", "批量操作"]
version: "2.0.0"
---

# Excel 配置表 Skill

游戏开发专用 Excel 配置表管理。SQL-over-Excel 引擎，26 个工具，支持高级 SQL 查询、批量操作、跨文件 JOIN。

## 调用方式

所有操作通过 CLI 可执行文件调用，输出 JSON。

**优先使用打包二进制**（无需 Python 环境）：

```bash
# 打包后的独立可执行文件（ZIP 包内 bin/excel-cli.exe 或 bin/excel-cli）
excel-cli <command> [options]
```

**回退到 Python 源码**（开发环境）：

```bash
python scripts/excel-cli.py <command> [options]
```

**输出格式**：`{success: bool, data: Any, message: str, meta: dict}`，退出码 0=成功 1=失败。

## 核心原则：SQL 优先

**优先使用 `query`** — 所有数据查询/分析任务
- 复杂条件筛选 → WHERE, LIKE, IN, BETWEEN, 子查询
- 聚合统计 → COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 多表关联 → 5 种 JOIN，支持跨文件
- 窗口函数 → ROW_NUMBER, RANK, DENSE_RANK
- 字符串函数 → UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING

## 工具选择决策树

```
═══ 读数据 ═══
首选：所有数据查询/分析 → query（SQL 引擎，批量分析首选）
│
├─ 已知精确坐标（如 A1:C10）────────────→ get-range
├─ 快速了解表结构（列名+类型+样本值）───→ describe-table
├─ 只需表头信息（中文+英文）───────────→ get-headers（更轻量）
├─ 定位文本位置───────────────────────→ search（返回 row/column）
├─ 跨文件搜索─────────────────────────→ search-directory
└─ 查找末行（追加数据前必用）──────────→ find-last-row

═══ 写数据（选对工具！） ═══
┌─ 批量修改多行？（改 10 行以上/按条件）──→ update-query（SQL UPDATE）
│  需要计算表达式 → SET 血量=血量*2
│  需要预览 → --dry-run
│
├─ 精确坐标写入（知道具体 A1:C10）───────→ update-range
│  默认覆盖！--insert-mode 才是追加
│  安全追加？→ 先 find-last-row → 再 update-range --insert-mode
│
├─ 按 ID 改单行────────────────────────→ upsert-row
│  存在则更新，不存在则插入，幂等安全
│
├─ 批量插入新行────────────────────────→ insert-query（SQL INSERT）
└─ 删行？
    按条件删？→ delete-query（SQL DELETE，必须 WHERE）
    按行号删？→ structure delete_rows

═══ 高级脚本 ═══
 循环/复杂逻辑/重复操作？→ run-python（直接执行 Python 代码）

═══ 结构操作 ═══
文件？    → create-file
工作表？  → list-sheets / create-sheet / delete-sheet / rename-sheet / copy-sheet
行/列？   → structure（insert/delete rows+columns）
改列名？  → rename-column
行高/列宽？→ set-layout

═══ 格式化 ═══
样式/合并/边框？ → format-cells（字体+合并+边框+预设样式）

═══ 对比 & 备份 ═══
按 ID 对比两表？ → compare-sheets
备份恢复？       → backup（create/list/restore）
```

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
