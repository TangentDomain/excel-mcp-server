---
name: excel
description: 游戏开发 Excel 配置表管理 — SQL 查询、批量操作、结构管理、格式化。
keywords: ["excel", "xlsx", "配置表", "SQL", "openpyxl", "游戏开发", "怪物表", "物品表", "技能表", "查询", "批量操作", "公式", "工作表", "表格"]
version: "3.2.0"
tags:
  - excel
  - sql
  - gamedev
---

# Excel 配置表 Skill

游戏开发专用 Excel 配置表管理。SQL-over-Excel 引擎，26 个工具，支持高级 SQL 查询、批量操作、跨文件 JOIN。

## 快速调用

### 方式一：Skill 接入（首推）

**安装**（二选一）：

```bash
# 方式 A：OMP harness 用户（自动注册 skill）
# skill 文件部署到 ~/.omp/agent/skills/excel/，harness 自动发现

# 方式 B：手动安装
git clone https://github.com/TangentDomain/excel-mcp-server.git
cp -r excel-mcp-server/.omp/skills/excel ~/.omp/agent/skills/excel
```

**使用**：

```bash
# 首次运行自动安装依赖（uv venv + pip install 从 GitHub）
python ~/.omp/agent/skills/excel/bin/excel-cli.py query --file data.xlsx --sql "SELECT * FROM Sheet1"

# Windows venv 安装后可直接用 exe（跳过自举，更快）
~/.omp/agent/skills/excel/.venv/Scripts/excel-cli.exe <command> [options]

# 检查更新
excel-cli self-update --check
excel-cli self-update          # 更新到最新版
```

Skill 入口：`~/.omp/agent/skills/excel/bin/excel-cli.py`，首次运行自动创建 venv + 从 GitHub 安装，后续通过 `self-update` 更新。

### 方式二：源码直接调用（开发环境）

本项目即为 excel-mcp-server 源码，可以用 `uv run` 直接调用：

```bash
uv run python -m excel_mcp_server_fastmcp.cli <command> [options]
```

**输出格式**（所有命令统一）：JSON `{success: bool, data: Any, message: str, meta: dict}`

**所有命令共用参数**：`--file <路径>` 是每个命令的必填参数（除 self-update）。路径用绝对路径。

## 核心决策树：选对命令

```
读数据？
├─ 筛选/聚合/GROUP BY/JOIN/窗口函数/子查询 → query
├─ 知道精确坐标 A1:C10                     → get-range
├─ 不确定有哪些列                          → describe-table（列名+类型+样本）
├─ 只要表头                                → get-headers
├─ 搜某个值在哪张表                        → search / search-directory
└─ 追加数据前找空行                        → find-last-row

写数据？
├─ 按条件批量改（WHERE 筛选）              → update-query
├─ 精确坐标写入                            → update-range（默认覆盖！追加加 --insert-mode）
├─ 按 ID 改/插单行（幂等）                  → upsert-row（推荐单行操作）
├─ SQL INSERT 插入                         → insert-query
├─ 按条件删行                              → delete-query（必须 WHERE）
└─ 按行号删/插入行列                       → structure

其他？
├─ 工作表管理 → create-sheet / delete-sheet / rename-sheet / copy-sheet
├─ 格式化     → format-cells（字体/边框/合并/背景色）
├─ 行高列宽   → set-layout
├─ 公式       → set-formula
├─ 复杂逻辑   → run-python（可用 query/update/insert/delete 函数 + ExcelOperations）
├─ 两表对比   → compare-sheets
└─ 备份恢复   → backup
```

## 完整命令签名

> `*` = 必填，`[]` = 可选

### 查询类

| 命令 | 签名 | 说明 |
|------|------|------|
| `query` | `--file*` 路径 `--sql*` "..." [`--format` 格式] [`--no-headers`] | SQL 查询（优先使用） |
| `describe-table` | `--file*` 路径 [`--sheet` S] | 查看表结构（列名+类型+样本值） |
| `get-headers` | `--file*` 路径 [`--header-row` N] [`--max-columns` N] [`--sheet` S] | 获取表头信息（中文+英文） |
| `list-sheets` | `--file*` 路径 | 列出所有工作表名称 |
| `get-range` | `--file*` 路径 `--range*` 范围 [`--formatting` JSON] [`--sheet` S] | 获取指定单元格范围的数据 |
| `search` | `--file*` 路径 `--pattern*` "关键词" [`--case-sensitive`] [`--range` 范围] [`--regex`] [`--sheet` S] [`--whole-word`] | 在 Excel 中搜索单元格 |
| `search-directory` | `--dir*` 路径 `--pattern*` "关键词" [`--extensions` 列表] [`--max-files` 路径] [`--recursive`] | 在目录下搜索 Excel 文件 |
| `find-last-row` | `--file*` 路径 `--sheet*` S [`--column` 列] | 查找最后一行 |
| `compare-sheets` | `--file1*` 路径 `--file2*` 路径 `--sheet1*` S `--sheet2*` S [`--header-row` N] [`--id-column` 列] | 按 ID 列比较两个工作表差异 |

### 写入类

| 命令 | 签名 | 说明 |
|------|------|------|
| `update-query` | `--file*` 路径 `--sql*` "..." [`--dry-run`] | SQL UPDATE 批量修改 |
| `insert-query` | `--file*` 路径 `--sql*` "..." [`--dry-run`] | SQL INSERT 插入数据 |
| `delete-query` | `--file*` 路径 `--sql*` "..." [`--dry-run`] | SQL DELETE 删除数据 |
| `update-range` | `--file*` 路径 `--data*` '[[...]]' `--range*` 范围 [`--insert-mode`] [`--no-preserve-formulas` "公式"] [`--preserve-formulas` "公式"] [`--sheet` S] | 精确坐标写入数据 |
| `upsert-row` | `--file*` 路径 `--key-column*` 列 `--key-value*` 值 `--sheet*` S `--updates*` '{...}' [`--header-row` N] | 按 key_column+key_value 插入或更新行 |
| `set-formula` | `--file*` 路径 `--cell*` A1 `--formula*` "公式" `--sheet*` S | 写入 Excel 公式 |

### 结构类

| 命令 | 签名 | 说明 |
|------|------|------|
| `create-file` | `--file*` 路径 [`--sheets` JSON数组] | 创建新 Excel 文件（`--sheets` 接受 JSON 数组，如 `'["表1","表2"]'`） |
| `create-sheet` | `--file*` 路径 `--name*` "名称" [`--index` N] | 创建新工作表 |
| `delete-sheet` | `--file*` 路径 `--name*` "名称" | 删除工作表 |
| `rename-sheet` | `--file*` 路径 `--new-name*` "名称" `--old-name*` "名称" | 重命名工作表 |
| `copy-sheet` | `--file*` 路径 `--source*` 源表 [`--index` N] [`--new-name` "名称"] | 复制工作表 |
| `structure` | `--file*` 路径 `--index*` N `--operation*` 操作 `--sheet*` S [`--count` N] | 插入或删除行和列 |
| `rename-column` | `--file*` 路径 `--new-header*` "列名" `--old-header*` "列名" `--sheet*` S [`--header-row` N] | 修改列名（表头） |

### 格式化类

| 命令 | 签名 | 说明 |
|------|------|------|
| `format-cells` | `--file*` 路径 `--range*` 范围 `--sheet*` S [`--formatting` JSON] [`--preset` 预设] | 设置单元格样式 |
| `set-layout` | `--file*` 路径 `--index*` N `--operation*` 操作 `--sheet*` S `--value*` 值 [`--count` N] | 设置行高或列宽 |

### 运行类

| 命令 | 签名 | 说明 |
|------|------|------|
| `run-python` | `--file*` 路径 `--code*` "..." [`--sheet` S] [`--timeout` N] | 直接执行 Python 代码操作 Excel |

### 运维类

| 命令 | 签名 | 说明 |
|------|------|------|
| `backup` | `--file*` 路径 `--operation*` 操作 [`--backup-dir` 路径] [`--backup-id` "ID"] | 备份与恢复 |
| `self-update` |  [`--check`] | 检查/更新 CLI 到最新版本 |

## 使用方法论（避免常见错误）

### 第一步：永远先 describe-table

**不要假设你知道列名。** 每次操作新文件前先跑：

```bash
# 先看有哪些 Sheet
uv run python -m excel_mcp_server_fastmcp.cli list-sheets --file F
# 再看目标 Sheet 的结构（自动检测双行表头）
uv run python -m excel_mcp_server_fastmcp.cli describe-table --file F [--sheet Sheet名]
```

返回的 `columns[].name` 就是 SQL 可用的列名。**用 describe-table 返回的列名写 SQL**，不要猜。

### 第二步：Sheet 名 ≠ 文件名

`Props.xlsx` 的 Sheet 可能叫 `PropList`、`PropName`。SQL 中 `FROM` 后面跟的是 **Sheet 名**，不是文件名：

```bash
# ❌ 错误：用文件名当表名
uv run ... query --file "Props.xlsx" --sql "SELECT * FROM Props"

# ✅ 正确：先查 Sheet 名，再用 Sheet 名
uv run ... query --file "Props.xlsx" --sql "SELECT * FROM PropList"
```

### 第三步：数值列可能是文本存储

游戏配置表中数值常以文本存储（如 `Hp = "2000"` 而非 `2000`）。排序/比较时用 `CAST`：

```bash
# ❌ 文本排序："100" < "20" < "3"（字典序）
uv run ... query --file F --sql "SELECT ID, Hp FROM Monster ORDER BY Hp DESC LIMIT 5"

# ✅ 数值排序：CAST 转换后正确排序
uv run ... query --file F --sql "SELECT ID, CAST(Hp AS INT) AS hp FROM Monster ORDER BY hp DESC LIMIT 5"
```

### 第四步：多表文件指定 --sheet

一个 xlsx 有多个 Sheet 时，用 `--sheet` 指定目标：

```bash
uv run ... query --file "Monster.xlsx" --sheet Monster --sql "SELECT * FROM Monster LIMIT 5"
```

### 第五步：写操作安全链

```bash
# 1. 备份
uv run ... backup --file F --operation create
# 2. dry-run 预览
uv run ... update-query --file F --sql "UPDATE 表 SET 列=值 WHERE 条件" --dry-run
# 3. 确认无误后执行
uv run ... update-query --file F --sql "UPDATE 表 SET 列=值 WHERE 条件"
# 4. 验证结果
uv run ... query --file F --sql "SELECT * FROM 表 WHERE 条件"
```

### 已知限制

- 空字段名的列会显示为 `Unnamed__N` → 用列序号或 describe-table 确认
- 数值以文本存储时需 CAST → 见第三步
- Excel 空字符串往返变为 NULL → xlsx 格式固有限制
- `run-python` 沙箱：`open`/`os`/`subprocess` 等文件/进程操作被禁；安全内置含 abs/all/dict/list/len/print/range/str/class/super/异常类 等 ~60 个 + `json`/`math`/`re`/`datetime`/`statistics` 等安全模块可导入。注入的 `query/update/insert/delete/ExcelOperations` 直接可用。**大数据写入走 `update-range` 内联 JSON（<8KB）或拆批 upsert-row；超 8KB 时改用外部 `python -c` + openpyxl 直接写文件**

## 防错自查（调用前检查）

| # | 自查 | 错误信号 | 修正 |
|---|------|---------|------|
| 1 | 追加还是覆盖？ | 返回含"覆盖模式" | 追加加 `--insert-mode`，覆盖是默认 |
| 2 | 单行还是批量？ | 改一行却用 update-query | 单行用 `upsert-row` |
| 3 | 多表文件范围 | `range="A1:C10"` 报错 | 用 `"Sheet名!A1:C10"` |
| 4 | SQL 类型对吗？ | SELECT 传给 update-query | 查→query / 改→update-query / 增→insert-query / 删→delete-query |
| 5 | 写完验证了？ | 写完就结束 | 写入→query 验证 |
| 6 | Sheet名对吗？ | "表 XX 不存在" | 先 `list-sheets`，用 Sheet 名不是文件名 |
| 7 | 数值排序对吗？ | 排序结果不对 | 文本列用 `CAST(col AS INT)` 再排序 |

## 双行表头

双行表头（第1行中文描述 + 第2行英文字段名）时：
- **SQL 工具**：中文英文列名都能用
- **describe-table 自动检测**：返回第2行英文名作为列名
- **upsert-row 的 --key-column**：中英文都能用
- **建议**：用 describe-table 返回的英文名

## SQL 功能与准确率

### 已支持（169 条测试验证，19 类别）

| 类别 | 功能 |
|------|------|
| 基础 | SELECT, DISTINCT, AS, `+-*/%`, 一元负号, 整数除法(截断向零), `t.*` qualified star |
| 条件 | WHERE, LIKE, IN, NOT IN, BETWEEN, AND/OR, 子查询, **WHERE 引用 SELECT 别名** |
| 聚合 | COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING |
| 排序 | ORDER BY, LIMIT, OFFSET, NULLS FIRST/LAST |
| 窗口 | ROW_NUMBER, RANK, DENSE_RANK (含 PARTITION BY), **WHERE 引用窗口别名** |
| 多表 | INNER/LEFT/RIGHT/FULL JOIN (同文件跨Sheet) |
| 高级 | CASE WHEN (简单/搜索/嵌入表达式), CTE(WITH), EXISTS, UNION/UNION ALL |
| 字符串 | UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING |
| NULL | IS NULL, IS NOT NULL, COALESCE |
| 脚本 | run-python: class 定义/继承, 异常捕获, json/math/re/statistics 等安全模块 |

### 关键语义（与 SQLite 对齐）

- **GROUP BY / ORDER BY**：NULL 排最前（ASC），排最后（DESC）
- **ROUND**：四舍五入（round half away from zero），非 banker's rounding
- **整数除法**：`int / int` 截断向零（`-1/2 = 0`，非 `-0.5`）
- **LENGTH(NULL)**：返回 NULL
- **float32 降级**：已移除，保留 float64 保障聚合精度

### SQL 限制

- 跨文件 JOIN 需用 `表名@'文件路径'` 语法（同文件跨 Sheet 直接用表名）
- Excel 空字符串 `""` 往返后变为 NULL（xlsx 格式限制）
- WHERE 引用窗口函数别名时自动重写为子查询（透明支持，无需手动包装）
- WHERE 引用 SELECT 别名（非窗口）已支持（物化为临时列，与 SQLite 对齐）

## 边界

- **只处理**：`.xlsx` / `.xls` 配置表的读写查询
- **不负责**：数据库管理、CSV 处理、Excel 透视表

## 🔧 自我维护

**这个 skill 是自治模块——自己为自己的行为负责。** agent 用它时发现以下信号，**当场提议改本 SKILL.md**，不等用户开口：

| 信号 | 改哪里 |
|------|--------|
| 命令报错 / 参数过期 / 路径错了 | 改正文命令 |
| 流程走不通 / 步骤缺失 | 改「工作流」段 |
| 踩了坑（环境/代理/权限） | 写进「故障排查」或「注意事项」段 |
| 没覆盖用户要的能力 | 补章节，或提议新 skill |
| 触发不准（漏触发/误触发） | 改 frontmatter `keywords` / `description` |

SOP：识别 → 当场提议 → 用户确认 → 改 SKILL.md → 验证 → commit。
