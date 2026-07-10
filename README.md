[简体中文](README.md) | [English](README.en.md)

# ExcelMCP: 游戏开发 Excel 配置表 MCP 服务器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1447-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-26-green.svg)
![SQL](https://img.shields.io/badge/SQL%20accuracy-100%25-brightgreen.svg)

> 基于 Python FastMCP + openpyxl + sqlglot 的 Excel 配置表 MCP 服务器。
> SQL-over-Excel 查询、批量操作、结构管理，为游戏开发者设计。

---

## 快速开始

### 方式一：Skill 接入（首推）

将 Excel skill 接入你的 OMP harness，首次调用自动创建 venv + 从 GitHub 安装，后续通过 `self-update` 更新。

#### 全局接入（所有项目可用）

```bash
# 1. 下载 skill 文件（只需 SKILL.md + bin/excel-cli.py）
git clone https://github.com/TangentDomain/excel-mcp-server.git
cp -r excel-mcp-server/.omp/skills/excel ~/.omp/agent/skills/excel

# 2. 首次运行——自动创建 venv + 从 GitHub 安装依赖
python ~/.omp/agent/skills/excel/bin/excel-cli.py query --file data.xlsx --sql "SELECT * FROM Sheet1"

# 3. 后续可直接用 exe（跳过自举，更快）
~/.omp/agent/skills/excel/.venv/Scripts/excel-cli.exe query --file data.xlsx --sql "SELECT * FROM Sheet1"
```

#### 项目级接入（仅当前项目）

```bash
# 在你的项目根目录下
cp -r /path/to/excel-mcp-server/.omp/skills/excel .omp/skills/excel

# 使用方式同上，路径改为项目内
python .omp/skills/excel/bin/excel-cli.py query --file data.xlsx --sql "SELECT * FROM Sheet1"
```

> **全局 vs 项目级**：全局装一次所有项目可用；项目级随项目走，适合团队协作（skill 定义随仓库提交）。
> 项目级优先级高于全局——同名 skill 项目版覆盖全局版。

#### 更新

```bash
excel-cli self-update --check   # 检查是否有新版本
excel-cli self-update           # 更新到最新版
```

详见 [SKILL.md](.omp/skills/excel/SKILL.md)。

### 方式二：MCP Server（AI 客户端集成）

```bash
# uvx（推荐，无需安装）
uvx excel-mcp-server-fastmcp

# pip
pip install excel-mcp-server-fastmcp
```

#### Cursor 配置
- 设置 → MCP → Add Server
- Name: `excelmcp`
- Command: `uvx`
- Args: `"excel-mcp-server-fastmcp"`

#### Claude Code 配置
```bash
claude mcp add excelmcp -- uvx excel-mcp-server-fastmcp
```

### 方式三：源码直接调用（开发环境）

```bash
uv run python -m excel_mcp_server_fastmcp.cli <command> [options]
```

---

## 26 个 MCP 工具

### 查询类（9 个）

| 工具 | 说明 |
|------|------|
| `excel_query` | **SQL 查询引擎**（首选）— WHERE/LIKE/IN/JOIN/窗口函数/CTE/UNION 等 |
| `excel_describe_table` | 查看表结构（列名+类型+样本值），支持双行表头自动检测 |
| `excel_get_headers` | 获取表头信息（中文+英文） |
| `excel_get_range` | 按精确坐标读取数据（如 A1:C10） |
| `excel_search` | 在工作表中搜索单元格文本 |
| `excel_search_directory` | 跨文件搜索 Excel |
| `excel_find_last_row` | 定位数据末行（追加数据前必用） |
| `excel_list_sheets` | 列出所有工作表名称 |
| `excel_compare_sheets` | 按 ID 列对比两个工作表差异 |

### 写入类（7 个）

| 工具 | 说明 |
|------|------|
| `excel_update_query` | **SQL UPDATE** 批量修改（支持 dry_run 预览） |
| `excel_insert_query` | **SQL INSERT** 插入数据（单行/多行） |
| `excel_delete_query` | **SQL DELETE** 删除数据（必须 WHERE） |
| `excel_update_range` | 精确坐标写入（默认覆盖，insert_mode=True 插入） |
| `excel_upsert_row` | 按主键插入或更新单行（幂等安全） |
| `excel_set_formula` | 写入 Excel 公式 |
| `excel_run_python` | 执行 Python 脚本（沙箱环境，注入 query/update/insert/delete） |

### 结构操作类（7 个）

| 工具 | 说明 |
|------|------|
| `excel_create_file` | 创建新 Excel 文件 |
| `excel_create_sheet` | 创建工作表 |
| `excel_delete_sheet` | 删除工作表 |
| `excel_rename_sheet` | 重命名工作表 |
| `excel_copy_sheet` | 复制工作表 |
| `excel_structure` | 插入/删除行列 |
| `excel_rename_column` | 重命名列（表头） |

### 格式化类（2 个）

| 工具 | 说明 |
|------|------|
| `excel_format_cells` | 设置样式（字体/合并/边框/预设样式） |
| `excel_set_layout` | 设置行高或列宽 |

### 备份类（1 个）

| 工具 | 说明 |
|------|------|
| `excel_backup` | 备份创建/列表/恢复 |

---

## SQL 功能

### 已支持（169 条差分测试验证，19 类别 100% 通过）

| 类别 | 功能 |
|------|------|
| 基础 | SELECT, DISTINCT, AS, `+-*/%`, 一元负号, 整数除法(截断向零), `t.*` qualified star |
| 条件 | WHERE, LIKE, IN, NOT IN, BETWEEN, AND/OR, 子查询, WHERE 引用 SELECT 别名 |
| 聚合 | COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING |
| 排序 | ORDER BY, LIMIT, OFFSET, NULLS FIRST/LAST |
| 窗口 | ROW_NUMBER, RANK, DENSE_RANK, NTILE, LAG, LEAD, FIRST_VALUE, LAST_VALUE, NTH_VALUE, AVG/SUM/MIN/MAX/COUNT OVER, GROUP_CONCAT, PARTITION BY, ROWS BETWEEN, WHERE 引用窗口别名 |
| 多表 | INNER/LEFT/RIGHT/FULL JOIN（同文件跨 Sheet + 跨文件 `表名@'路径'`） |
| 高级 | CASE WHEN, CTE(WITH), EXISTS, UNION/UNION ALL, INTERSECT/EXCEPT, NULLIF, COALESCE |
| 字符串 | UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING |
| 数学 | ABS, CEIL, FLOOR, SQRT, POWER, ROUND |
| NULL | IS NULL, IS NOT NULL, COALESCE, 三值逻辑（与 SQLite 对齐） |

### 关键语义（与 SQLite 3.x 对齐）

- **GROUP BY / ORDER BY**：NULL 排最前（ASC），排最后（DESC）
- **ROUND**：四舍五入（round half away from zero），非 banker's rounding
- **整数除法**：截断向零（int(a/b)），与 SQLite 一致
- **NULL 三值逻辑**：NULL = NULL → UNKNOWN(FALSE)，NULL != 0 → UNKNOWN(FALSE)
- **LIKE**：`%` 匹配任意字符，`_` 匹配单字符，大小写不敏感

### SQL 限制

- 跨文件 JOIN 需用 `表名@'文件路径'` 语法（同文件跨 Sheet 直接用表名）
- Excel 空字符串 `""` 往返后变为 NULL（xlsx 格式固有限制）
- WHERE 引用窗口函数别名时自动重写为子查询（透明支持）
- WHERE 引用 SELECT 别名（非窗口）已支持（物化为临时列）

---

## 技术规格

- **版本**: 1.17.0
- **Python**: >= 3.10
- **依赖**: FastMCP / openpyxl / sqlglot / pandas
- **测试**: 1447 passed, 3 skipped, 1 xfailed
- **SQL 准确率**: 169 条差分测试 100% 通过（与 SQLite 交叉校验）
- **工具数量**: 26 个 MCP 工具
- **支持格式**: .xlsx, .xlsm

---

## 架构

```
server.py                    MCP 工具层 (FastMCP) — 26 个工具
  └─ api/
       ├─ advanced_sql_query.py   SQL 查询引擎 (10395 行)
       ├─ excel_operations.py     通用 Excel 操作 (2776 行)
       ├─ script_runner.py        Python 脚本沙箱 (281 行)
       └─ header_analyzer.py      双行表头检测
  └─ core/
       ├─ excel_reader.py         读取 (calamine → openpyxl 降级)
       ├─ excel_writer.py         写入 (传统模式, 1948 行)
       ├─ excel_manager.py        工作表管理
       ├─ excel_search.py         搜索
       ├─ excel_compare.py        比较
       └─ excel_converter.py      格式转换
  └─ utils/
       ├─ validators.py           SecurityValidator + ExcelValidator
       ├─ formatter.py            结果格式化
       ├─ formula_cache.py        公式计算缓存
       └─ concurrent_utils.py     并发工具
  └─ calibrator/
       └─ core.py                 SQLite 交叉校准工具
  └─ verification/
       └─ runner.py               baseline 驱动验证
```

### 关键设计

- **双表头支持**：自动检测游戏配表常见的双层表头（中文描述 + 英文字段名），SQL 工具中英文名都可用
- **性能路径**：calamine (Rust 引擎) 纯数据读取 → openpyxl 格式化/公式读取降级 → StreamingWriter 大文件流式写入
- **安全**：所有工具路径验证（防穿越/符号链接），SQL 通过 sqlglot AST 解析（非拼接），run-python 沙箱限制文件/进程操作
- **SQL 校准器**：将 Excel 导入 SQLite 跑同一条 SQL 做对比，定位引擎 bug（开发调试用）

---

## SQL 校准器（开发调试工具）

将 Excel 导入 SQLite 后跑同一条 SQL，跟 `excel_query` 的返回结果做对比，定位 bug。

### CLI 使用

```bash
# 导入 Excel 到 SQLite
python -m excel_mcp_server_fastmcp.calibrate import <xlsx路径> [数据库名]

# 执行查询
python -m excel_mcp_server_fastmcp.calibrate query <数据库名> "<SQL>"

# 列出所有表
python -m excel_mcp_server_fastmcp.calibrate tables [数据库名]

# 查看表结构
python -m excel_mcp_server_fastmcp.calibrate schema <数据库名> <表名>
```

### Python API

```python
from excel_mcp_server_fastmcp.calibrator.core import cmd_import, cmd_query, cmd_tables, cmd_schema

result = cmd_import("/path/to/data.xlsx", "my_db")
result = cmd_query("my_db", "SELECT * FROM table1 LIMIT 10")
result = cmd_tables("my_db")
result = cmd_schema("my_db", "table1")
```

---

## 开发

```bash
# 安装依赖
uv sync --extra dev

# 运行全部测试
uv run python -m pytest tests/ -q --timeout=60

# 运行不变量测试
uv run python -m pytest tests/invariants/ -q --timeout=30

# Lint
ruff check src/ tests/ && ruff format --check src/ tests/
```

### 测试体系

| 层级 | 目录 | 说明 |
|------|------|------|
| L1 结果结构 | `tests/invariants/test_l1_result_structure.py` | API 返回格式不变量 |
| L2 架构 | `tests/invariants/test_l2_architecture.py` | 代码结构不变量 |
| L3 SQL 功能 | `tests/invariants/test_l3_*.py` | SQL 准确率差分测试 |
| L4 限制消除 | `tests/invariants/test_l4_limit_fixes.py` | 引擎限制修复验证 |
| 对抗测试 | `tests/adversarial/` | 随机 fuzz 读写 |
| 功能测试 | `tests/test_*.py` | 各模块功能测试 |

详见 [开发者指南](docs/DEVELOPMENT.md)。

---

## 常见问题

### MCP 连接失败
```bash
uv --version              # 确认 uv 已安装
uvx excel-mcp-server-fastmcp --force-reinstall  # 重装
# 重启 AI 客户端
```

### Excel 文件读取失败
- 文件路径要完整（不要用 `~/`）
- 确认文件是 `.xlsx` 格式
- 文件没有被 Excel 软件打开

### 大文件卡顿
- 用 WHERE 过滤减少数据量
- 分批处理：`"先读取前1000行"`

---

## 贡献

- [GitHub Issues](https://github.com/TangentDomain/excel-mcp-server/issues) — 报告 Bug / 功能建议
- [贡献指南](docs/CONTRIBUTING.md)
- [更新日志](docs/CHANGELOG.md)

## 许可证

[MIT License](LICENSE)
