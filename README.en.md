[简体中文](README.md) | [English](README.en.md)

# ExcelMCP: Excel Config Table MCP Server for Game Development

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1447-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-26-green.svg)
![SQL](https://img.shields.io/badge/SQL%20accuracy-100%25-brightgreen.svg)

> Excel configuration table MCP server built on Python FastMCP + openpyxl + sqlglot.
> SQL-over-Excel queries, batch operations, and structure management designed for game developers.

---

## Quick Start

### Option 1: Skill Integration (Recommended)

Install via global skill with self-update and self-bootstrap support:

```bash
excel-cli <command> [options]
```

The skill entry auto-manages venv and forwards to CLI. See [SKILL.md](.omp/skills/excel/SKILL.md).

### Option 2: MCP Server (AI Client Integration)

```bash
# uvx (recommended, no install needed)
uvx excel-mcp-server-fastmcp

# pip
pip install excel-mcp-server-fastmcp
```

#### Cursor
- Settings → MCP → Add Server
- Name: `excelmcp`, Command: `uvx`, Args: `"excel-mcp-server-fastmcp"`

#### Claude Code
```bash
claude mcp add excelmcp -- uvx excel-mcp-server-fastmcp
```

### Option 3: Source (Development)

```bash
uv run python -m excel_mcp_server_fastmcp.cli <command> [options]
```

---

## 26 MCP Tools

### Query (9)

| Tool | Description |
|------|-------------|
| `excel_query` | **SQL engine** (primary) — WHERE/LIKE/IN/JOIN/window functions/CTE/UNION |
| `excel_describe_table` | Table structure (column names + types + sample values) |
| `excel_get_headers` | Header info (Chinese + English) |
| `excel_get_range` | Read by exact range (e.g. A1:C10) |
| `excel_search` | Search cell text in a worksheet |
| `excel_search_directory` | Search across Excel files |
| `excel_find_last_row` | Find last data row (use before appending) |
| `excel_list_sheets` | List all worksheet names |
| `excel_compare_sheets` | Compare two sheets by ID column |

### Write (7)

| Tool | Description |
|------|-------------|
| `excel_update_query` | **SQL UPDATE** batch modify (supports dry_run) |
| `excel_insert_query` | **SQL INSERT** (single/multi-row) |
| `excel_delete_query` | **SQL DELETE** (WHERE required) |
| `excel_update_range` | Write by exact range (overwrite by default, insert_mode=True for insert) |
| `excel_upsert_row` | Insert or update single row by key (idempotent) |
| `excel_set_formula` | Write Excel formula |
| `excel_run_python` | Execute Python script (sandboxed, with query/update/insert/delete injected) |

### Structure (7)

| Tool | Description |
|------|-------------|
| `excel_create_file` | Create new Excel file |
| `excel_create_sheet` | Create worksheet |
| `excel_delete_sheet` | Delete worksheet |
| `excel_rename_sheet` | Rename worksheet |
| `excel_copy_sheet` | Copy worksheet |
| `excel_structure` | Insert/delete rows and columns |
| `excel_rename_column` | Rename column header |

### Formatting (2)

| Tool | Description |
|------|-------------|
| `excel_format_cells` | Set cell style (font/merge/border/preset) |
| `excel_set_layout` | Set row height or column width |

### Backup (1)

| Tool | Description |
|------|-------------|
| `excel_backup` | Backup create/list/restore |

---

## SQL Features

### Supported (169 differential tests, 19 categories, 100% pass rate)

| Category | Features |
|----------|----------|
| Basic | SELECT, DISTINCT, AS, `+-*/%`, unary minus, integer division (trunc toward zero), `t.*` qualified star |
| Conditions | WHERE, LIKE, IN, NOT IN, BETWEEN, AND/OR, subqueries, WHERE referencing SELECT aliases |
| Aggregation | COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING |
| Sorting | ORDER BY, LIMIT, OFFSET, NULLS FIRST/LAST |
| Window | ROW_NUMBER, RANK, DENSE_RANK, NTILE, LAG, LEAD, FIRST_VALUE, LAST_VALUE, NTH_VALUE, AVG/SUM/MIN/MAX/COUNT OVER, GROUP_CONCAT, PARTITION BY, ROWS BETWEEN, WHERE referencing window aliases |
| Multi-table | INNER/LEFT/RIGHT/FULL JOIN (same-file cross-sheet + cross-file `table@'path'`) |
| Advanced | CASE WHEN, CTE(WITH), EXISTS, UNION/UNION ALL, INTERSECT/EXCEPT, NULLIF, COALESCE |
| String | UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING |
| Math | ABS, CEIL, FLOOR, SQRT, POWER, ROUND |
| NULL | IS NULL, IS NOT NULL, COALESCE, three-valued logic (aligned with SQLite) |

### Key Semantics (aligned with SQLite 3.x)

- **GROUP BY / ORDER BY**: NULL sorts first (ASC), last (DESC)
- **ROUND**: Round half away from zero (not banker's rounding)
- **Integer division**: Truncation toward zero, same as SQLite
- **NULL three-valued logic**: NULL = NULL → UNKNOWN(FALSE)
- **LIKE**: `%` matches any chars, `_` matches single char, case-insensitive

### SQL Limitations

- Cross-file JOIN requires `table@'file_path'` syntax
- Excel empty string `""` round-trips to NULL (xlsx format limitation)
- WHERE referencing window function aliases is auto-rewritten to subquery (transparent)
- WHERE referencing SELECT aliases (non-window) is supported via temp column materialization

---

## Technical Specs

- **Version**: 1.17.0
- **Python**: >= 3.10
- **Dependencies**: FastMCP / openpyxl / sqlglot / pandas
- **Tests**: 1447 passed, 3 skipped, 1 xfailed
- **SQL accuracy**: 169 differential tests 100% pass (cross-validated with SQLite)
- **Tools**: 26 MCP tools
- **Formats**: .xlsx, .xlsm

---

## Development

```bash
uv sync --extra dev                                          # Install deps
uv run python -m pytest tests/ -q --timeout=60               # Run all tests
uv run python -m pytest tests/invariants/ -q --timeout=30     # Invariant tests
ruff check src/ tests/ && ruff format --check src/ tests/     # Lint
```

See [Development Guide](docs/DEVELOPMENT.md).

---

## Troubleshooting

### MCP connection failed
```bash
uv --version
uvx excel-mcp-server-fastmcp --force-reinstall
# Restart AI client
```

### Excel file read failed
- Use full file path (not `~/`)
- Confirm `.xlsx` format
- Close the file in Excel

### Large file slow
- Use WHERE to filter data
- Process in batches

---

## Contributing

- [GitHub Issues](https://github.com/TangentDomain/excel-mcp-server/issues)
- [Contributing Guide](docs/CONTRIBUTING.md)
- [Changelog](docs/CHANGELOG.md)

## License

[MIT License](LICENSE)
