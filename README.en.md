<div align="center">

[简体中文](README.md) ｜ [English](README.en.md)

</div>

# 🎮 ExcelMCP: Game Dev Excel Configuration Manager

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Powered by: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![Status](https://img.shields.io/badge/status-stable-green.svg)
![Tests](https://img.shields.io/badge/tests-773%20tests-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-41%20verified%20tools-green.svg)

**ExcelMCP** is an Excel configuration table management MCP server designed for game development. Use AI natural language commands to intelligently manage game configs like skill tables, equipment data, and monster attributes. Built with **FastMCP** and **openpyxl**, featuring **41 professional tools** and **773 test cases** for enterprise-grade reliability.

🎯 **Core Features**: Skill systems, equipment management, monster configuration, numerical balancing, version comparison, designer toolchain

📦 **One-line install**: `uvx excel-mcp-server-fastmcp` — run directly from PyPI, zero config

---

## 🚀 Quick Start

### Option 1: uvx One-line Run (Recommended)

No need to clone — run directly from PyPI:

```bash
uvx excel-mcp-server-fastmcp
```

MCP client configuration:
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["excel-mcp-server-fastmcp"]
    }
  }
}
```

### Option 2: Install from Source

1. **Clone the project**
   ```bash
   git clone https://github.com/TangentDomain/excel-mcp-server.git
   cd excel-mcp-server
   ```

2. **Install dependencies**
   ```bash
   # Recommended: Use uv (faster)
   pip install uv && uv sync

   # Alternative: Use pip
   pip install -e .
   ```

3. **Configure MCP client**
   ```json
   {
     "mcpServers": {
       "excelmcp": {
         "command": "python",
         "args": ["-m", "excel_mcp_server_fastmcp"]
       }
     }
   }
   ```

4. **Start using**
   Ready! Let your AI assistant control Excel files through natural language.

### Verify Installation
```bash
python -m pytest tests/ --tb=short -q
```

---

## ⚡ Quick Reference

### 🎯 Common Command Cheat Sheet

#### ⭐ Basic Operations (Beginner)
```text
Read data:      "Read data from range A1:C10 in sales.xlsx"
File info:      "Get basic information about report.xlsx"
Simple search:  "Find 'Fireball' in skills.xlsx"
```

#### ⭐⭐ Data Operations (Advanced)
```text
Update data:    "Multiply all values in column 2 of skills.xlsx by 1.2"
Format setting: "Make the first row of report.xlsx bold with light blue background"
Insert rows:    "Insert 3 empty rows at row 5 in inventory.xlsx"
```

#### ⭐⭐⭐ Game Dev Specialized (Expert)
```text
Config comparison:   "Compare skill tables v1.0 and v1.1, generate change report"
Batch analysis:      "Analyze HP/attack ratios for all level 20-30 monsters"
Attribute adjustment: "Increase attributes of legendary equipment by 25%"
```

### 🎮 Game Development Scenario Quick Reference

| Scenario | Recommended Tools | Example Command |
|----------|-------------------|-----------------|
| Skill balance adjustment | `excel_search` + `excel_update_range` | "Increase damage of all fire skills by 20%" |
| Equipment config management | `excel_format_cells` + `excel_get_range` | "Mark all legendary equipment with gold color" |
| Monster data validation | `excel_check_duplicate_ids` + `excel_search` | "Ensure monster IDs are unique and HP is reasonable" |
| Version comparison | `excel_compare_sheets` + `excel_compare_files` | "Compare old vs new version config tables" |
| Data statistics query | `excel_query` | "Query average attack power by class in skill table" |
| Conditional batch update | `excel_update_query` | "Increase fire skill damage by 20%" |
| Pre-modification preview | `excel_preview_operation` + `excel_assess_data_impact` | "Preview impact of deleting rows 5-10" |
| Pre-modification backup | `excel_create_backup` | "Backup skill table before modifying" |
| Formula evaluation | `excel_evaluate_formula` | "Temporarily calculate SUM(A2:A100) to see result" |

### 🔧 Range Expression Reference

| Format | Description | Example |
|--------|-------------|---------|
| `Sheet1!A1:C10` | Standard range | "SkillTable!A1:D50" |
| `Sheet1!1:5` | Row range | "ConfigTable!2:100" |
| `Sheet1!B:D` | Column range | "DataTable!B:G" |
| `Sheet1!A1` | Single cell | "SettingsTable!A1" |

---

## 🛠️ Complete Tool List (41 Professional Tools)

### 📁 File & Worksheet Management
- `excel_create_file` - Create new Excel files with custom worksheets
- `excel_create_sheet` - Add new worksheets
- `excel_delete_sheet` - Delete worksheets
- `excel_list_sheets` - List worksheet names
- `excel_rename_sheet` - Rename worksheets
- `excel_get_file_info` - Get file metadata
- `excel_get_sheet_headers` - Get all worksheet headers
- `excel_merge_files` - Merge multiple Excel files

### 📊 Data Operations
- `excel_get_range` - Read cell/row/column ranges
- `excel_update_range` - Write/update data ranges with formula preservation
- `excel_get_headers` - Extract headers from any row
- `excel_insert_rows` - Insert empty rows
- `excel_delete_rows` - Delete row ranges
- `excel_insert_columns` - Insert empty columns
- `excel_delete_columns` - Delete column ranges
- `excel_find_last_row` - Find last row with data
- `excel_set_formula` - Set cell formulas (auto-calculate)
- `excel_evaluate_formula` - Evaluate formulas without modifying files

### 🔍 Search & Analysis
- `excel_search` - Regex expression search
- `excel_search_directory` - Directory batch search
- `excel_query` - SQL query (dual-row headers, WHERE/GROUP BY/HAVING/ORDER BY/LIMIT/OFFSET/DISTINCT/JOIN/math expressions)
- `excel_update_query` - SQL UPDATE batch modification (SET constant/column reference/arithmetic, WHERE conditions, dry_run preview)
- `excel_describe_table` - View table structure (column names, types, descriptions, sample values, auto-detect dual-row headers)
- `excel_compare_sheets` - Worksheet comparison (game config optimized)
- `excel_compare_files` - Multi-worksheet file comparison
- `excel_check_duplicate_ids` - ID duplicate detection

### 🛡️ Safety & Backup
- `excel_create_backup` - Create file auto-backups
- `excel_restore_backup` - Restore from backup
- `excel_list_backups` - List all backup records
- `excel_preview_operation` - Preview operation scope and current data
- `excel_assess_data_impact` - Comprehensively assess potential operation impact

### 📜 Operation History
- `excel_get_operation_history` - Get operation history and statistics

### 🎨 Formatting & Styling
- `excel_format_cells` - Apply fonts, colors, alignment formats
- `excel_set_borders` - Set cell borders
- `excel_merge_cells` - Merge cell ranges
- `excel_unmerge_cells` - Unmerge cell ranges
- `excel_set_column_width` - Adjust column width
- `excel_set_row_height` - Adjust row height

### 🔄 Data Conversion
- `excel_export_to_csv` - Export CSV format
- `excel_import_from_csv` - Create Excel files from CSV
- `excel_convert_format` - Format conversion (.xlsx/.xlsm/.csv/.json)

---

## 📖 Usage Guide

### 🎮 Game Configuration Table Standard Format

**Dual-row header system** (Auto-detected for game dev):
```
Row 1 (Description): ['Skill ID Desc', 'Skill Name Desc', 'Skill Type Desc']
Row 2 (Field):       ['skill_id', 'skill_name', 'skill_type']
```

`excel_query` auto-detects dual-row headers (Row 1 descriptions + Row 2 field names). Query results include `column_descriptions` mapping.

**Common configuration table structures**:
- **Skill Table**: ID|Name|Type|Level|Cost|Cooldown|Damage|Description
- **Equipment Table**: ID|Name|Type|Quality|Attributes|Set|Acquisition
- **Monster Table**: ID|Name|Level|HP|Attack|Defense|Skills|Drops

### 📋 Standard Workflow

1. **Search & Locate**: Use `excel_search` to understand data distribution
2. **Determine Boundaries**: Use `excel_find_last_row` to confirm data range
3. **Read Current State**: Use `excel_get_range` to get current configuration
4. **Update Data**: Use `excel_update_range` for safe updates
5. **Beautify Display**: Use `excel_format_cells` to mark important data
6. **Verify Results**: Re-read to confirm successful updates

### 🔍 SQL Query Reference

`excel_query` is built on sqlglot + pandas, supporting the following SQL syntax:

**Supported syntax:**
```sql
-- Basic queries
SELECT * FROM SkillTable WHERE level >= 10 LIMIT 20
SELECT skill_name, damage FROM SkillTable ORDER BY damage DESC

-- Aggregation
SELECT skill_type, AVG(damage) as avg_dmg, COUNT(*) as cnt FROM SkillTable GROUP BY skill_type

-- HAVING filter
SELECT skill_type, SUM(damage) as total FROM SkillTable GROUP BY skill_type HAVING total > 1000

-- Math expressions
SELECT skill_name, damage * 1.2 as boosted_dmg FROM SkillTable WHERE level >= 5

-- LIKE fuzzy search
SELECT * FROM SkillTable WHERE skill_name LIKE '%fire%'

-- DISTINCT deduplication
SELECT DISTINCT skill_type FROM SkillTable

-- IN conditions
SELECT * FROM SkillTable WHERE skill_type IN ('Attack', 'Support')

-- BETWEEN range
SELECT * FROM MonsterTable WHERE level BETWEEN 10 AND 20

-- IS NULL / IS NOT NULL
SELECT * FROM SkillTable WHERE description IS NULL

-- OFFSET pagination
SELECT * FROM MonsterTable ORDER BY level LIMIT 20 OFFSET 0

-- NOT LIKE / NOT IN
SELECT * FROM SkillTable WHERE skill_name NOT LIKE '%test%'

-- JOIN cross-table queries (within same file)
SELECT a.skill_name, b.equip_name FROM SkillTable a INNER JOIN EquipTable b ON a.equip_id = b.equip_id
```

**SQL UPDATE batch modification:**
```sql
-- Constant update
UPDATE SkillTable SET damage = 500 WHERE skill_type = 'Ultimate'

-- Arithmetic expression (column reference)
UPDATE SkillTable SET damage = damage * 1.1 WHERE element = 'Fire'

-- Multi-column update
UPDATE SkillTable SET damage = damage * 1.1, cooldown = cooldown - 1 WHERE level >= 20

-- dry_run preview mode (no actual changes)
UPDATE SkillTable SET damage = damage * 1.1 WHERE element = 'Fire'  -- dry_run=True
```

**Unsupported syntax (with clear alternative suggestions):**
- CASE WHEN (suggest: conditional queries or external processing)
- Subqueries / CTE / Window functions
- RIGHT JOIN / CROSS JOIN (rarely used in game scenarios)

**Query performance:**
- Same-file repeated queries auto-cache, 30-100x speedup
- Small table (10 rows): first ~30-47ms, cached 2-5ms
- Large table (2000 rows): first ~230ms, cached 2-8ms
- Cache auto-invalidates on file modification

**Common problem solutions:**
- **File locked**: Close Excel program and retry
- **Encoding issues**: Ensure UTF-8 encoding
- **Large file slow**: Use precise ranges, process in batches
- **Memory insufficient**: Reduce single processing amount, close workbooks promptly
- **Permission issues**: Use admin privileges or check file properties

---

## 🔒 Security Mechanisms

ExcelMCP has built-in multi-layer security protections:

### Path Security (SecurityValidator)
- **Path traversal protection**: Rejects `../` directory traversal attacks
- **Symlink rejection**: Does not follow symlinks to prevent pointing to sensitive files
- **Hidden file rejection**: Does not process files starting with `.`
- **Extension whitelist**: Only allows `.xlsx`/`.xlsm`/`.xls`/`.csv`/`.json`/`.bak`
- **File size limit**: Max 50MB per file

### Formula Injection Protection
- **DDE detection**: Rejects `=DDE()` dynamic data exchange formulas
- **CMD detection**: Rejects `=CMD()` system command execution
- **SHELL detection**: Rejects `=SHELL()` shell command formulas
- **REGISTER detection**: Rejects `=REGISTER()` DLL registration formulas
- **Pipe detection**: Rejects dangerous formulas containing pipe characters

### Data Security
- **File lock**: `excel_update_query` uses file locks (fcntl LOCK_EX) to prevent concurrent write conflicts
- **Transaction protection**: Auto-creates backups before UPDATE, auto-rollback on failure
- **Temp file cleanup**: Auto-cleans orphan `.bak` temp files older than 1 hour on startup

### Error Messages
- Security errors prefixed with 🔒, including specific rejection reason
- Example: `🔒 Security validation failed: path contains illegal characters '..'`

---

## 🏗️ Technical Architecture

### Package Structure
```
src/excel_mcp_server_fastmcp/    # Main package (directly importable after pip install)
├── __init__.py                   # Package entry point, exposes main()
├── server.py                     # MCP interface layer (41 tool definitions)
├── api/                          # API business logic layer
│   ├── excel_operations.py       # Unified Excel operations entry
│   └── advanced_sql_query.py     # SQL query engine
├── core/                         # Core operation layer
│   ├── excel_reader.py           # Read operations
│   ├── excel_writer.py           # Write operations
│   ├── excel_search.py           # Search operations
│   ├── excel_manager.py          # Workbook management
│   ├── excel_compare.py          # Version comparison
│   └── excel_converter.py        # Format conversion
├── models/                       # Data models
│   └── types.py                  # Type definitions
└── utils/                        # Utility layer
    ├── validators.py             # Path/data validation + security
    ├── error_handler.py          # Unified error handling
    ├── formatter.py              # Result formatting
    ├── parsers.py                # Parameter parsing
    ├── temp_file_manager.py      # Temp file management
    ├── formula_cache.py          # Formula cache
    └── exceptions.py             # Custom exceptions
```

### Layered Design Pattern
```
MCP Interface Layer (Pure Delegation)
    ↓
API Business Logic Layer (Centralized Processing)
    ↓
Core Operation Layer (Single Responsibility)
    ↓
Tool Layer (Common Functions)
```

### Core Features
- **Pure Delegation Pattern**: Interface layer has zero business logic
- **Centralized Processing**: Unified validation, error handling, result formatting
- **1-Based Indexing**: Matches Excel user habits (Row 1 = First row)
- **Workbook Caching**: 75% performance improvement on cache hits
- **Realistic Concurrency Handling**: Properly handles Excel file concurrency limits

### Performance Optimization
- **Precise Range Reading**: 60-80% faster than reading entire tables
- **Batch Operations**: 15-20x faster than individual operations
- **Batch Processing**: 70% memory usage reduction for large files

---

## 📊 Project Information

### Quality Metrics
- **Test Cases**: 773 (behavior-validated, no coverage padding)
- **Test Files**: 33 test files
- **Test Code**: 13,574 lines
- **Tool Count**: 41 (verified with @mcp.tool decorators)
- **Architecture Layers**: 4-layer design (MCP→API→Core→Utils)

### Verification Commands
```bash
# Run full test suite (parallel accelerated)
python -m pytest tests/ -q --tb=short -n auto --timeout=30

# Verify tool completeness
grep -c "def excel_" src/excel_mcp_server_fastmcp/server.py  # Should output: 41

# Generate coverage report
python -m pytest tests/ --cov=src --cov-report=html
```

### Development Standards
- **Pure Delegation Pattern**: server.py strictly delegates to ExcelOperations
- **Centralized Business Logic**: Unified validation, error handling, result formatting
- **Branch Naming**: All feature branches must start with `feature/`
- **Test Coverage**: Maintain 80%+ test coverage

---

## ❓ FAQ

### Basic Questions
**Q: Which Excel formats are supported?**
A: Supports `.xlsx`, `.xlsm` formats, with `.csv` support through import/export

**Q: How to handle non-English worksheet names?**
A: Fully supports CJK and other worksheet names and content

**Q: How is large file processing performance?**
A: SQL queries auto-cache DataFrames, 30-100x speedup on repeated queries. Large tables (2000 rows): first ~230ms, cached 2-8ms.

**Q: How to ensure data security?**
A: Complete error handling, formula preservation by default, operation preview support

### Game Development
**Q: What is the dual-row header system?**
A: Game config table standard format: Row 1 field descriptions, Row 2 field names

**Q: How to perform version comparison?**
A: Use specialized config table comparison tools with ID object tracking

---

## 🤝 Contributing

**Ways to contribute**:
- 🐛 **Report Bugs**: Report issues through GitHub Issues
- 💡 **Feature Suggestions**: Propose new feature requirements
- 📝 **Documentation**: Improve usage guides and technical docs
- 🔧 **Code**: Follow development standards, submit PRs

**License**: MIT License - See [LICENSE](LICENSE) file for details

---

<div align="center">

## 🔝 Quick Navigation

| 🎯 **Quick Start** | 🛠️ **Tool Reference** | 📚 **Learning Guide** |
|-------------------|------------------------|---------------------|
| [🚀 Installation](#-quick-start) | [📋 Complete Tool List](#️-complete-tool-list-41-professional-tools) | [📖 Usage Guide](#-usage-guide) |
| [⚡ Command Cheat Sheet](#-quick-reference) | [🏗️ Technical Architecture](#️-technical-architecture) | [🔒 Security](#-security-mechanisms) |
| [🎮 Game Config Management](#-usage-guide) | [📊 Project Info](#-project-information) | [❓ FAQ](#-faq) |

**[⬆️ Back to Top](#-excelmcp-game-dev-excel-configuration-manager)**

*✨ Making game configuration table management simple and efficient ✨*

</div>
