<div align="center">

[简体中文](README.md) ｜ [English](README.en.md)

</div>

# 🎮 ExcelMCP: Game Dev Excel Configuration Table Manager

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Powered by: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![Status](https://img.shields.io/badge/status-stable-green.svg)
![Tests](https://img.shields.io/badge/tests-938%20tests-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-46%20verified%20tools-green.svg)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)

**ExcelMCP** is an Excel configuration table management MCP server designed for game development. Through AI natural language commands, it enables intelligent operations on game configurations such as skill tables, equipment data, and monster attributes. Built with **FastMCP**, reads use **python-calamine** (Rust engine, 2300x speedup), writes use **openpyxl**. Features **46 professional tools** and **938 test cases**, ensuring enterprise-grade reliability.

🎯 **Core Features**: Skill systems, equipment management, monster configuration, numerical balancing, version comparison, designer toolchain

📦 **One-line install**: `uvx --force excel-mcp-server-fastmcp` — run from PyPI, auto-updates, zero config

---

## 🚀 Quick Start

### Option 1: uvx One-line Run (Recommended)

No need to clone the repo — run directly from PyPI:

```bash
uvx excel-mcp-server-fastmcp
```

MCP client configuration:
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["--force", "excel-mcp-server-fastmcp"]
    }
  }
}
```

> ⚡ **Recommended: add `--force`**: Skips local cache and automatically fetches the latest version from PyPI. Won't re-download when already up-to-date (just 1-2s check). New versions are picked up automatically without manual intervention.

> 💡 **Debug mode**: Set environment variable `EXCEL_MCP_DEBUG=1` to enable verbose logging (default: WARNING level). Set `EXCEL_MCP_JSON_LOG=1` for structured JSON logging (one JSON object per line with ts/level/tool/duration_ms fields).

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
# Check version
excel-mcp-server-fastmcp --version

# Run tests
python -m pytest tests/ --tb=short -q

# Run performance benchmarks
python scripts/benchmark.py --quick        # Quick mode (~30s)
python scripts/benchmark.py                # Full mode (includes large table tests)
python scripts/benchmark.py --compare      # Compare with previous results
```

---

## ⚡ Quick Reference

### 🎯 Common Command Cheat Sheet

#### ⭐ Basic Operations (Beginner Level)
```text
Read data:      "Read data from range A1:C10 in sales.xlsx"
File info:      "Get basic information about report.xlsx"
Simple search:  "Find 'Fireball' in skills.xlsx"
```

#### ⭐⭐ Data Operations (Advanced Level)
```text
Update data:    "Multiply all values in column 2 of skills.xlsx by 1.2"
Format setting: "Make the first row of report.xlsx bold with light blue background"
Insert rows:    "Insert 3 empty rows at row 5 in inventory.xlsx"
```

#### ⭐⭐⭐ Game Development Specialized (Expert Level)
```text
Config comparison: "Compare skill tables v1.0 and v1.1, generate change report"
Batch analysis:    "Analyze HP/attack ratios for all level 20-30 monsters"
Attribute adjustment: "Increase attributes of legendary equipment by 25%"
```

### 🎮 Game Development Scenario Quick Reference

| Scenario | Recommended Tools | Example Command |
|----------|-------------------|-----------------|
| Skill balance adjustment | `excel_search` + `excel_update_range` | "Increase damage of all fire skills by 20%" |
| Equipment config management | `excel_format_cells` + `excel_get_range` | "Mark all legendary equipment with gold color" |
| Monster data validation | `excel_check_duplicate_ids` + `excel_search` | "Ensure monster IDs are unique and HP is reasonable" |
| Version comparison | `excel_compare_sheets` + `excel_compare_files` | "Compare differences between old and new version config tables" |
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

## 🎮 Game Designer Complete Workflow Tutorial

> Step-by-step guide to operating Excel configuration tables with natural language. No need to memorize any command format — just describe what you want in Chinese.

### 📦 Step 1: Understand Your Table (DESCRIBE)

```
"View the skill table structure of skills.xlsx"
→ excel_describe_table("skills.xlsx", "SkillConfig")
```

Returns:
```
Column       | Type   | Description | Non-null | Sample Values
skill_id     | int    | Skill ID    | 10/10    | 1001, 1002
skill_name   | str    | Skill Name  | 10/10    | Fireball, Heal
damage       | float  | Damage      | 9/10     | 150.0, 200.5
cooldown     | int    | Cooldown    | 10/10    | 5, 10
```

💡 **Tip**: Dual-row header config tables (Row 1 Chinese description + Row 2 English field name) are auto-detected.

### 🔍 Step 2: Search & Locate (SEARCH)

```
"Search for all fire skills in skills.xlsx"
→ excel_search("skills.xlsx", "Fire")

"Search for all equipment containing 'Legendary'"
→ excel_search("equipment.xlsx", "Legendary", "EquipmentConfig")
```

### 📊 Step 3: SQL Query Analysis (QUERY)

The most powerful feature. Use standard SQL syntax to query config tables:

**Basic Queries — Find Data:**
```sql
-- View all skills above level 10
SELECT * FROM SkillConfig WHERE level >= 10

-- Show skill names and damage, sorted by damage
SELECT skill_name, damage FROM SkillConfig ORDER BY damage DESC LIMIT 10

-- Pagination: 5 per page, page 3
SELECT * FROM MonsterConfig ORDER BY level LIMIT 5 OFFSET 10
```

**Chinese Column Name Queries — Designer Friendly:**
```sql
-- Use Chinese column names directly with dual-row headers
SELECT 技能名称, 伤害值 FROM SkillConfig WHERE 等级 >= 10

-- Mix Chinese and English column names
SELECT skill_name, 伤害值 FROM SkillConfig WHERE 技能类型 = '攻击'
```

**Aggregation Statistics — Numerical Analysis:**
```sql
-- Average damage per class
SELECT skill_type, AVG(damage) as avg_dmg, COUNT(*) as cnt
FROM SkillConfig GROUP BY skill_type

-- Skill types with total damage over 1000
SELECT skill_type, SUM(damage) as total
FROM SkillConfig GROUP BY skill_type HAVING total > 1000

-- Equipment quality distribution
SELECT DISTINCT quality FROM EquipmentConfig
```

**DPM Numerical Balancing:**
```sql
-- Damage per second ranking (DPM = damage / cooldown)
SELECT skill_name, damage * 1.0 / cooldown as dpm
FROM SkillConfig ORDER BY dpm DESC LIMIT 10
```

**Data Quality Checks:**
```sql
-- Find configs with missing values
SELECT skill_name, description FROM SkillConfig WHERE description IS NULL

-- Find monsters in specific level range
SELECT name, level, hp FROM MonsterConfig WHERE level BETWEEN 10 AND 20

-- Exclude test data
SELECT * FROM SkillConfig WHERE skill_name NOT LIKE '%测试%'
```

### ✏️ Step 4: Batch Modification (UPDATE)

**Method 1: SQL UPDATE (Recommended, precise conditional modification):**
```
"Increase all fire skill damage by 20%"
→ excel_update_query("skills.xlsx", "UPDATE SkillConfig SET damage = damage * 1.2 WHERE skill_type = 'Fire'")
```

⚠️ **Preview before modifying:**
```
"Preview what fire skill damage +20% would change"
→ excel_update_query("skills.xlsx", "UPDATE SkillConfig SET damage = damage * 1.2 WHERE skill_type = 'Fire'", dry_run=True)
```

**Method 2: Range Write (known area batch write):**
```
"Multiply damage column values in rows 2-50 of skills.xlsx by 1.15"
→ excel_update_range("skills.xlsx", "SkillConfig!E2:E50", [[...]])
```

⚠️ **Always backup before modifying:**
```
"Backup skills.xlsx"
→ excel_create_backup("skills.xlsx")
```

### 🔄 Step 5: Version Comparison (COMPARE)

```
"Compare skill table differences between v1.0 and v1.1"
→ excel_compare_sheets("skills_v1.0.xlsx", "SkillConfig",
                        "skills_v1.1.xlsx", "SkillConfig")
```

### 📋 Common Designer Scenarios

| I want to... | How to say it |
|-------------|---------------|
| See what's in the table | "View xxx table structure" |
| Find a specific skill/equipment | "Search xxx" |
| Filter by conditions | "Query skills with level > 10" |
| Count by type | "How many skills per class" |
| Find strongest skill | "Top 10 skills by DPM" |
| Find problematic data | "Which skills have empty descriptions" |
| Batch modify values | "Increase all fire skill damage by 20%" |
| Conditional batch update | "UPDATE SkillTable SET Damage=Damage*1.1 WHERE Element='Fire'" |
| Compare version differences | "Compare v1 and v2 config tables" |

### ❓ Common Errors & Solutions

**Q: What if I misspell a column name?**
A: The system auto-recommends similar column names. E.g., `skil_name` → "Did you mean: skill_name?"

**Q: Table too large, slow query?**
A: Repeated queries on the same table auto-cache. Second query is 30-100x faster. 2000-row table: first ~230ms, cached 2-8ms.

**Q: How to use JOIN?**
A: Supports cross-sheet joins within the same file:
```sql
SELECT a.skill_name, b.equip_name FROM SkillConfig a INNER JOIN EquipConfig b ON a.equip_id = b.equip_id
```

---

## 🛠️ Complete Tool List (46 Professional Tools)

### 📁 File & Worksheet Management
- `excel_create_file` - Create new Excel files with custom worksheets
- `excel_create_sheet` - Add new worksheets
- `excel_delete_sheet` - Delete worksheets
- `excel_list_sheets` - List worksheet names
- `excel_rename_sheet` - Rename worksheets
- `excel_copy_sheet` - Copy worksheets (with data and formatting), for creating config table variants
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
- `excel_rename_column` - Rename columns (modify header cell values, supports dual-header tables)
- `excel_upsert_row` - Upsert row (find by key column, update if exists, insert if not)
- `excel_batch_insert_rows` - Batch insert multiple rows at end of sheet
- `excel_set_formula` - Set cell formulas (auto-calculate)
- `excel_evaluate_formula` - Evaluate formulas without writing to files

### 🔍 Search & Analysis
- `excel_search` - Regex expression search
- `excel_search_directory` - Directory batch search
- `excel_query` - SQL query (supports dual-row headers, WHERE/GROUP BY/HAVING/ORDER BY/LIMIT/OFFSET/DISTINCT/JOIN/UNION/subqueries/CTE/CASE WHEN/COALESCE/string functions/math expressions/window functions)
- `excel_update_query` - SQL UPDATE batch modification (SET constant/column ref/arithmetic, WHERE condition, dry_run preview)
- `excel_describe_table` - View table structure (column names, types, descriptions, sample values, auto-detect dual-row headers)
- `excel_compare_sheets` - Worksheet comparison (game config optimized)
- `excel_compare_files` - Multi-worksheet file comparison
- `excel_check_duplicate_ids` - ID duplicate detection
- `excel_server_stats` - Server runtime statistics (tool call count, latency, error rate, error classification)

### 🛡️ Safety & Backup
- `excel_create_backup` - Create file auto-backup
- `excel_restore_backup` - Restore from backup files
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

**Dual-row header system** (Game dev specialized, auto-detected):
```
Row 1 (Description): ['Skill ID Desc', 'Skill Name Desc', 'Skill Type Desc']
Row 2 (Field):      ['skill_id', 'skill_name', 'skill_type']
```

`excel_query` auto-detects dual-row header format (Row 1 Chinese description + Row 2 English field name). Query results include `column_descriptions` mapping for easy understanding.

**Common config table structures**:
- **Skill Config**: ID|Name|Type|Level|Cost|Cooldown|Damage|Description
- **Equipment Config**: ID|Name|Type|Quality|Attributes|Set|Acquisition
- **Monster Config**: ID|Name|Level|HP|Attack|Defense|Skills|Drops

### 📋 Standard Workflow

1. **Search & Locate**: Use `excel_search` to understand data distribution
2. **Determine Boundaries**: Use `excel_find_last_row` to confirm data range
3. **Read Current State**: Use `excel_get_range` to get current configuration
4. **Update Data**: Use `excel_update_range` for safe updates
5. **Beautify Display**: Use `excel_format_cells` to mark important data
6. **Verify Results**: Re-read to confirm successful updates

### 🔍 SQL Query Reference

`excel_query` is built on sqlglot + pandas, supporting the following SQL syntax:

**Supported Syntax:**
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
SELECT * FROM SkillTable WHERE skill_name LIKE '%Fire%'

-- DISTINCT dedup
SELECT DISTINCT skill_type FROM SkillTable

-- IN condition
SELECT * FROM SkillTable WHERE skill_type IN ('Attack', 'Support')

-- BETWEEN range
SELECT * FROM MonsterTable WHERE level BETWEEN 10 AND 20

-- IS NULL / IS NOT NULL
SELECT * FROM SkillTable WHERE description IS NULL
SELECT * FROM SkillTable WHERE description IS NOT NULL

-- OFFSET pagination
SELECT * FROM MonsterTable ORDER BY level LIMIT 20 OFFSET 0

-- NOT LIKE / NOT IN
SELECT * FROM SkillTable WHERE skill_name NOT LIKE '%test%'
SELECT * FROM EquipTable WHERE quality NOT IN ('deprecated', 'internal_test')

-- JOIN cross-table queries (within same file)
SELECT a.skill_name, b.equip_name FROM SkillTable a INNER JOIN EquipTable b ON a.equip_id = b.equip_id
SELECT a.name, b.hp FROM MonsterTable a LEFT JOIN MonsterDropTable b ON a.id = b.monster_id WHERE a.level > 10

-- Subqueries (WHERE col IN / NOT IN / Scalar)
SELECT * FROM SkillTable WHERE skill_type IN (SELECT DISTINCT skill_type FROM SkillTable WHERE damage > 200)
SELECT * FROM SkillTable WHERE damage > (SELECT AVG(damage) FROM SkillTable)

-- EXISTS subqueries (correlated)
SELECT * FROM MonsterTable WHERE EXISTS (SELECT 1 FROM DropTable WHERE DropTable.monster_id = MonsterTable.id)

-- CASE WHEN conditional expressions
SELECT skill_name, CASE WHEN damage > 200 THEN 'High' WHEN damage > 100 THEN 'Mid' ELSE 'Low' END as tier FROM SkillTable

-- CTE (WITH ... AS ...)
WITH high_dmg AS (SELECT * FROM SkillTable WHERE damage > 150) SELECT * FROM high_dmg ORDER BY damage DESC

-- UNION / UNION ALL combine query results
SELECT name, damage FROM SkillTable WHERE skill_type='Mage' UNION ALL SELECT name, damage FROM SkillTable WHERE skill_type='Warrior' ORDER BY damage DESC LIMIT 10

-- COALESCE / IFNULL null replacement
SELECT skill_name, COALESCE(description, 'N/A') as desc FROM SkillTable

-- String functions
SELECT UPPER(skill_name), LOWER(skill_type) FROM SkillTable
SELECT CONCAT(skill_type, ':', skill_name) as label FROM SkillTable
SELECT REPLACE(description, 'attack', 'strike') FROM SkillTable
```

**SQL UPDATE Batch Modification:**
```sql
-- Constant modification
UPDATE SkillTable SET damage = 500 WHERE skill_type = 'Ultimate'

-- Arithmetic expressions (column references)
UPDATE SkillTable SET damage = damage * 1.1 WHERE element = 'Fire'

-- Multi-column modification
UPDATE SkillTable SET damage = damage * 1.1, cooldown = cooldown - 1 WHERE level >= 20

-- dry_run preview mode (no actual changes)
UPDATE SkillTable SET damage = damage * 1.1 WHERE element = 'Fire'  -- dry_run=True
```

**Unsupported Syntax (with clear alternative suggestions):**
- FROM subqueries `FROM (SELECT ...)` (suggest: use WHERE subqueries or CTEs)

**Window Functions (ROW_NUMBER/RANK/DENSE_RANK):**
```sql
-- Rank by damage descending
SELECT skill_name, damage, ROW_NUMBER() OVER (ORDER BY damage DESC) as rn FROM SkillConfig

-- Rank within each class
SELECT skill_name, skill_type, ROW_NUMBER() OVER (PARTITION BY skill_type ORDER BY damage DESC) as rn FROM SkillConfig

-- RANK vs DENSE_RANK: tied values get same rank
SELECT skill_name, damage, RANK() OVER (ORDER BY damage DESC) as r, DENSE_RANK() OVER (ORDER BY damage DESC) as dr FROM SkillConfig
```

**Query Performance:**
- Same-file repeated queries auto-cache, 30-100x speedup
- python-calamine Rust engine: get_range from 1.6s to 0.7ms (2300x speedup)
- SQL cold query: ~10ms (calamine) vs ~200ms (openpyxl)
- Cache auto-invalidates on file modification

**Common Problem Solutions**:
- **File locked**: Close Excel program and retry
- **Encoding issues**: Ensure UTF-8 encoding
- **Large file slow**: Use precise ranges, process in batches
- **Memory insufficient**: Reduce single processing amount, close workbooks promptly
- **Permission issues**: Use admin privileges or check file properties

---

## 🔒 Security Mechanisms

ExcelMCP includes multi-layer security protections:

### Path Security (SecurityValidator)
- **Path traversal protection**: Rejects `../` directory traversal attacks
- **Symlink rejection**: Does not follow symlinks to prevent pointing to sensitive files
- **Hidden file rejection**: Does not process files starting with `.`
- **Extension whitelist**: Only allows `.xlsx`/`.xlsm`/`.xls`/`.csv`/`.json`/`.bak`
- **File size limit**: Maximum 50MB per file

### Formula Injection Protection
- **DDE detection**: Rejects `=DDE()` dynamic data exchange formulas
- **CMD detection**: Rejects `=CMD()` system command execution
- **SHELL detection**: Rejects `=SHELL()` shell command formulas
- **REGISTER detection**: Rejects `=REGISTER()` DLL registration formulas
- **Pipe detection**: Rejects formulas containing pipe characters

### Data Security
- **File locking**: `excel_update_query` uses file locks (fcntl LOCK_EX) to prevent concurrent write conflicts
- **Transaction protection**: Auto-creates backup before UPDATE, auto-rollback on failure
- **Temp file cleanup**: Auto-cleans orphaned `.bak` temp files older than 1 hour on startup

### Error Messages
- Security errors are prefixed with 🔒, including specific rejection reasons
- Example: `🔒 Security validation failed: Path contains illegal characters '..'`

---

## 🏗️ Technical Architecture

### Package Structure
```
src/excel_mcp_server_fastmcp/    # Main package (directly importable after pip install)
├── __init__.py                   # Package entry point, exposes main()
├── server.py                     # MCP interface layer (46 tool definitions)
├── api/                          # API business logic layer
│   ├── excel_operations.py       # Excel operations unified entry
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
    ├── formula_cache.py          # Formula caching
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
- **Pure Delegation Pattern**: Interface layer has zero business logic, fully delegates
- **Centralized Processing**: Unified validation, error handling, result formatting
- **1-Based Indexing**: Matches Excel user habits (Row 1 = First row)
- **Workbook Caching**: 75% performance improvement when cache hits
- **Realistic Concurrency Handling**: Properly handles Excel file concurrency limitations

### Performance Optimization
- **python-calamine Read Engine**: Rust native parsing, get_range from 1.6s to 0.7ms (2300x speedup)
- **Precise Range Reading**: 60-80% faster than reading entire tables
- **Batch Operations**: 15-20x faster than individual operations
- **Batch Processing**: 70% memory usage reduction for large files

---

## 📊 Project Information

### Quality Validation Metrics
- **Test Cases**: 798 (behavior validation, no coverage padding)
- **Test Files**: 34 test files
- **Test Code**: 13,574 lines
- **Tool Count: 46 (@mcp.tool decorator verified)
- **Architecture Layers**: 4-layer design (MCP→API→Core→Utils)

### Verification Commands
```bash
# Run complete test suite (parallel accelerated)
python -m pytest tests/ -q --tb=short -n auto --timeout=30

# Verify tool completeness
grep -c "def excel_" src/excel_mcp_server_fastmcp/server.py  # Should output: 46

# Generate coverage report
python -m pytest tests/ --cov=src --cov-report=html
```

### Development Standards
- **Pure Delegation Pattern**: server.py strictly delegates to ExcelOperations
- **Centralized Business Logic**: Unified validation, error handling, result formatting
- **Branch Naming**: All feature branches must start with `feature/`
- **Test Coverage**: Maintain 80%+ test coverage

---

## ❓ Frequently Asked Questions

### Basic Questions
**Q: Which Excel formats are supported?**
A: Supports `.xlsx`, `.xlsm` formats, with `.csv` support through import/export

**Q: How to handle Chinese worksheet names?**
A: Fully supports Chinese worksheet names and content

**Q: How is large file processing performance?**
A: SQL queries auto-cache DataFrame, repeated queries on the same file are 30-100x faster. Large table (2000 rows): first ~230ms, cached 2-8ms.

**Q: How to ensure data security?**
A: Complete error handling, formula preservation by default, operation preview support

### Game Development Specialized
**Q: What is the dual-row header system?**
A: Game config table standard format: Row 1 field descriptions, Row 2 field names

**Q: How to perform version comparison?**
A: Use specialized config table comparison tools with ID object tracking

---

## 🤝 Contributing Guide

**Contribution Methods**:
- 🐛 **Report Bugs**: Report issues through GitHub Issues
- 💡 **Feature Suggestions**: Propose new feature requirements
- 📝 **Documentation Improvements**: Improve usage guides and technical documentation
- 🔧 **Code Contributions**: Follow development standards, submit PRs

**License**: MIT License - See [LICENSE](LICENSE) file for details

---

<div align="center">

## 🔝 Quick Navigation

| 🎯 **Quick Start** | 🛠️ **Tool Reference** | 📚 **Learning Guide** |
|-------------------|------------------------|---------------------|
| [🚀 Installation](#-quick-start) | [📋 Complete Tool List](#️-complete-tool-list-41-professional-tools) | [📖 Usage Guide](#-usage-guide) |
| [⚡ Command Cheat Sheet](#-quick-reference) | [🏗️ Technical Architecture](#️-technical-architecture) | [🚨 Troubleshooting](#-troubleshooting) |
| [🎮 Game Config Management](#-usage-guide) | [📊 Project Info](#-project-information) | [❓ FAQ](#-frequently-asked-questions) |

**[⬆️ Back to Top](#-excelmcp-game-dev-excel-configuration-table-manager)**

*✨ Making game configuration table management simple and efficient ✨*

</div>
