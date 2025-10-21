<div align="center">

[简体中文](README.md) ｜ [English](README.en.md)

</div>

# 🎮 ExcelMCP: Game Development Excel Configuration Table Manager

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Powered by: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![Status](https://img.shields.io/badge/status-stable-green.svg)
![Tests](https://img.shields.io/badge/tests-698%20passed-brightgreen.svg)
![Coverage](https://img.shields.io/badge/coverage-78.58%25-blue.svg)
![Tools](https://img.shields.io/badge/tools-38%20verified%20tools-green.svg)

**ExcelMCP** is an Excel configuration table management MCP server specially designed for game development. Through AI natural language commands, it enables intelligent operations on game configurations such as skill tables, equipment data, and monster attributes. Built with **FastMCP** and **openpyxl**, it features **38 professional tools** and **698 test cases**, ensuring enterprise-grade reliability.

🎯 **Core Features**: Skill systems, equipment management, monster configuration, numerical balancing, version comparison, designer toolchain

---

## 🚀 Quick Start (3-Minute Setup)

### Installation Steps

1. **Clone the project**
   ```bash
   git clone https://github.com/tangjian/excel-mcp-server.git
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
         "args": ["-m", "src.server"],
         "env": {"PYTHONPATH": "${workspaceRoot}"}
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
| Equipment configuration management | `excel_format_cells` + `excel_get_range` | "Mark all legendary equipment with gold color" |
| Monster data validation | `excel_check_duplicate_ids` + `excel_search` | "Ensure monster IDs are unique and HP is reasonable" |
| Version comparison analysis | `excel_compare_sheets` | "Compare differences between old and new version config tables" |

### 🔧 Range Expression Reference

| Format | Description | Example |
|--------|-------------|---------|
| `Sheet1!A1:C10` | Standard range | "SkillTable!A1:D50" |
| `Sheet1!1:5` | Row range | "ConfigTable!2:100" |
| `Sheet1!B:D` | Column range | "DataTable!B:G" |
| `Sheet1!A1` | Single cell | "SettingsTable!A1" |

---

## 🛠️ Complete Tool List (38 Professional Tools)

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

### 🔍 Search & Analysis
- `excel_search` - Regex expression search
- `excel_search_directory` - Directory batch search
- `excel_compare_sheets` - Worksheet comparison (game config optimized)
- `excel_check_duplicate_ids` - ID duplicate detection

### 🎨 Formatting & Styling
- `excel_format_cells` - Apply fonts, colors, alignment formats
- `excel_set_borders` - Set cell borders
- `excel_merge_cells` - Merge cell ranges
- `excel_unmerge_cells` - Unmerge cells
- `excel_set_column_width` - Adjust column width
- `excel_set_row_height` - Adjust row height

### 🔄 Data Conversion
- `excel_export_to_csv` - Export CSV format
- `excel_import_from_csv` - Create Excel files from CSV
- `excel_convert_format` - Format conversion (.xlsx/.xlsm/.csv/.json)

---

## 📖 Usage Guide

### 🎮 Game Configuration Table Standard Format

**Dual-row header system** (Game development specialized):
```
Row 1 (Description): ['Skill ID Description', 'Skill Name Description', 'Skill Type Description']
Row 2 (Field):      ['skill_id', 'skill_name', 'skill_type']
```

**Common configuration table structures**:
- **Skill Configuration Table**: ID|Name|Type|Level|Cost|Cooldown|Damage|Description
- **Equipment Configuration Table**: ID|Name|Type|Quality|Attributes|Set|Acquisition
- **Monster Configuration Table**: ID|Name|Level|HP|Attack|Defense|Skills|Drops

### 📋 Standard Workflow

1. **Search & Locate**: Use `excel_search` to understand data distribution
2. **Determine Boundaries**: Use `excel_find_last_row` to confirm data range
3. **Read Current State**: Use `excel_get_range` to get current configuration
4. **Update Data**: Use `excel_update_range` for safe updates
5. **Beautify Display**: Use `excel_format_cells` to mark important data
6. **Verify Results**: Re-read to confirm successful updates

### 🚨 Troubleshooting

**Common Problem Solutions**:
- **File locked**: Close Excel program and retry
- **Chinese garbled**: Ensure UTF-8 encoding, check Python environment encoding
- **Large file slow**: Use precise ranges, process data in batches
- **Memory insufficient**: Reduce single processing data amount, close workbooks promptly
- **Permission issues**: Use administrator privileges or check file properties

---

## 🏗️ Technical Architecture

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
- **Precise Range Reading**: 60-80% faster than reading entire tables
- **Batch Operations**: 15-20x faster than individual operations
- **Batch Processing**: 70% memory usage reduction for large files

---

## 📊 Project Information

### Quality Validation Metrics
- **Test Cases**: 699 (698 passed, 1 skipped)
- **Test Code**: 13,515 lines (comprehensive validation)
- **Tool Count**: 38 (verified with @mcp.tool decorators)
- **Test Coverage**: 78.58%
- **Architecture Layers**: 4-layer design (MCP→API→Core→Utils)

### Verification Commands
```bash
# Run complete test suite
python -m pytest tests/ -v

# Verify tool completeness
grep -r "@mcp.tool" src/ | wc -l  # Should output: 38

# Generate coverage report
python -m pytest tests/ --cov=src --cov-report=html
```

### Development Standards
- **Pure Delegation Pattern**: server.py strictly delegates to ExcelOperations
- **Centralized Business Logic**: Unified validation, error handling, result formatting
- **Branch Naming**: All feature branches must start with `feature/`
- **Test Coverage**: Maintain 78%+ test coverage

---

## ❓ Frequently Asked Questions

### Basic Questions
**Q: Which Excel formats are supported?**
A: Supports `.xlsx`, `.xlsm` formats, with CSV support through import/export

**Q: How to handle Chinese worksheet names?**
A: Fully supports Chinese worksheet names and content

**Q: How is large file processing performance?**
A: Based on openpyxl performance, recommends batch processing for large files

**Q: How to ensure data security?**
A: Complete error handling, formula preservation by default, operation preview support

### Game Development Specialized
**Q: What is the dual-row header system?**
A: Game configuration table standard format: Row 1 field descriptions, Row 2 field names

**Q: How to perform version comparison?**
A: Use specialized configuration table comparison tools with ID object tracking

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
| [🚀 Installation](#-quick-start-3-minute-setup) | [📋 Complete Tool List](#️-complete-tool-list-38-professional-tools) | [📖 Usage Guide](#-usage-guide) |
| [⚡ Command Cheat Sheet](#-quick-reference) | [🏗️ Technical Architecture](#️-technical-architecture) | [🚨 Troubleshooting](#-troubleshooting) |
| [🎮 Game Config Management](#-usage-guide) | [📊 Project Info](#-project-information) | [❓ FAQ](#-frequently-asked-questions) |

**[⬆️ Back to Top](#-excelmcp-game-development-excel-configuration-table-manager)**

*✨ Making game configuration table management simple and efficient ✨*

</div>