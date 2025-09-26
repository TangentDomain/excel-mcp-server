
<div align="center">

[简体中文](README.md) ｜ [English](README.en.md)

</div>

# ExcelMCP: Powerful Excel MCP Server 🚀

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Powered by: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
![Status](https://img.shields.io/badge/status-production-success.svg)
![Tests](https://img.shields.io/badge/tests-295%20passed-brightgreen.svg)

**ExcelMCP** is a comprehensive Model Context Protocol (MCP) server that revolutionizes Excel file manipulation through AI. Built with **FastMCP** and **openpyxl**, it provides 32+ powerful tools enabling AI assistants to perform complex Excel operations through natural language commands. From regex searches across thousands of files to advanced data manipulation and formatting - all with enterprise-grade reliability.

🎯 **Perfect for:** Game development configuration tables, data analysis workflows, automated reporting, bulk file processing, and intelligent office automation.

---

## ✨ Key Features

- ⚡️ **32+ Advanced Tools**: Complete Excel manipulation suite from basic CRUD to complex formatting
- 🔍 **Powerful Search Engine**: Regex search across files with directory-wide operations
- 🧠 **Smart Data Operations**: Range-based read/write, row/column management, formula preservation
- 🎨 **Professional Formatting**: Preset styles, custom formatting, borders, merging, sizing
- 🗂️ **File Lifecycle Management**: Create, convert, merge, import/export CSV, file information
- 🎮 **Game Development Optimized**: Specialized Excel config table comparison for game development
- 🔒 **Enterprise-Ready**: Centralized error handling, comprehensive validation, 100% test coverage

---

### 🎬 Quick Demo

*(Here you could insert a GIF showing a user typing "Find all emails in `report.xlsx` and highlight them in yellow" and the server executing it)*

**Example Prompt:**

```text
"In `quarterly_sales.xlsx`, find all rows where the 'Region' is 'North' and the 'Sale Amount' is over 5000. Copy them to a new sheet named 'Top Performers' and format the header in blue."
```

---

## 🚀 Getting Started (3-Minute Setup)

Get ExcelMCP running in your favorite MCP client (VS Code with Continue, Cursor, Claude Desktop, or any MCP-compatible client).

### Prerequisites

- Python 3.10+
- An MCP-compatible client

### Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/tangjian/excel-mcp-server.git
    cd excel-mcp-server
    ```

2. **Install dependencies:**

    Using **uv** (recommended for speed):

    ```bash
    pip install uv
    uv sync
    ```

    Or using **pip**:

    ```bash
    pip install -e .
    ```

3. **Configure your MCP client:**

    Add to your MCP client configuration (`.vscode/mcp.json`, `.cursor/mcp.json`, etc.):

    ```json
    {
      "mcpServers": {
        "excelmcp": {
          "command": "python",
          "args": ["-m", "src.server"],
          "env": {
            "PYTHONPATH": "${workspaceRoot}"
          }
        }
      }
    }
    ```

4. **Start automating!**

    You're ready! Ask your AI assistant to control Excel files with natural language.

---

## 🛠️ Available Tools (32 Advanced Excel Operations)

ExcelMCP provides a comprehensive suite of Excel manipulation tools:

### 📋 File & Sheet Management

| Tool | Description |
|------|------------|
| `excel_list_sheets` | Lists all worksheet names in an Excel file |
| `excel_create_file` | Creates new Excel files with optional named sheets |
| `excel_create_sheet` | Adds new worksheets to existing files |
| `excel_delete_sheet` | Removes worksheets from files |
| `excel_rename_sheet` | Renames existing worksheets |
| `excel_get_file_info` | Retrieves detailed file information (size, format, etc.) |

### 🔍 Search & Data Discovery

| Tool | Description |
|------|------------|
| `excel_search` | Regex search within single Excel files with range support |
| `excel_search_directory` | Batch regex search across entire directories |
| `excel_get_range` | Reads data from specified ranges (cells/rows/columns) |
| `excel_get_headers` | Extracts column headers from worksheets |
| `excel_get_sheet_headers` | Gets headers from all worksheets in a file |

### ✏️ Data Manipulation

| Tool | Description |
|------|------------|
| `excel_update_range` | Updates cell ranges with new data and formula preservation |
| `excel_insert_rows` | Inserts empty rows at specified positions |
| `excel_insert_columns` | Inserts empty columns at specified positions |
| `excel_delete_rows` | Removes rows from worksheets |
| `excel_delete_columns` | Removes columns from worksheets |

### 🎨 Formatting & Styling

| Tool | Description |
|------|------------|
| `excel_format_cells` | Applies fonts, colors, alignment with presets or custom styles |
| `excel_merge_cells` | Merges cell ranges for headers and layouts |
| `excel_unmerge_cells` | Unmerges previously merged cell ranges |
| `excel_set_borders` | Adds borders with various styles (thin, thick, dotted, etc.) |
| `excel_set_row_height` | Adjusts row heights in points |
| `excel_set_column_width` | Adjusts column widths in character units |

### 🔄 Import/Export & Conversion

| Tool | Description |
|------|------------|
| `excel_export_to_csv` | Exports worksheets to CSV with encoding options |
| `excel_import_from_csv` | Creates Excel files from CSV data |
| `excel_convert_format` | Converts between Excel formats (xlsx, xlsm, csv, json) |
| `excel_merge_files` | Combines multiple Excel files with different merge modes |
| `excel_compare_sheets` | Compares Excel sheets to identify differences (game dev optimized) |

All tools return structured JSON responses with success indicators, detailed results, and comprehensive error information.

---

## 💡 Use Cases & Examples

### Real-World Applications

- **Game Development**: "Compare TrSkill.xlsx configuration tables between versions and highlight changes in damage values"
- **Data Cleaning**: "In all `.xlsx` files in `/reports`, find cells containing 'N/A' and replace with empty values"
- **Automated Reporting**: "Create summary.xlsx with Sales sheet (copy A1:F20 from sales_data.xlsx) and Inventory sheet (copy A1:D15 from inventory.xlsx)"
- **Bulk Processing**: "Search all Excel files in directory for email patterns and export matches to emails.csv"
- **Professional Formatting**: "Apply company header style to A1:E1 range with blue background and white bold text"

### Command Examples

```plaintext
Natural Language → AI Assistant → ExcelMCP

"Find all cells containing currency symbols in my finance folder"
→ Uses excel_search_directory with regex pattern [$€¥£]

"Create a new report with three sheets: Data, Charts, Summary"
→ Uses excel_create_file with custom sheet names

"Make the header row bold and add borders to the data table"
→ Uses excel_format_cells with preset="header" + excel_set_borders

"Compare Q3 and Q4 sales sheets and show me what changed"
→ Uses excel_compare_sheets to identify differences
```

---

## 🏗️ Architecture & Dependencies

### Core Technologies

- **[FastMCP](https://github.com/jlowin/fastmcp)**: Modern MCP server framework
- **[openpyxl](https://openpyxl.readthedocs.io/)**: Core Excel file manipulation
- **[xlcalculator](https://pypi.org/project/xlcalculator/)**: Formula evaluation engine
- **[xlwings](https://www.xlwings.org/)**: Optional Excel application integration

### Project Structure

```text
src/
├── server.py              # MCP tool definitions (pure delegation)
├── api/excel_operations.py # Centralized business logic
├── core/                  # Excel operation modules
│   ├── excel_reader.py    # Read operations
│   ├── excel_writer.py    # Write operations
│   ├── excel_manager.py   # File/sheet management
│   └── excel_search.py    # Search & comparison
├── utils/                 # Validators, parsers, formatters
└── models/                # Type definitions
```

### Quality Assurance

- **295 comprehensive tests** with 100% passing rate
- **Centralized error handling** with structured responses
- **Type safety** with full TypeScript-style annotations
- **Game development optimized** with specialized config table tools

---

## 🤝 Contributing

We welcome contributions! Whether it's adding new features, improving documentation, or reporting bugs:

1. **Fork the repository** and create your feature branch
2. **Add tests** for any new functionality (maintain our 100% pass rate!)
3. **Follow code style**: Use type hints, docstrings, and our error handling patterns
4. **Submit a PR** with clear description of changes

### Development Setup

```bash
git clone https://github.com/tangjian/excel-mcp-server.git
cd excel-mcp-server
uv sync --dev  # Install with development dependencies
pytest tests/ # Run the full test suite (221 tests)
```

---

## 📜 License

Licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

---

## ⭐ Support

If ExcelMCP helps your workflow, please:

- ⭐ Star this repository
- 🐛 Report issues on GitHub
- 💡 Suggest new features
- 📖 Contribute to documentation

Built with ❤️ for the AI and Excel automation community.
