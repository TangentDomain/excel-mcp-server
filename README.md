

<div align="center">
<a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a>
</div>

# ExcelMCP: Powerful Excel MCP Server 🚀


[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Powered by: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/your-fastmcp-repo)
[![Status](https://img.shields.io/badge/status-active-success.svg)]()

**ExcelMCP** is a powerful Model Context Protocol (MCP) server that transforms how you interact with Excel spreadsheets. Say goodbye to complex formulas and manual data wrangling. With ExcelMCP, you can manage, query, and automate your Excel workflows using simple natural language commands. Let AI assistants directly control your Excel files for truly intelligent office automation.

---

### ✨ Key Features

*   ⚡️ **Blazing-Fast Search**: Instantly find data across thousands of cells and files with powerful regex searches.
*   📊 **Effortless Data Management**: Read, write, and update cell ranges, rows, and columns with simple instructions.
*   🗂️ **Full Workspace Control**: Create, delete, and manage Excel files and worksheets on the fly.
*   🎨 **Dynamic Formatting**: Apply beautiful, consistent formatting to your data with presets or custom styles.
*   🔍 **Directory-Wide Operations**: Run commands on an entire folder of Excel files at once for true automation.
*   🔒 **Robust & Reliable**: Built with a centralized error-handling system for stable and predictable performance.

---

### 🎬 Quick Demo

*(Here you could insert a GIF showing a user typing "Find all emails in `report.xlsx` and highlight them in yellow" and the server executing it)*

**Example Prompt:**
```
"In `quarterly_sales.xlsx`, find all rows where the 'Region' is 'North' and the 'Sale Amount' is over 5000. Copy them to a new sheet named 'Top Performers' and format the header in blue."
```

---

### 🚀 Getting Started (5-Minute Setup)

Get ExcelMCP running in your favorite MCP client (like VS Code, Cursor, or Claude Desktop) with just a few steps.

**Prerequisites:**
*   Python 3.8+
*   An MCP-compatible client.

**Installation:**

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/excel-mcp.git
    cd excel-mcp
    ```

2.  **Install dependencies:**
    We recommend using `uv` for fast installation.
    ```bash
    pip install uv
    uv pip install -r requirements.txt
    ```

3.  **Configure your MCP client:**
    Add the following configuration to your client's MCP settings file (e.g., `.vscode/mcp.json`, `.cursor/mcp.json`):

    ```json
    {
      "mcpServers": {
        "excelmcp": {
          "command": "python",
          "args": [
            "-m",
            "src.server"
          ],
          "env": {
            "PYTHONPATH": "${workspaceRoot}"
          }
        }
      }
    }
    ```
    *Make sure the `PYTHONPATH` points to the root of the project directory.*

4.  **Start Automating!**
    You're all set! Start giving natural language commands to your AI assistant to control Excel.

---

### 🛠️ Available Tools

ExcelMCP exposes a rich set of tools to your AI assistant:

| Tool Name                      | Description                                                              |
| ------------------------------ | ------------------------------------------------------------------------ |
| `excel_list_sheets`            | Lists all worksheet names in a given Excel file.                         |
| `excel_regex_search`           | Searches for content matching a regex pattern within a single file.      |
| `excel_regex_search_directory` | Searches for content across all Excel files in a specified directory.    |
| `excel_get_range`              | Reads and returns data from a specified range (e.g., "A1:C10").          |
| `excel_update_range`           | Updates a specified range with new data.                                 |
| `excel_insert_rows`            | Inserts a specified number of empty rows at a given position.            |
| `excel_insert_columns`         | Inserts a specified number of empty columns at a given position.         |
| `excel_delete_rows`            | Deletes a specified number of rows from a given position.                |
| `excel_delete_columns`         | Deletes a specified number of columns from a given position.             |
| `excel_create_file`            | Creates a new, empty `.xlsx` file, optionally with named sheets.         |
| `excel_create_sheet`           | Adds a new worksheet to an existing file.                                |
| `excel_delete_sheet`           | Deletes a worksheet from a file.                                         |
| `excel_rename_sheet`           | Renames an existing worksheet.                                           |
| `excel_format_cells`           | Applies styling (font, color, alignment) to a range of cells.            |

---

### 💡 Use Cases

*   **Data Cleaning**: "In all `.xlsx` files in the `/reports` directory, find cells containing `N/A` and replace them with an empty value."
*   **Automated Reporting**: "Create a new file `summary.xlsx`. Copy the range `A1:F20` from `sales_data.xlsx` into a sheet named 'Sales', and copy `A1:D15` from `inventory.xlsx` into a sheet named 'Inventory'."
*   **Data Extraction**: "Get all values from column D in `contacts.xlsx` where column A is 'Active'."
*   **Bulk Formatting**: "In `financials.xlsx`, make the entire first row bold and set its background color to light gray."

---

### 🤝 Contributing

Contributions are welcome! Whether it's adding new features, improving documentation, or reporting bugs, we'd love to have your help. Please check out our `CONTRIBUTING.md` for more details on how to get started.

### 📜 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
