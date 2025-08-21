
<div align="center">
<a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a>
</div>

# SheetPilot: 您的 AI 驱动的 Excel 副驾驶 🚀

[![许可证: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 版本](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![技术支持: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/your-fastmcp-repo)
[![状态](https://img.shields.io/badge/status-active-success.svg)]()

**SheetPilot** 是一个强大的模型上下文协议 (MCP) 服务器，它改变了您与 Excel 电子表格的交互方式。告别复杂的公式和手动数据整理。借助 SheetPilot，您可以使用简单的自然语言命令来管理、查询和自动化您的 Excel 工作流。这就像为您的所有电子表格任务配备了一位 AI 副驾驶。

---

### ✨ 主要功能

*   ⚡️ **闪电般快速的搜索**: 使用强大的正则表达式搜索，即时在数千个单元格和文件中查找数据。
*   📊 **轻松的数据管理**: 通过简单的指令读取、写入和更新单元格范围、行和列。
*   🗂️ **完全的工作区控制**: 即时创建、删除和管理 Excel 文件和工作表。
*   🎨 **动态格式化**: 使用预设或自定义样式，为您的数据应用美观、一致的格式。
*   🔍 **目录范围的操作**: 一次性在整个 Excel 文件文件夹上运行命令，实现真正的自动化。
*   🔒 **稳健可靠**: 内置集中式错误处理系统，性能稳定可预测。

---

### 🎬 快速演示

*（此处可以插入一个 GIF，展示用户输入“在 `report.xlsx` 中查找所有电子邮件并用黄色突出显示”，然后服务器执行该命令）*

**示例提示:**
```
"在 `quarterly_sales.xlsx` 中，查找‘地区’为‘北部’且‘销售额’超过 5000 的所有行。将它们复制到一个名为‘Top Performers’的新工作表中，并将标题格式设置为蓝色。"
```

---

### 🚀 快速入门 (5 分钟设置)

只需几个步骤，即可在您喜欢的 MCP 客户端（如 VS Code、Cursor 或 Claude Desktop）中运行 SheetPilot。

**先决条件:**
*   Python 3.8+
*   一个与 MCP 兼容的客户端。

**安装:**

1.  **克隆存储库:**
    ```bash
    git clone https://github.com/your-username/sheet-pilot.git
    cd sheet-pilot
    ```

2.  **安装依赖项:**
    我们建议使用 `uv` 进行快速安装。
    ```bash
    pip install uv
    uv pip install -r requirements.txt
    ```

3.  **配置您的 MCP 客户端:**
    将以下配置添加到您客户端的 MCP 设置文件中 (例如 `.vscode/mcp.json`, `.cursor/mcp.json`):

    ```json
    {
      "mcpServers": {
        "sheetpilot": {
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
    *请确保 `PYTHONPATH` 指向项目目录的根目录。*

4.  **开始自动化！**
    一切就绪！开始向您的 AI 助手发出自然语言命令来控制 Excel。

---

### 🛠️ 可用工具

SheetPilot 向您的 AI 助手公开了一套丰富的工具集:

| 工具名称                       | 描述                                                                 |
| ------------------------------ | -------------------------------------------------------------------- |
| `excel_list_sheets`            | 列出给定 Excel 文件中的所有工作表名称。                              |
| `excel_regex_search`           | 在单个文件中搜索与正则表达式模式匹配的内容。                         |
| `excel_regex_search_directory` | 在指定目录中的所有 Excel 文件中搜索内容。                            |
| `excel_get_range`              | 读取并返回指定范围（例如 "A1:C10"）的数据。                          |
| `excel_update_range`           | 使用新数据更新指定范围。                                             |
| `excel_insert_rows`            | 在给定位置插入指定数量的空行。                                       |
| `excel_insert_columns`         | 在给定位置插入指定数量的空列。                                       |
| `excel_delete_rows`            | 从给定位置删除指定数量的行。                                         |
| `excel_delete_columns`         | 从给定位置删除指定数量的列。                                         |
| `excel_create_file`            | 创建一个新的、空的 `.xlsx` 文件，可选择带有命名的工作表。            |
| `excel_create_sheet`           | 向现有文件添加新工作表。                                             |
| `excel_delete_sheet`           | 从文件中删除工作表。                                                 |
| `excel_rename_sheet`           | 重命名现有工作表。                                                   |
| `excel_format_cells`           | 将样式（字体、颜色、对齐方式）应用于单元格范围。                     |

---

### 💡 用例

*   **数据清理**: "在 `/reports` 目录中的所有 `.xlsx` 文件中，查找包含 `N/A` 的单元格，并将其替换为空值。"
*   **自动报告**: "创建一个新文件 `summary.xlsx`。将 `sales_data.xlsx` 中的范围 `A1:F20` 复制到名为‘Sales’的工作表中，并将 `inventory.xlsx` 中的 `A1:D15` 复制到名为‘Inventory’的工作表中。"
*   **数据提取**: "获取 `contacts.xlsx` 中 A 列为‘Active’的所有 D 列的值。"
*   **批量格式化**: "在 `financials.xlsx` 中，将整个第一行加粗，并将其背景颜色设置为浅灰色。"

---

### 🤝 贡献

欢迎贡献！无论是添加新功能、改进文档还是报告错误，我们都希望得到您的帮助。请查看我们的 `CONTRIBUTING.md` 以获取有关如何开始的更多详细信息。

### 📜 许可证

该项目根据 MIT 许可证授权。有关详细信息，请参阅 [LICENSE](LICENSE) 文件。
