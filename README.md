
<div align="center">
<a href="README.md">### 🎬 游戏开发场景演示

**技能配置表管理:**

```bash
"在技能配置表中查找所有火系技能，将伤害值统一提升20%，并用红色高亮显示"
```

**装备数据分析:**

```bash
"比较新旧版本的装备配置表，找出所有属性变更的装备，生成详细的变更报告"
```

**游戏数值平衡:**

```bash
"检查怪物配置表中所有等级20-30的怪物，确保血量和攻击力的比例在合理范围内"
```| <a href="README.en.md">English</a>
</div>

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![许可证: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 版本](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![技术支持: FastMCP](https://img.shields.io/badge/Powered%20by-FastMCP-orange)](https://github.com/jlowin/fastmcp)
[![状态](https://img.shields.io/badge/status-production-success.svg)](#)
[![测试覆盖](https://img.shields.io/badge/tests-295%20passed-brightgreen.svg)](#)
[![工具数量](https://img.shields.io/badge/tools-32个专业工具-blue.svg)](#)

**ExcelMCP** 是专为**游戏开发**设计的 Excel 配置表管理 MCP (Model Context Protocol) 服务器。通过 AI 自然语言指令，实现技能配置表、装备数据、怪物属性等游戏配置的智能化操作。基于 **FastMCP** 和 **openpyxl** 构建，拥有 **32个专业工具** 和 **295个测试用例**，确保企业级可靠性。

🎯 **专为游戏开发而生：** 技能系统、装备管理、怪物配置、游戏数值平衡、版本对比、策划工具链。

---

## 🎮 游戏配置表专业管理

### 🎯 游戏开发核心功能
- **🛡️ 技能配置表** (TrSkill): 技能ID、名称、类型、等级、消耗、冷却、伤害、描述
- **⚔️ 装备配置表** (TrItem): 装备ID、名称、类型、品质、属性、套装、获取方式
- **👹 怪物配置表** (TrMonster): 怪物ID、名称、等级、血量、攻击、防御、技能、掉落
- **🎁 道具配置表** (TrProps): 道具ID、名称、类型、数量、效果、获取、描述

### ✨ 32个专业工具分类
- **⚡️ 文件管理** (8工具): 创建、转换、合并、导入导出、文件信息
- **🔍 智能搜索** (4工具): 正则搜索、目录搜索、游戏表头分析、重复ID检测
- **📊 数据操作** (8工具): 范围读写、行列管理、数据更新、公式保护
- **🎨 格式美化** (6工具): 预设样式、边框设置、合并单元格、行列尺寸
- **🔄 版本对比** (3工具): 游戏配置表比较、对象变更跟踪、ID对象管理
- **🛠️ 高级功能** (3工具): 中文支持、Unicode处理、批量操作

### 🔒 企业级可靠性
- **295个测试用例** 100% 通过，覆盖所有功能模块
- **委托架构设计** 纯MCP接口 → API层 → 核心层 → 工具层
- **统一错误处理** 集中化异常管理和用户友好提示
- **性能优化** 工作簿缓存、内存管理、大文件处理

---

### 🎬 快速演示

*（此处可以插入一个 GIF，展示用户输入“在 `report.xlsx` 中查找所有电子邮件并用黄色突出显示”，然后服务器执行该命令）*

**示例提示:**

```
"在 `quarterly_sales.xlsx` 中，查找‘地区’为‘北部’且‘销售额’超过 5000 的所有行。将它们复制到一个名为‘Top Performers’的新工作表中，并将标题格式设置为蓝色。"
```

---

## 🚀 快速入门 (3 分钟设置)

在您喜欢的 MCP 客户端（VS Code 配 Continue、Cursor、Claude Desktop 或任何 MCP 兼容客户端）中运行 ExcelMCP。

### 先决条件

- Python 3.10+
- 一个与 MCP 兼容的客户端

### 安装

1. **克隆存储库:**

    ```bash
    git clone https://github.com/tangjian/excel-mcp-server.git
    cd excel-mcp-server
    ```

2. **安装依赖项:**

    使用 **uv**（推荐，速度更快）:

    ```bash
    pip install uv
    uv sync
    ```

    或使用 **pip**:

    ```bash
    pip install -e .
    ```

3. **配置您的 MCP 客户端:**

    添加到您的 MCP 客户端配置中（`.vscode/mcp.json`、`.cursor/mcp.json` 等）:

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

4. **开始自动化！**

    准备就绪！让您的 AI 助手通过自然语言控制 Excel 文件。

---

## � 完整工具列表（27个工具）

### 📁 文件与工作表管理

| 工具 | 用途 |
|------|------|
| `excel_create_file` | 创建新的 Excel 文件（.xlsx/.xlsm），支持自定义工作表 |
| `excel_create_sheet` | 在现有文件中添加新工作表 |
| `excel_delete_sheet` | 删除指定工作表 |
| `excel_list_sheets` | 列出工作表名称和获取文件信息 |
| `excel_rename_sheet` | 重命名工作表 |
| `excel_get_file_info` | 获取文件元数据（大小、创建日期等） |

### 📊 数据操作

| 工具 | 用途 |
|------|------|
| `excel_get_range` | 读取单元格/行/列范围（支持 A1:C10、行范围、列范围等） |
| `excel_update_range` | 写入/更新数据范围，支持公式保留 |
| `excel_get_headers` | 从任意行提取表头 |
| `excel_get_sheet_headers` | 获取所有工作表的表头 |
| `excel_insert_rows` | 插入空行到指定位置 |
| `excel_delete_rows` | 删除行范围 |
| `excel_insert_columns` | 插入空列到指定位置 |
| `excel_delete_columns` | 删除列范围 |

### 🔍 搜索与分析

| 工具 | 用途 |
|------|------|
| `excel_search` | 在工作表中进行正则表达式搜索 |
| `excel_search_directory` | 在目录中的所有 Excel 文件中批量搜索 |
| `excel_compare_sheets` | 比较两个工作表，检测变化（针对游戏配置优化） |

### 🎨 格式化与样式

| 工具 | 用途 |
|------|------|
| `excel_format_cells` | 应用字体、颜色、对齐等格式（预设或自定义） |
| `excel_set_borders` | 设置单元格边框样式 |
| `excel_merge_cells` | 合并单元格范围 |
| `excel_unmerge_cells` | 取消合并单元格 |
| `excel_set_column_width` | 调整列宽 |
| `excel_set_row_height` | 调整行高 |

### 🔄 数据转换

| 工具 | 用途 |
|------|------|
| `excel_export_to_csv` | 导出工作表为 CSV 格式 |
| `excel_import_from_csv` | 从 CSV 创建 Excel 文件 |
| `excel_convert_format` | 在 Excel 格式间转换（.xlsx、.xlsm、.csv、.json） |
| `excel_merge_files` | 合并多个 Excel 文件 |

---

### 💡 用例

- **数据清理**: "在 `/reports` 目录中的所有 `.xlsx` 文件中，查找包含 `N/A` 的单元格，并将其替换为空值。"
- **自动报告**: "创建一个新文件 `summary.xlsx`。将 `sales_data.xlsx` 中的范围 `A1:F20` 复制到名为‘Sales’的工作表中，并将 `inventory.xlsx` 中的 `A1:D15` 复制到名为‘Inventory’的工作表中。"
- **数据提取**: "获取 `contacts.xlsx` 中 A 列为‘Active’的所有 D 列的值。"
- **批量格式化**: "在 `financials.xlsx` 中，将整个第一行加粗，并将其背景颜色设置为浅灰色。"

---

### 🤝 贡献

欢迎贡献！无论是添加新功能、改进文档还是报告错误，我们都希望得到您的帮助。请查看我们的 `CONTRIBUTING.md` 以获取有关如何开始的更多详细信息。

### 📜 许可证

该项目根据 MIT 许可证授权。有关详细信息，请参阅 [LICENSE](LICENSE) 文件。
