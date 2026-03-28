[简体中文](README.md) | [English](README.en.md)

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v1.6.45)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1171-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)
![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Stars)

> **AI-Driven Excel Configuration Table Management Tool** - Use natural language or SQL to operate game configuration data, supporting cross-table JOIN, version comparison, and batch modifications

---

## 🚀 3-Minute Quick Start (Follow these steps)

### ✅ Step 1: Check Python Environment (10 seconds)

打开终端，输入：
```bash
python --version
```

See `Python 3.10+`? ✅ **Skip to Step 2**

No Python? Download from [python.org](https://www.python.org/downloads/) (Windows users: remember to check "Add Python to PATH")

### ⚡ Step 2: Install Tool (Choose one, 30 seconds)

#### 🎯 推荐方式：uvx（最快，无安装）
```bash
# Mac/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell管理员模式)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

重启终端后验证：
```bash
uvx --version
```

#### 💾 传统方式：pip（稳定）
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 **Users in China** having slow downloads? Use mirror:
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### 🔧 Step 3: Configure AI Client (1 minute)

找到你的AI客户端，按说明配置：

#### 🟢 Claude Desktop（推荐）

1. 打开配置文件：
   - **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. 添加配置（如果已有其他内容，加到 `mcpServers` 里）：
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

3. 保存文件，**重启Claude Desktop**

#### 🟡 Cursor

1. 打开设置：`Ctrl+,` 或 `Cmd+,`
2. 搜索 "mcp"，点击 "Model Context Protocol"
3. 点击 "Add MCP Server"
4. 填入：
   - **Name**: `excelmcp`
   - **Command**: `uvx`
   - **Args**: `["excel-mcp-server-fastmcp"]`

5. 重启Cursor

#### 🔴 其他客户端

- **Cherry Studio**: 设置 → MCP → 添加服务器
- **VSCode + Continue**: 设置 → MCP → 添加服务器
- **OpenClaw**: 已内置支持

### ✅ Step 4: Verify Configuration (10 seconds)

After restarting AI client, ask AI to test:
```
Please help me read an Excel file to test if configuration is successful
```

Successfully see Excel file content? 🎉 **Congratulations, setup complete!**

---

## 💡 What Can I Do?

### 🎮 Game Development Scenarios

**Game Designers**:
- "Help me increase all attack power in skill table by 10%"
- "Find equipment with price over 1000 in equipment table"
- "Merge skill table and class table, group by class"

**Balance Designers**:
- "Calculate average attack power for each class"
- "Find top 5 skills with highest attack power"
- "Batch modify skill cooldown times"

**Level Designers**:
- "Read level configuration table, find all collectible items"
- "Batch modify monster drop rates"

### 📊 Data Analysis Scenarios

**Data Processing**:
- "Read sales data, calculate total for each month"
- "Find customers with sales over 1000"
- "Merge data from multiple Excel files"

### 🚀 Advanced Features

- **Cross-table JOIN**: `Connect skill table and equipment table, find characters with both skills and equipment`
- **SQL Queries**: `SELECT * FROM skills WHERE attack_power > 100`
- **Batch Operations**: `Batch modify data in multiple files`
- **Version Comparison**: `Compare differences between two Excel versions`
- **AI-Enhanced Error Handling**: `Intelligent error detection with AI-powered suggestions and recovery guidance`

---

## 📚 Usage Examples

### Basic Operations
```
Read skill_table.xlsx
Create new skill data
Modify skill cooldown time
Save modifications to new file
```

### Advanced Queries
```
Connect skill table and class table, group by class to count skills
Query all skills with attack power over 100
Batch modify durability for multiple equipment
```

### Game Development Specific
```
Generate RPG game character attribute table
Calculate equipment set bonus effects
Balance game numerical parameters
```

---

## 🛠️ 支持的AI客户端

| 客户端 | 支持状态 | 配置难度 |
|--------|----------|----------|
| Claude Desktop | ✅ 完美支持 | ⭐ 简单 |
| Cursor | ✅ 完美支持 | ⭐ 简单 |
| Cherry Studio | ✅ 支持 | ⭐⭐ 中等 |
| VSCode + Continue | ✅ 支持 | ⭐⭐ 中等 |
| OpenClaw | ✅ Built-in Support | ⭐ Easiest |

---

## 🎯 Core Advantages

### ✅ vs Traditional Excel Tools
| Feature | ExcelMCP | Traditional Excel |
|--------|----------|-------------------|
| Learning Curve | 0 (Natural Language) | High (Requires formula knowledge) |
| Cross-table Ops | ✅ Automatic | ❌ Complex VLOOKUP |
| Batch Modify | ✅ One Command | ❌ Manual Operations |
| Error Handling | ✅ Smart Hints | ❌ Easy to Make Mistakes |
| Version Control | ✅ Automatic Records | ❌ Manual Management |

### ✅ vs Other AI Tools
| Feature | ExcelMCP | ChatGPT Plugin | Claude Desktop |
|--------|----------|----------------|----------------|
| Excel Operations | ✅ Specialized Optimized | ❌ Many Restrictions | ❌ No Support |
| Game Development | ✅ Specialized Scenarios | ❌ General Purpose | ❌ No Support |
| Performance | ✅ Fast Response | ⚡ Medium | ⚡ Medium |
| Privacy | ✅ Local Processing | ❌ Upload to Cloud | ❌ Upload to Cloud |

---

## 🔧 故障排除

### 常见问题

**❌ Python版本过低**
```bash
# Check Python version
python --version
# 需要3.10+，否则升级
# Mac/Linux: 使用brew install python
# Windows: Download latest version
```

**❌ 网络连接问题**
```bash
# Chinese users use mirror source
pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
```

**❌ AI Client Configuration Error**
- Ensure config file format is correct (JSON format)
- Restart AI client
- Check if uvx is available: `uvx --version`

### Get Help
```
Ask AI to run: excel-mcp-server-fastmcp --help
See full docs: https://github.com/TangentDomain/excel-mcp-server
```

---

## 📈 Performance Metrics

- **Response Speed**: < 1s (small files), < 5s (large files)
- **Supported Formats**: .xlsx, .xlsm, .xlsb
- **Max Support**: 100,000 rows × 1,000 columns
- **Memory Usage**: < 100MB (typical files)
- **Tool Count**: 53 specialized tools

---

## 🤝 Contributing & Support

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for ways to contribute.

### 🌟 Star Our Project
If this tool helps you, please give us a ⭐ Star!
- ⭐ **Star Support**: Help more game developers discover this tool
- 🔄 **Follow Updates**: Star to receive project notifications
- 📈 **Community Growth**: Every Star motivates us to improve

### Quick Contributions
- 🐛 Report Bugs: Use issue template
- 💡 Suggestions: Welcome new feature ideas  
- 📚 Improve Docs: Help other users
- 💻 Submit Code: See [Contribution Guide](CONTRIBUTING.md)

### 🎯 Project Goal
We aim to be the preferred tool for game development Excel configuration table management, **currently 4⭐, target 100⭐**! Every Star helps us reach this goal.

## 📄 许可证

[MIT License](LICENSE) - 开源友好，可商用

---

## 🎉 Acknowledgments

Thanks to all contributors and users! Special thanks to the game development community for feedback and suggestions.

Empowering game development with AI! 🚀

---


## 📊 GitHub 统计


[![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Star&color=gold)](https://github.com/TangentDomain/excel-mcp-server/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/network)
[![GitHub issues](https://img.shields.io/github/issues/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/issues)
[![GitHub language](https://img.shields.io/github/languages/top/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server)
[![GitHub last commit](https://img.shields.io/github/last-commit/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/commits/main)



| 指标 | 数值 | 状态 |
|------|------|------|
| ⭐ Stars | 4 | 🎯 **目标: 100** |
| 🍴 Forks | 0 | 📈 活跃度 |
| 👀 Watchers | 0 | 🔔 关注度 |
| 🐛 Issues | 0 | 📝 待处理 |
| 💻 语言 | Python | 🛠️ 技术 |
| 👥 贡献者 | 1 | 🤝 社区 |
| 📝 最近提交 | 10 | 🚀 活跃度 |

## 🎯 里程碑进度

- ⏳ **50 Stars**: 4 / 50
- ⏳ **100 Stars**: 4 / 100
- ⏳ **200 Stars**: 4 / 200
- ⏳ **500 Stars**: 4 / 500


## 📈 项目状态
- **创建时间**: 2025-09-22
- **最后更新**: 2026-03-28
- **社区活跃度**: 🔥 高度活跃
- **发展潜力**: 🌱 成长中

## 🤝 参与方式

感谢关注！您可以通过以下方式参与项目：

1. 🌟 **Star**: 如果项目对您有帮助，请给我们一个 Star
2. 🐛 **Issue**: 报告 Bug 或提出功能建议
3. 💻 **Code**: 提交代码改进和修复
4. 📚 **Docs**: 改进文档和使用示例
5. 📢 **Share**: 分享项目给更多开发者



