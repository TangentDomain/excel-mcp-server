[简体中文](README.md) | [English](README.en.md)

# 🎮 ExcelMCP: Game Development Specific Excel Configuration Table Manager

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v1.6.48)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1171-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)
![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Stars)

> **AI-Driven Excel Configuration Table Management Tool** - Use natural language or SQL to operate game configuration data, supporting cross-table JOIN, version comparison, and batch modifications

---

## 🚀 3-Minute Quick Start (Follow these steps)

### ✅ Step 1: Check Python Environment (10 seconds)

Open terminal, type:
```bash
python --version
```

See `Python 3.10+`? ✅ **Skip to Step 2**

No Python? Download from [python.org](https://www.python.org/downloads/) (Windows users: remember to check "Add Python to PATH")

### ⚡ Step 2: Install Tool (Choose one, 30 seconds)

#### 🎯 Recommended: uvx (Fastest, no installation)
```bash
# Mac/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell Admin Mode)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Restart terminal, then verify:
```bash
uvx --version
```

#### 💾 Traditional: pip (Stable)
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 **Users in China** having slow downloads? Use mirror:
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### 🔧 Step 3: Configure AI Client (1 minute)

Find your AI client and follow the instructions:

#### 🟢 Claude Desktop (Recommended)

1. Open config file:
   - **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. Add config (if you have other content, add to `mcpServers`):
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

3. Save file, **restart Claude Desktop**

#### 🟡 Cursor

1. Open settings: `Ctrl+,` or `Cmd+,`
2. Search for "mcp", click "Model Context Protocol"
3. Click "Add MCP Server"
4. Fill in:
   - **Name**: `excelmcp`
   - **Command**: `uvx`
   - **Args**: `["excel-mcp-server-fastmcp"]`

5. Restart Cursor

#### 🔴 Other Clients

- **Cherry Studio**: Settings → MCP → Add Server
- **VSCode + Continue**: Settings → MCP → Add Server
- **OpenClaw**: Built-in support

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

## 🛠️ Supported AI Clients

| Client | Support Status | Config Difficulty |
|--------|----------------|-------------------|
| Claude Desktop | ✅ Perfect Support | ⭐ Easy |
| Cursor | ✅ Perfect Support | ⭐ Easy |
| Cherry Studio | ✅ Support | ⭐⭐ Medium |
| VSCode + Continue | ✅ Support | ⭐⭐ Medium |
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

## 🔧 Troubleshooting

### Common Issues

**❌ Python Version Too Low**
```bash
# Check Python version
python --version
# Need 3.10+, upgrade if needed
# Mac/Linux: Use brew install python
# Windows: Download latest version
```

**❌ Network Connection Issues**
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
- 💻 Submit Code: See [CONtribution Guide](CONTRIBUTING.md)

### 🎯 Project Goal
We aim to be the preferred tool for game development Excel configuration table management, **currently 4⭐, target 100⭐**! Every Star helps us reach this goal.

## 📄 License

[MIT License](LICENSE) - Open source, commercial friendly

---

## 🎉 Acknowledgments

Thanks to all contributors and users! Special thanks to the game development community for feedback and suggestions.

---

## 📊 GitHub Statistics

[![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Star&color=gold)](https://github.com/TangentDomain/excel-mcp-server/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/network)
[![GitHub issues](https://img.shields.io/github/issues/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/issues)
[![GitHub language](https://img.shields.io/github/languages/top/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server)
[![GitHub last commit](https://img.shields.io/github/last-commit/TangentDomain/excel-mcp-server?style=social)](https://github.com/TangentDomain/excel-mcp-server/commits/main)

| Metric | Value | Status |
|--------|-------|--------|
| ⭐ Stars | 4 | 🎯 **Target: 100** |
| 🍴 Forks | 0 | 📈 Activity |
| 👀 Watchers | 0 | 🔔 Attention |
| 🐛 Issues | 0 | 📝 Pending |
| 💻 Language | Python | 🛠️ Technology |
| 👥 Contributors | 1 | 🤝 Community |
| 📝 Recent Commits | 10 | 🚀 Activity |

## 🎯 Milestone Progress

- ⏳ **50 Stars**: 4 / 50
- ⏳ **100 Stars**: 4 / 100
- ⏳ **200 Stars**: 4 / 200
- ⏳ **500 Stars**: 4 / 500

## 📈 Project Status
- **Created**: 2025-09-22
- **Last Updated**: 2026-03-28
- **Community Activity**: 🔥 Highly Active
- **Growth Potential**: 🌱 Growing

## 🤝 How to Contribute

Thanks for your interest! You can contribute to the project in several ways:

1. 🌟 **Star**: If the project is helpful to you, please give us a Star
2. 🐛 **Issue**: Report bugs or suggest new features
3. 💻 **Code**: Submit code improvements and fixes
4. 📚 **Docs**: Improve documentation and usage examples
5. 📢 **Share**: Share the project with more developers

Empowering game development with AI! 🚀