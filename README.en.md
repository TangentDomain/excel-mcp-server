[简体中文](README.md) | [English](README.en.md)

# 🎮 ExcelMCP: Game Development Specific Excel Configuration Table Manager

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v1.6.48)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1171-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)
![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Stars)

> **AI-Driven Excel Configuration Management Tool for Game Developers**
> 
> Use natural language or SQL to operate game configuration data, supporting cross-table JOIN, version comparison, and batch modifications

---

## 🎯 One Sentence Introduction

> "I want to increase all skill attack power by 10%, sort equipment by rarity, and find all mage skills"

**Just say this sentence, and ExcelMCP automatically completes all operations for you!**

---

## 🚀 Quick Start (2 minutes)

### 🔥 Super Simple Installation (Choose one)

#### 🎯 Recommended: uvx (Easiest, no installation)
```bash
# Mac/Linux one-line command
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows PowerShell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Use:
```bash
uvx excel-mcp-server-fastmcp
```

#### 📦 Traditional: pip
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 **Users in China**: `pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple`

### 🔗 AI Client Configuration (1 minute)

#### Claude Desktop (Recommended)
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

#### Cursor
- Settings → MCP → Add Server
- Name: `excelmcp`
- Command: `uvx`
- Args: `["excel-mcp-server-fastmcp"]`

#### Other Clients
- Cherry Studio / VSCode + Continue: Same configuration
- OpenClaw: Built-in support, no configuration needed

### ✅ Verify Success
In your AI client, say: "Please help me read a skill table to test"
See Excel data? 🎉 **Setup complete!**

---

## 🎮 Game Development Scenarios

### 🎯 Game Designer Daily Work
```bash
# Natural language commands
"Help me increase all mage skill attack power by 20%"
"Find all epic equipment with price over 1000 in the equipment table"
"Connect skill table and class table, count skills per class"
"Copy skill table to new file named skill_backup_v2.xlsx"
```

### 🔢 Balance Designer Tasks
```bash
# Data analysis
"Calculate average attack and defense for each class"
"Find top 10 skills with highest attack power"
"Batch adjust all skill cooldown times, multiply by 0.8"
"Generate character attribute balance report"
```

### 🏗️ Level Designer Requirements
```bash
# Level configuration
"Read level configuration table, find all collectible items"
"Batch modify monster drop rates, increase rare items by 50%"
"Generate level progress statistics report"
```

---

## 📊 Core Feature Comparison

| Scenario | ExcelMCP | Traditional Excel | ChatGPT |
|----------|----------|------------------|---------|
| **Learning Curve** | 🟢 0 (Speak naturally) | 🔴 Need formula learning | 🟡 Need clear descriptions |
| **Cross-table Ops** | 🟢 Automatic JOIN | 🔴 Complex VLOOKUP | 🔴 No support |
| **Batch Modify** | 🟢 One command搞定 | 🔴 Manual operations | 🟡 Need detailed description |
| **Error Handling** | 🟢 Smart hints | 🔴 Easy to make mistakes | 🟡 Depends on AI capability |
| **Game Optimization** | 🟢 Specialized | 🔴 General features | 🔴 Not professional |

---

## 🛠️ Supported Game Types

| Game Type | Supported Scenarios | Special Features |
|-----------|---------------------|------------------|
| **RPG** | Skill systems, equipment sets, attribute growth | CTE queries, equipment bonus calculation |
| **MMO** | Large data configs, version management | Streaming writes, cache optimization |
| **Card Games** | Card effects, probability calculations | Conditional formatting, data validation |
| **Strategy** | Unit configs, combat calculations | Cross-file JOIN, batch operations |
| **Casual** | Level configs, item management | Simple queries, quick modifications |

---

## 💡 Usage Tips

### 🎯 High-Efficiency Command Examples
```bash
# Data analysis
"Analyze skill balance, find skills with too high damage"
"Calculate equipment set bonus effects, sort by total value"
"Count monster drops, find most valuable loot"

# Batch operations
"Batch increase all weapon durability by 20%"
"Copy equipment table to different quality categories"
"Generate class equipment recommendations"

# Version management
"Compare skill table old and new version differences"
"Create configuration file backup"
"Rollback to specific version"
```

### 🚀 Performance Optimization
- **Large files**: Use streaming writes, supports 100K+ rows
- **Complex queries**: Auto-index optimization, response < 3 seconds
- **Memory usage**: Typical files < 100MB
- **Supported formats**: .xlsx, .xlsm, .xlsb

---

## 📚 Documentation Resources

### 📖 Quick Start Guides
- [Basic Tutorial](docs/README-gaming.md) - Game development getting started
- [Performance Optimization](docs/README-performance.md) - Large file handling techniques
- [SQL Reference](docs/README-sql.md) - Advanced query syntax

### 🎮 Example Code
- [Game Development Examples](examples/README.md) - Complete skill systems, equipment management cases
- [Batch Operations](examples/进阶操作/) - Data batch processing, version comparison
- [Real Cases](examples/实战案例/) - Complete game numerical balance solutions

---

## 🔧 Troubleshooting

### ❌ Common Issues

**Python Version Problems**
```bash
python --version  # Requires 3.10+
# Upgrade: https://www.python.org/downloads/
```

**Network Issues**
```bash
# China mirror source
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple excel-mcp-server-fastmcp
```

**Configuration Errors**
- Check JSON format is correct
- Restart AI client
- Verify uvx: `uvx --version`

### 🆘 Get Help
```bash
# Command line help
excel-mcp-server-fastmcp --help

# Project documentation
https://github.com/TangentDomain/excel-mcp-server

# Report issues
GitHub Issues → Use Bug Report Template
```

---

## 📈 Technical Specifications

- **Response Speed**: < 1s (small files), < 5s (large files)
- **Data Scale**: 100K rows × 1000 columns
- **Tool Count**: 53 game-specific tools
- **Memory Usage**: < 100MB (typical files)
- **Supported Formats**: .xlsx, .xlsm, .xlsb

---

## 🤝 Contributing

### 🌟 Give us a Star!
If this tool helps you, please light up ⭐ Star
- 🔍 **Discover Tool**: Help more game developers find us
- 🔔 **Get Updates**: Star to receive feature updates first
- 🎮 **Drive Ecosystem**: Every Star motivates us to improve

### 💪 How to Contribute
- 🐛 **Report Bugs**: Use [Issue Template](https://github.com/TangentDomain/excel-mcp-server/issues/new)
- 💡 **Feature Suggestions**: Welcome new game development needs
- 📚 **Improve Documentation**: Help other developers get started faster
- 💻 **Submit Code**: See [Contribution Guide](CONTRIBUTING.md)

---

## 📄 License

[MIT License](LICENSE) - Open source, commercial friendly

---

## 🎉 Acknowledgments

Thanks to all contributors and the game developer community! Special thanks to:
- Game designers and balance designers for valuable feedback
- Test users providing real-world usage scenarios
- Developer community code contributions

---

**Redeveloping game development configuration management with AI!** 🚀
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