[简体中文](README.md) ｜ [English](README.en.md)

# 🎮 ExcelMCP: Game Development Specific Excel Configuration Table Manager

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1187-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)

> AI-driven Excel configuration table management tool. Use natural language or SQL to operate game configuration data, supporting cross-table JOIN, version comparison, and batch modifications.

## 🤔 What is this?

**One sentence**: Let AI read and write Excel configuration tables for you.

**Who uses it**: Game designers (skill tables, equipment tables, numerical balancing), data analysts, anyone who needs AI to operate Excel.

**What you need**:
- ✅ An AI client (Claude Desktop / Cursor / Cherry Studio / any MCP-supported client)
- ✅ Python 3.10+ (you probably already have it)
- ❌ No need to clone code or manually start services

---

## 🚀 5-Minute Quick Start

### Step 1: Check Python (you might already have it)

Open terminal (Mac: `Cmd+space` search "terminal", Windows: `Win+R` type `cmd`), enter:
```bash
python --version
```

See `Python 3.10` or higher? **Skip to step 2**.

Not installed? Go to [python.org](https://www.python.org/downloads/) to download and install.
⚠️ **Windows users**: Make sure to check **"Add Python to PATH"** during installation.

### Step 2: Install the tool (choose one)

**Option A: Use uvx (recommended, faster)**

uvx is a uv command that can run Python tools directly from PyPI without manual installation.

First install uv:
```bash
# Mac / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell, right-click Start menu → "Terminal(Admin)")
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

After installing, **restart terminal**, then verify:
```bash
uvx --version
```

If you see a version number, it's working.

**Option B: Use pip (more traditional, but more stable)**

pip is Python's built-in package manager, no extra installation needed:
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 **Users in China** if download is slow, add a mirror source:
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### Step 3: Configure your AI client

Find your client and configure it according to the instructions:

#### 🟠 Claude Desktop

1. Open the config file:
   - **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
2. Add this section (if you already have other content, add it to `mcpServers`):

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

> If using option B (pip), change `"command"` to `"excel-mcp-server-fastmcp"` and remove the `"args"` line.

3. Save the file, **restart Claude Desktop**

#### 🟢 Cursor

1. Open Cursor → Settings → Features → MCP
2. Click "Add MCP Server"
3. Fill in:
   - **Name**: `excelmcp`
   - **Command**: `uvx`
   - **Args**: `excel-mcp-server-fastmcp`
4. Save and restart Cursor

#### 🔵 Cherry Studio / Other Clients

Find the MCP server configuration page and fill in the same `uvx` and `excel-mcp-server-fastmcp`.
The principle is the same: tell the client "what command to use to start this tool".

### Step 4: Start using

Restart your AI client, then just say to it:

```
→ "Help me open D:/game/config/skills.xlsx"
→ "What columns are in this table?"
→ "Search for all fire skills"
→ "Increase damage of all fire skills by 20%"
→ "Create a new copy called skills_v2.xlsx"
→ "Use SQL to find the top 10 skills by damage"
```

**Done.** You only need to know how to speak, no need to write code or learn commands.
Of course, if you know SQL, you can also use SQL directly - AI will help you translate and execute it.

---

## ❓ Common Questions

<details>
<summary><b>Error installing uv / network timeout?</b></summary>

Users in China may need to configure a proxy. Or just use option B (pip), which is more stable.
</details>

<details>
<summary><b>Getting "command not found: uvx"?</b></summary>

1. Confirm uv is installed: `uv --version`
2. Mac/Linux needs to restart terminal, or run `source ~/.bashrc` (or `source ~/.zshrc`)
3. If still not working, use option B (pip) instead
</details>

<details>
<summary><b>How to confirm installation succeeded?</b></summary>

```bash
uvx excel-mcp-server-fastmcp --help
```

Seeing help message means it's OK.
</details>

<details>
<summary><b>Which AI clients are supported?</b></summary>

Any client that supports MCP (Model Context Protocol):
- Claude Desktop / Claude Code
- Cursor
- Cherry Studio
- OpenClaw
- VS Code + Continue / Cline plugins
- Any MCP client that supports stdio mode
</details>

<details>
<summary><b>What's the difference from regular Excel tools?</b></summary>

ExcelMCP is an MCP server - it doesn't run separately, but is embedded in AI clients.
Your AI assistant (Claude/Cursor, etc.) can directly call it to read and write Excel files.
You don't need to learn any commands, just describe what you want to do in natural language.
</details>

## 📚 More Resources

Need more detailed information? Please check:

- [🏆 Competitor Comparison](./docs/README-comparison.md)
- [⚡ Performance Optimization](./docs/README-performance.md)
- [🎮 Game Development Scenarios](./docs/README-gaming.md)
- [📖 Complete Tool List](./docs/README-tools.md)
- [📊 Technical Architecture](./docs/README-architecture.md)
- [🔧 SQL Query Guide](./docs/README-sql.md)
- [📋 Changelog](./CHANGELOG.md)

---

## 🤝 Contributing

Contributions welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for how to participate in project development.

## 📄 License

[MIT License](LICENSE)