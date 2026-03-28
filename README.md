[简体中文](README.md) ｜ [English](README.en.md)

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1161-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)

> AI驱动的Excel配置表管理工具。用自然语言或SQL操作游戏配置数据，支持跨表JOIN、版本对比、批量修改。

## 🤔 这是什么？

**一句话**：让AI帮你读写Excel配置表。

**谁在用**：游戏策划（技能表、装备表、数值平衡）、数据分析师、任何需要用AI操作Excel的人。

**需要什么**：
- ✅ 一个AI客户端（Claude Desktop / Cursor / Cherry Studio / 任何支持MCP的客户端）
- ✅ Python 3.10+（你的电脑上可能已经有了）
- ❌ 不需要克隆代码、不需要手动启动服务

---

## 🚀 5分钟上手（跟着做就行）

### 第1步：确认有Python（可能已经有了）

打开终端（Mac按`Cmd+空格`搜"终端"，Windows按`Win+R`输入`cmd`），输入：
```bash
python --version
```

看到 `Python 3.10` 或更高版本？**直接跳到第2步**。

没装？去 [python.org](https://www.python.org/downloads/) 下载安装。
⚠️ **Windows用户**：安装时一定勾选 **"Add Python to PATH"**。

### 第2步：装运行工具（二选一）

**方式A：用 uvx（推荐，更快）**

uvx 是 uv 的一个命令，能直接从PyPI运行Python工具，不用手动安装。

先装 uv：
```bash
# Mac / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell，右键开始菜单→"终端(管理员)")
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

装完后**重启终端**，然后验证：
```bash
uvx --version
```

看到版本号就说明装好了。

**方式B：用 pip（更传统，但更稳定）**

pip 是Python自带的包管理器，不用额外装：
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 **中国大陆用户**如果下载慢，加个镜像源：
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### 第3步：配置你的AI客户端

找到你用的客户端，按说明配置：

#### 🟠 Claude Desktop

1. 打开配置文件：
   - **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
2. 在里面加这段（如果已有其他内容，加到 `mcpServers` 里）：

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

> 如果用的方式B（pip），把 `"command"` 改成 `"excel-mcp-server-fastmcp"`，删掉 `"args"` 那行。

3. 保存文件，**重启Claude Desktop**

#### 🟢 Cursor

1. 打开 Cursor → Settings → Features → MCP
2. 点 "Add MCP Server"
3. 填入：
   - **Name**: `excelmcp`
   - **Command**: `uvx`
   - **Args**: `excel-mcp-server-fastmcp`
4. 保存，重启Cursor

#### 🔵 Cherry Studio / 其他客户端

找到 MCP服务器配置页面，填入相同的 `uvx` 和 `excel-mcp-server-fastmcp`。
原理都一样：告诉客户端"用什么命令启动这个工具"。

### 第4步：开始用

重启你的AI客户端，然后直接对它说：

```
→ "帮我打开 D:/game/config/skills.xlsx"
→ "查看这张表有哪些列"
→ "搜索所有火系技能"
→ "把所有火系技能伤害提升20%"
→ "新建一个叫 skills_v2.xlsx 的副本"
→ "用SQL查一下伤害前10的技能"
```

**搞定。** 你只需要会说话，不需要写代码、不需要学命令。
当然，如果你懂SQL，也可以直接用SQL查询——AI会帮你翻译执行。

---

## ❓ 常见问题

<details>
<summary><b>装uv的时候报错 / 网络超时？</b></summary>

中国大陆用户可能需要配代理。或者直接用方式B（pip），更稳定。
</details>

<details>
<summary><b>提示 "command not found: uvx"？</b></summary>

1. 确认uv装成功了：`uv --version`
2. Mac/Linux 需要重启终端，或运行 `source ~/.bashrc`（或 `source ~/.zshrc`）
3. 如果还是不行，用方式B（pip）代替
</details>

<details>
<summary><b>怎么确认安装成功了？</b></summary>

```bash
uvx excel-mcp-server-fastmcp@1.6.41 --help
```

看到帮助信息就OK了。
</details>

<details>
<summary><b>支持哪些AI客户端？</b></summary>

任何支持MCP（Model Context Protocol）的客户端：
- Claude Desktop / Claude Code
- Cursor
- Cherry Studio
- OpenClaw
- VS Code + Continue / Cline 插件
- 任何支持 stdio 模式的MCP客户端
</details>

<details>
<summary><b>和普通Excel工具有什么区别？</b></summary>

ExcelMCP是MCP服务器——它不单独运行，而是嵌入到AI客户端里。
你的AI助手（Claude/Cursor等）可以直接调用它来读写Excel文件。
你不需要学任何命令，直接用自然语言描述你要做什么。
</details>

## 📚 更多资源

需要更详细的信息？请查看：

- [🏆 竞品对比](./docs/README-comparison.md)
- [⚡ 性能优化](./docs/README-performance.md)
- [🎮 游戏开发场景](./docs/README-gaming.md)
- [📖 完整工具列表](./docs/README-tools.md)
- [📊 技术架构](./docs/README-architecture.md)
- [🔧 SQL查询指南](./docs/README-sql.md)
- [📋 更新日志](./CHANGELOG.md)

---

## 🤝 贡献

欢迎贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解如何参与项目开发。

## 📄 许可证

[MIT License](LICENSE)