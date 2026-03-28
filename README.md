[简体中文](README.md) | [English](README.en.md)

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v1.6.45)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1171-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)
![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Stars)

> 用自然语言或 SQL 操作游戏配置数据，支持跨表 JOIN、版本对比、批量修改

---

## 🚀 3分钟快速上手

### ✅ 第一步：检查 Python 环境（10秒）

打开终端，输入：
```bash
python --version
```

看到 `Python 3.10+`？✅ **直接跳到第二步**

没装 Python？去 [python.org](https://www.python.org/downloads/) 下载（Windows 用户记得勾选 "Add Python to PATH"）

### ⚡ 第二步：安装工具（任选其一，30秒）

#### 🎯 推荐方式：uvx（最快，无需安装）
```bash
# Mac/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows（PowerShell 管理员模式）
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

> 💡 国内下载慢？用镜像源：
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### 🔧 第三步：配置 AI 客户端（1分钟）

找到你的 AI 客户端，按说明配置：

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

3. 保存文件，**重启 Claude Desktop**

#### 🟡 Cursor

1. 打开设置：`Ctrl+,` 或 `Cmd+,`
2. 搜索 "mcp"，点击 "Model Context Protocol"
3. 点击 "Add MCP Server"
4. 填入：
   - **Name**: `excelmcp`
   - **Command**: `uvx`
   - **Args**: `["excel-mcp-server-fastmcp"]`

5. 重启 Cursor

#### 🔴 其他客户端

- **Cherry Studio**: 设置 → MCP → 添加服务器
- **VSCode + Continue**: 设置 → MCP → 添加服务器
- **OpenClaw**: 已内置支持

### ✅ 第四步：验证配置（10秒）

重启 AI 客户端后，让 AI 帮你测试：
```
帮我读一个 Excel 文件，测试配置是否成功
```

能看到 Excel 文件内容？🎉 **恭喜，配置完成！**

---

## 💡 能做什么？

### 🎮 游戏开发场景

**策划**：
- "帮我把技能表里所有攻击力加 10%"
- "在装备表里找价格超过 1000 的装备"
- "把技能表和职业表关联，按职业分组"

**数值策划**：
- "算一下每个职业的平均攻击力"
- "找出攻击力最高的前 5 个技能"
- "批量修改技能冷却时间"

**关卡策划**：
- "读取关卡配置表，找出所有可收集道具"
- "批量修改怪物掉落概率"

### 📊 数据分析场景

- "读取销售数据，算每个月的总额"
- "找销售额超过 1000 的客户"
- "合并多个 Excel 文件的数据"

### 🚀 高级功能

- **跨表 JOIN**：关联技能表和装备表，找出同时拥有技能和装备的角色
- **SQL 查询**：`SELECT * FROM skills WHERE attack_power > 100`
- **批量操作**：批量修改多个文件中的数据
- **版本对比**：比较两个 Excel 版本的差异
- **AI 智能错误处理**：智能检测错误并提供修复建议

---

## 📚 使用示例

### 基础操作
```
读取技能表.xlsx
创建新的技能数据
修改技能冷却时间
把修改保存到新文件
```

### 高级查询
```
关联技能表和职业表，按职业统计技能数量
查询攻击力超过 100 的所有技能
批量修改多个装备的耐久度
```

### 游戏开发专用
```
生成 RPG 游戏角色属性表
计算装备套装加成效果
平衡游戏数值参数
```

---

## 🛠️ 支持的 AI 客户端

| 客户端 | 支持状态 | 配置难度 |
|--------|----------|----------|
| Claude Desktop | ✅ 完美支持 | ⭐ 简单 |
| Cursor | ✅ 完美支持 | ⭐ 简单 |
| Cherry Studio | ✅ 支持 | ⭐⭐ 中等 |
| VSCode + Continue | ✅ 支持 | ⭐⭐ 中等 |
| OpenClaw | ✅ 内置支持 | ⭐ 最简单 |

---

## 🎯 核心优势

### ✅ 对比传统 Excel 工具
| 功能 | ExcelMCP | 传统 Excel |
|------|----------|------------|
| 学习成本 | 0（自然语言） | 高（需要公式知识） |
| 跨表操作 | ✅ 自动关联 | ❌ 复杂的 VLOOKUP |
| 批量修改 | ✅ 一条指令 | ❌ 手动操作 |
| 错误处理 | ✅ 智能提示 | ❌ 容易出错 |
| 版本管理 | ✅ 自动记录 | ❌ 手动管理 |

### ✅ 对比其他 AI 工具
| 功能 | ExcelMCP | ChatGPT 插件 | Claude Desktop |
|------|----------|--------------|----------------|
| Excel 操作 | ✅ 专业优化 | ❌ 限制多 | ❌ 不支持 |
| 游戏开发 | ✅ 专属场景 | ❌ 通用场景 | ❌ 不支持 |
| 响应速度 | ✅ 快速 | ⚡ 中等 | ⚡ 中等 |
| 隐私安全 | ✅ 本地处理 | ❌ 上传云端 | ❌ 上传云端 |

---

## 🔧 故障排除

### 常见问题

**❌ Python 版本过低**
```bash
python --version
# 需要 3.10+，否则升级
# Mac/Linux: brew install python
# Windows: 去 python.org 下载最新版
```

**❌ 网络问题**
```bash
# 国内用户用镜像源
pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
```

**❌ AI 客户端配置错误**
- 检查配置文件格式是否正确（JSON 格式）
- 重启 AI 客户端
- 检查 uvx 是否可用：`uvx --version`

### 获取帮助
```
让 AI 运行: excel-mcp-server-fastmcp --help
查看文档: https://github.com/TangentDomain/excel-mcp-server
```

---

## 📈 性能指标

- **响应速度**：小文件 < 1秒，大文件 < 5秒
- **支持格式**：.xlsx、.xlsm、.xlsb
- **最大支持**：10万行 × 1000列
- **内存占用**：< 100MB（典型文件）
- **工具数量**：53 个专业工具

---

## 🤝 参与贡献

欢迎贡献代码！查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解贡献方式。

### 🌟 Star 支持我们
如果这个工具对你有帮助，请给个 ⭐ Star！
- ⭐ **Star 支持**：帮助更多游戏开发者发现这个工具
- 🔄 **关注更新**：Star 后会收到项目通知
- 📈 **社区成长**：每一个 Star 都是我们改进的动力

### 快速贡献
- 🐛 报告 Bug：使用 issue 模板
- 💡 功能建议：欢迎提出新想法
- 📚 改进文档：帮助其他用户
- 💻 提交代码：查看 [贡献指南](CONTRIBUTING.md)

---

## 📄 许可证

[MIT License](LICENSE) - 开源友好，可商用

---

## 🎉 鸣谢

感谢所有贡献者和用户！特别感谢游戏开发社区的反馈和建议。

用 AI 赋能游戏开发！🚀
