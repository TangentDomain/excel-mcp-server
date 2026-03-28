[简体中文](README.md) | [English](README.en.md)

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/v1.6.48)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1171-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)
![GitHub stars](https://img.shields.io/github/stars/TangentDomain/excel-mcp-server?style=social&label=Stars)

> **专为游戏开发者打造的 AI 驱动 Excel 配置管理工具**
> 
> 用自然语言或 SQL 操作游戏配置数据，支持跨表 JOIN、版本对比、批量修改

---

## 🎯 一句话介绍

> "我要把技能攻击力全部加10%，装备按稀有度排序，找出法师职业所有技能"

**只需要说这句话，ExcelMCP 自动帮你完成所有操作！**

---

## 🚀 快速开始（2分钟）

### 🔥 超简单安装（任选一种）

#### 🎯 推荐：uvx（最简单，无安装）
```bash
# Mac/Linux 一行命令
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows PowerShell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

使用：
```bash
uvx excel-mcp-server-fastmcp
```

#### 📦 传统：pip
```bash
pip install excel-mcp-server-fastmcp
```

> 💡 国内用户：`pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple`

### 🔗 AI 客户端配置（1分钟）

#### Claude Desktop（推荐）
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
- 设置 → MCP → Add Server
- Name: `excelmcp`
- Command: `uvx`
- Args: `["excel-mcp-server-fastmcp"]`

#### 其他客户端
- Cherry Studio / VSCode + Continue：同样配置
- OpenClaw：内置支持，无需配置

### ✅ 验证成功
在 AI 客户端说："帮我读取技能表测试一下" 
看到 Excel 数据？🎉 **配置完成！**

---

## 🎮 游戏开发场景演示

### 🎯 策划日常工作
```bash
# 自然语言指令
"帮我把技能表里所有法师技能攻击力加 20%"
"找出装备表中价格超过 1000 的史诗装备"
"把技能表和职业表关联，统计每个职业的技能数量"
"复制技能表到新文件，命名为技能备份_v2.xlsx"
```

### 🔢 数值策划任务
```bash
# 数据分析
"计算每个职业的平均攻击力和防御力"
"找出攻击力最高的前 10 个技能"
"批量调整所有技能冷却时间，乘以 0.8"
"生成角色属性平衡报告"
```

### 🏗️ 关卡策划需求
```bash
# 关卡配置
"读取关卡配置表，找出所有可收集道具"
"批量修改怪物掉落概率，稀有物品提升 50%"
"生成关卡进度统计报表"
```

---

## 📊 核心功能对比

| 场景 | ExcelMCP | 传统 Excel | ChatGPT |
|------|----------|------------|---------|
| **学习成本** | 🟢 0（直接说人话） | 🔴 需要公式学习 | 🟡 需要描述清楚 |
| **跨表操作** | 🟢 自动 JOIN | 🔴 复杂 VLOOKUP | 🔴 不支持 |
| **批量修改** | 🟢 一条指令搞定 | 🔴 手动操作 | 🟡 需要详细描述 |
| **错误处理** | 🟢 智能提示 | 🔴 容易出错 | 🟡 依赖 AI 能力 |
| **游戏优化** | 🟢 专属优化 | 🔴 通用功能 | 🔴 不专业 |

---

## 🛠️ 支持的游戏类型

| 游戏类型 | 支持场景 | 特色功能 |
|----------|----------|----------|
| **RPG** | 技能系统、装备套装、属性成长 | CTE 查询、装备加成计算 |
| **MMO** | 大数据量配置、版本管理 | 流式写入、缓存优化 |
| **卡牌** | 卡牌效果、概率计算 | 条件格式、数据验证 |
| **策略** | 单位配置、战斗计算 | 跨文件 JOIN、批量操作 |
| **休闲** | 关卡配置、道具管理 | 简单查询、快速修改 |

---

## 💡 使用技巧

### 🎯 高效指令示例
```bash
# 数据分析
"分析技能平衡性，找出伤害过高的技能"
"计算装备套装加成效果，按总价排序"
"统计怪物掉落，找出最值钱的掉落"

# 批量操作
"批量修改所有武器耐久度 +20%"
"复制装备表到不同品质分类"
"生成职业配装推荐"

# 版本管理
"比较技能表新旧版本差异"
"创建配置文件备份"
"回滚到指定版本"
```

### 🚀 性能优化
- **大文件**：使用流式写入，支持 10万+ 行数据
- **复杂查询**：自动索引优化，响应 < 3 秒
- **内存占用**：典型文件 < 100MB
- **支持格式**：.xlsx、.xlsm、.xlsb

---

## 📚 文档资源

### 📖 快速上手
- [基础教程](docs/README-gaming.md) - 游戏开发入门指南
- [性能优化](docs/README-performance.md) - 大文件处理技巧
- [SQL 参考](docs/README-sql.md) - 高级查询语法

### 🎮 示例代码
- [游戏开发示例](examples/README.md) - 完整技能系统、装备管理案例
- [批量操作示例](examples/进阶操作/) - 数据批处理、版本对比
- [实战案例](examples/实战案例/) - 完整游戏数值平衡方案

---

## 🔧 故障排除

### ❌ 常见问题

**Python 版本问题**
```bash
python --version  # 需要 3.10+
# 升级：https://www.python.org/downloads/
```

**网络问题**
```bash
# 国内镜像源
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple excel-mcp-server-fastmcp
```

**配置错误**
- 检查 JSON 格式是否正确
- 重启 AI 客户端
- 验证 uvx：`uvx --version`

### 🆘 获取帮助
```bash
# 命令行帮助
excel-mcp-server-fastmcp --help

# 项目文档
https://github.com/TangentDomain/excel-mcp-server

# 提交问题
GitHub Issues → 使用 Bug 报告模板
```

---

## 📈 技术规格

- **响应速度**：小文件 < 1秒，大文件 < 5秒
- **数据规模**：10万行 × 1000列
- **工具数量**：53 个游戏专用工具
- **内存占用**：< 100MB（典型文件）
- **支持格式**：.xlsx、.xlsm、.xlsb

---

## 🤝 参与贡献

### 🌟 给个 Star 吧！
如果这个工具对你有帮助，请点亮 ⭐ Star
- 🔍 **发现工具**：帮助更多游戏开发者找到我们
- 🔔 **获取更新**：Star 后第一时间收到功能更新
- 🎮 **推动生态**：每一个 Star 都是我们改进的动力

### 💪 如何贡献
- 🐛 **报告 Bug**：使用 [Issue 模板](https://github.com/TangentDomain/excel-mcp-server/issues/new)
- 💡 **功能建议**：欢迎提出游戏开发新需求
- 📚 **改进文档**：让其他开发者更容易上手
- 💻 **提交代码**：查看 [贡献指南](CONTRIBUTING.md)

---

## 📄 许可证

[MIT License](LICENSE) - 开源友好，可商用

---

## 🎉 致谢

感谢所有贡献者和游戏开发者社区！特别感谢：
- 游戏策划和数值策划的宝贵反馈
- 测试用户提供的真实使用场景
- 开发者社区的代码贡献

---

**用 AI 重新定义游戏开发配置管理！** 🚀
