[简体中文](README.md) ｜ [English](README.en.md)

# 🎮 ExcelMCP: 游戏开发专用 Excel 配置表管理器

[![PyPI](https://img.shields.io/pypi/v/excel-mcp-server-fastmcp.svg)](https://pypi.org/project/excel-mcp-server-fastmcp/)
[![CI](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/TangentDomain/excel-mcp-server/actions/workflows/ci.yml)
![Tests](https://img.shields.io/badge/tests-1161-brightgreen.svg)
![Tools](https://img.shields.io/badge/tools-53-green.svg)

> **AI驱动的Excel配置表管理工具** - 用自然语言或SQL操作游戏配置数据，支持跨表JOIN、版本对比、批量修改

---

## 🚀 3分钟快速上手（跟着做就行）

### ✅ 第1步：确认Python环境（10秒）

打开终端，输入：
```bash
python --version
```

看到 `Python 3.10+`？✅ **直接跳到第2步**

没装Python？去 [python.org](https://www.python.org/downloads/) 下载安装（Windows用户记得勾选 "Add Python to PATH"）

### ⚡ 第2步：安装工具（二选一，30秒）

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

> 💡 **国内用户**下载慢？加镜像源：
> ```bash
> pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
> ```

### 🔧 第3步：配置AI客户端（1分钟）

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

### ✅ 第4步：验证配置（10秒）

重启AI客户端后，让AI助手帮你测试：
```
请帮我读取Excel文件测试配置是否成功
```

成功看到Excel文件内容？🎉 **恭喜你，配置完成！**

---

## 💡 我能做什么？

### 🎮 游戏开发场景

**游戏策划**：
- "帮我把技能表里的所有攻击力提升10%"
- "找出装备表中价格超过1000的装备"
- "合并技能表和职业表，按职业分组"

**数值策划**：
- "计算每个职业的平均攻击力"
- "找出攻击力最高的前5个技能"
- "批量修改技能冷却时间"

**关卡策划**：
- "读取关卡配置表，找出所有需要收集的道具"
- "批量修改怪物掉落概率"

### 📊 数据分析场景

**数据处理**：
- "读取销售数据，计算每个月的总额"
- "找出销售额超过1000的客户"
- "合并多个Excel文件的数据"

### 🚀 高级功能

- **跨表JOIN**：`连接技能表和装备表，找出同时拥有技能和装备的角色`
- **SQL查询**：`SELECT * FROM 技能表 WHERE 攻击力 > 100`
- **批量操作**：`批量修改多个文件的数据`
- **版本对比**：`对比两个版本的Excel文件差异`

---

## 📚 使用示例

### 基础操作
```
读取技能表.xlsx
创建新技能数据
修改技能冷却时间
保存修改到新文件
```

### 高级查询
```
连接技能表和职业表，按职业分组统计技能数量
查询所有攻击力超过100的技能
批量修改多个装备的耐久度
```

### 游戏开发专用
```
生成RPG游戏角色属性表
计算装备套装加成效果
平衡游戏数值参数
```

---

## 🛠️ 支持的AI客户端

| 客户端 | 支持状态 | 配置难度 |
|--------|----------|----------|
| Claude Desktop | ✅ 完美支持 | ⭐ 简单 |
| Cursor | ✅ 完美支持 | ⭐ 简单 |
| Cherry Studio | ✅ 支持 | ⭐⭐ 中等 |
| VSCode + Continue | ✅ 支持 | ⭐⭐ 中等 |
| OpenClaw | ✅ 内置支持 | ⭐ 最简单 |

---

## 🎯 核心优势

### ✅ 对比传统Excel工具
| 特性 | ExcelMCP | 传统Excel |
|------|----------|-----------|
| 学习成本 | 0（自然语言） | 高（需要函数知识） |
| 跨表操作 | ✅ 自动处理 | ❌ 复杂VLOOKUP |
| 批量修改 | ✅ 一行指令 | ❌ 手动操作 |
| 错误处理 | ✅ 智能提示 | ❌ 容易出错 |
| 版本管理 | ✅ 自动记录 | ❌ 需要手动 |

### ✅ 对比其他AI工具
| 特性 | ExcelMCP | ChatGPT插件 | Claude Desktop |
|------|----------|--------------|----------------|
| Excel操作 | ✅ 专用优化 | ❌ 限制多 | ❌ 无支持 |
| 游戏开发 | ✅ 专用场景 | ❌ 通用 | ❌ 无支持 |
| 性能表现 | ✅ 快速响应 | ⚡ 中等 | ⚡ 中等 |
| 隐私保护 | ✅ 本地处理 | ❌ 上传云端 | ❌ 上传云端 |

---

## 🔧 故障排除

### 常见问题

**❌ Python版本过低**
```bash
# 检查Python版本
python --version
# 需要3.10+，否则升级
# Mac/Linux: 使用brew install python
# Windows: 下载最新版本
```

**❌ 网络连接问题**
```bash
# 国内用户使用镜像源
pip install excel-mcp-server-fastmcp -i https://pypi.tuna.tsinghua.edu.cn/simple
```

**❌ AI客户端配置错误**
- 确保配置文件格式正确（JSON格式）
- 重启AI客户端
- 检查uvx是否可用：`uvx --version`

### 获取帮助
```
让AI助手运行：excel-mcp-server-fastmcp --help
查看完整文档：https://github.com/TangentDomain/excel-mcp-server
```

---

## 📈 性能指标

- **响应速度**：< 1秒（小文件），< 5秒（大文件）
- **支持格式**：.xlsx, .xlsm, .xlsb
- **最大支持**：10万行 × 1000列
- **内存占用**：< 100MB（典型文件）
- **工具数量**：53个专用工具

---

## 🤝 贡献

欢迎贡献代码和建议！查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解参与方式。

### 快速贡献
- 🐛 报告Bug：使用issue模板
- 💡 提出建议：欢迎新的功能想法  
- 📚 完善文档：帮助其他用户
- ⭐ Star支持：让更多人发现这个工具

## 📄 许可证

[MIT License](LICENSE) - 开源友好，可商用

---

## 🎉 致谢

感谢所有贡献者和用户的支持！特别感谢游戏开发社区的反馈和建议。

让AI为游戏开发赋能！ 🚀