# Excel MCP Server

一个基于FastMCP和openpyxl的Excel文件操作服务器，提供MCP（Model Context Protocol）接口。

## 🚀 功能特性

- **Excel文件操作**: 读取、写入、搜索Excel文件
- **工作表管理**: 列出、创建、重命名、删除工作表
- **数据操作**: 获取和更新单元格范围数据
- **行列操作**: 插入、删除行和列
- **搜索功能**: 正则表达式搜索和替换

## 📁 项目结构

```
excel-mcp-server/
├── src/excel_mcp/          # 源码目录
│   ├── models/             # 数据模型
│   ├── utils/              # 工具模块
│   ├── core/               # 核心功能
│   └── server_new.py       # MCP服务器接口
├── tests/                  # 测试目录
├── data/                   # 数据文件
├── docs/                   # 文档
├── scripts/                # 脚本文件
└── archive/                # 归档文件
```

## 🛠️ 环境设置

### 使用uv（推荐）

```bash
# 创建虚拟环境
uv venv

# 安装依赖
uv pip install -e ".[dev]"
```

### 使用pip

```bash
# 创建虚拟环境
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# 或 .venv\Scripts\activate  # Windows

# 安装依赖
pip install -e ".[dev]"
```

## 🧪 运行测试

```bash
# 运行完整测试套件
uv run python run_tests.py

# 或单独运行测试
uv run pytest tests/ -v
uv run python tests/test_runner.py
```

## 🏃 启动服务器

```bash
uv run python src/excel_mcp/server_new.py
```

## 📋 测试覆盖

- ✅ **基础功能测试**: 4个测试全部通过
- ✅ **解析器测试**: 8个测试全部通过  
- ✅ **验证器测试**: 17个测试全部通过
- 📊 **总计**: 29个测试全部通过

## 🏗️ 架构设计

### 模块化架构

- **`models/`**: 数据类型和枚举定义
- **`utils/`**: 验证器、解析器、异常处理
- **`core/`**: Excel读取、写入、管理、搜索核心功能
- **`server_new.py`**: 纯MCP接口层

### 设计原则

- 单一职责原则
- 接口分离
- 依赖注入
- 统一错误处理
- 全面测试覆盖

## 📦 依赖包

- **fastmcp**: MCP服务器框架
- **openpyxl**: Excel文件操作
- **pytest**: 测试框架

## 🤝 开发指南

1. 所有新功能必须包含测试
2. 遵循现有的模块化架构
3. 使用类型提示
4. 包含适当的错误处理
5. 运行测试确保通过

## 📄 许可证

[添加您的许可证信息]