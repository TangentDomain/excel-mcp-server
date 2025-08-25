# Excel MCP Server 快速部署指南

## 项目概述

Excel MCP Server 是一个基于 FastMCP 框架的 Excel 文件操作服务器，提供以下核心功能：

- ⚡ **正则搜索**: 在Excel文件中快速搜索符合条件的数据
- 📊 **数据操作**: 读写、修改Excel单元格范围
- 🗂️ **文件管理**: 创建、删除Excel文件和工作表
- 🎨 **格式化**: 动态设置单元格格式
- 🔍 **批量操作**: 目录级别的Excel文件操作

## 核心架构

```
src/
├── server.py              # FastMCP服务器主入口
├── core/                  # 核心功能模块
│   ├── excel_reader.py    # Excel读取操作
│   ├── excel_writer.py    # Excel写入操作
│   ├── excel_manager.py   # 文件和工作表管理
│   ├── excel_search.py    # 正则搜索功能
│   └── excel_compare.py   # 游戏配置表比较
├── utils/                 # 工具模块
│   ├── error_handler.py   # 统一错误处理
│   ├── formatter.py       # 结果格式化
│   ├── validators.py      # 数据验证
│   └── parsers.py         # 数据解析
└── models/                # 数据模型
    └── types.py           # 类型定义
```

## 快速部署方式

### 方式一：使用UV包管理器（推荐）

**前提条件**: 已安装 `uv` 包管理器

```bash
# 1. 进入项目目录
cd D:\excel-mcp-server

# 2. 同步依赖
uv sync

# 3. 使用配置文件
# 将 mcp-windows.json 内容复制到你的MCP客户端配置中
```

**MCP客户端配置**:
```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uv",
      "args": [
        "--directory",
        "D:/excel-mcp-server",
        "run",
        "python",
        "-m",
        "src.server"
      ]
    }
  }
}
```

### 方式二：直接Python运行

**前提条件**: Python 3.10+ + 已安装依赖

```bash
# 1. 进入项目目录
cd D:\excel-mcp-server

# 2. 激活虚拟环境（如果使用）
.venv\Scripts\activate

# 3. 安装依赖（如果未安装）
pip install fastmcp openpyxl mcp xlcalculator formulas xlwings

# 4. 使用配置文件
# 将 mcp-direct.json 内容复制到你的MCP客户端配置中
```

**MCP客户端配置**:
```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "python",
      "args": [
        "-m",
        "src.server"
      ],
      "cwd": "D:/excel-mcp-server"
    }
  }
}
```

### 方式三：构建独立包

```bash
# 1. 构建wheel包
cd D:\excel-mcp-server
uv build

# 2. 安装到系统环境
pip install dist/excel_mcp_server_fastmcp-1.0.0-py3-none-any.whl

# 3. 直接运行（需要配置PYTHONPATH）
python -m src.server
```

## 部署验证

### 测试服务器启动
```bash
# 使用UV方式测试
uv run python -m src.server

# 或直接Python方式测试
python -m src.server
```

**注意**: MCP服务器通过标准输入输出通信，正常启动时不会显示任何输出，这是正常现象。

### 功能测试
可以通过MCP客户端测试以下功能：

1. **创建Excel文件**: `excel_create_file`
2. **读取数据**: `excel_get_range`
3. **写入数据**: `excel_update_range`
4. **搜索功能**: `excel_regex_search`
5. **工作表管理**: `excel_create_sheet`, `excel_delete_sheet`

## 常见问题

### 1. 导入错误
**问题**: `ImportError: 缺少必要的依赖包`
**解决**: 确保已安装所有依赖 `uv sync` 或 `pip install -r requirements.txt`

### 2. 路径问题
**问题**: 找不到模块或文件
**解决**: 确保使用绝对路径，Windows用户注意使用正斜杠

### 3. 权限问题
**问题**: 无法访问Excel文件
**解决**: 确保Python进程有读写Excel文件的权限

## 依赖清单

核心依赖:
- `fastmcp >= 0.1.0` - MCP服务器框架
- `openpyxl >= 3.1.0` - Excel文件操作
- `mcp >= 1.0.0` - MCP协议支持
- `xlcalculator >= 0.5.0` - Excel公式计算
- `formulas >= 1.3.0` - 公式解析引擎
- `xlwings >= 0.33.15` - Excel应用集成

开发依赖:
- `pytest >= 7.0.0`
- `pytest-asyncio >= 0.21.0`
- `coverage >= 7.10.4`

## 技术特性

- **统一错误处理**: 所有MCP工具使用 `@unified_error_handler` 装饰器
- **模块化设计**: 核心功能分离到独立模块，便于维护和测试
- **类型安全**: 完整的类型提示和数据验证
- **游戏优化**: 专门为游戏配置表比较优化的功能
- **性能优化**: 支持大文件处理和批量操作

## 联系信息

项目基于MIT许可证开源，支持Python 3.10+环境。
