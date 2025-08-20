# Excel MCP Server (FastMCP)

基于 FastMCP 和 openpyxl 实现的 Excel 操作 MCP 服务器，为 Claude Desktop 和其他 MCP 客户端提供强大的 Excel 文件操作能力。

## 功能特性

- **🔍 正则搜索**: 在Excel文件中使用正则表达式搜索单元格内容
- **📊 范围获取**: 读取指定范围的Excel数据，支持格式信息
- **✏️ 范围修改**: 修改指定范围的数据，保持公式和格式完整性
- **➕ 行列插入**: 在指定位置插入空白行或列
- **📋 工作表管理**: 列出所有工作表和相关信息

## 📋 环境要求

- **Python**: 3.10 或更高版本
- **操作系统**: Windows, macOS, Linux
- **内存**: 建议 512MB 以上可用内存
- **磁盘空间**: 至少 100MB 可用空间

## 🚀 快速开始

### 方式一：使用自动化脚本（推荐）

**Windows 用户：**

```powershell
# 1. 克隆项目
git clone <repository-url>
cd excel-mcp-server

# 2. 运行自动化部署脚本（如果存在）
# 注意：项目中包含 start.ps1 启动脚本
./start.ps1
```

### 方式二：手动安装

```bash
# 1. 克隆或下载项目
git clone <repository-url>
cd excel-mcp-server

# 2. 创建虚拟环境
python -m venv venv

# 3. 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/Mac:
# source venv/bin/activate

# 4. 安装依赖
pip install fastmcp openpyxl mcp

# 5. 验证安装
python server.py --help
```

## ⚙️ 配置说明

### Claude Desktop 配置

1. **找到 Claude Desktop 配置文件**：
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Linux**: `~/.config/claude/claude_desktop_config.json`

2. **添加 MCP 服务器配置**：

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["path/to/excel-mcp-server/server.py"],
      "cwd": "path/to/excel-mcp-server"
    }
  }
}
```

**配置示例（Windows）**：

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "D:/excel-mcp-server/venv/Scripts/python.exe",
      "args": ["D:/excel-mcp-server/server.py"],
      "cwd": "D:/excel-mcp-server"
    }
  }
}
```

### 其他 MCP 客户端配置

对于其他支持 MCP 协议的客户端，配置服务器的标准方式：

- **协议**: stdio
- **命令**: `python server.py`
- **工作目录**: 项目根目录

## 🎯 启动和运行

### 使用启动脚本（Windows）

项目提供了 `start.ps1` 自动化启动脚本：

```powershell
./start.ps1
```

启动脚本会自动：
1. 检查虚拟环境是否存在
2. 激活虚拟环境
3. 启动 MCP 服务器
4. 等待客户端连接

### 手动启动

```bash
# 激活虚拟环境
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

# 启动服务器
python server.py
```

### 验证运行状态

服务器启动成功后，你应该能看到：
- 服务器在 stdio 模式下运行
- 等待 MCP 客户端连接的提示信息
- 没有错误信息输出

## 📚 API 参考

### excel_regex_search
在Excel文件中搜索符合正则表达式的单元格
- `file_path`: Excel文件路径
- `pattern`: 正则表达式模式
- `flags`: 正则标志 (i=忽略大小写, m=多行, s=点匹配换行)
- `search_values`: 是否搜索显示值
- `search_formulas`: 是否搜索公式

### excel_list_sheets
列出Excel文件中所有工作表信息
- `file_path`: Excel文件路径

### excel_get_range
获取Excel文件指定范围的数据
- `file_path`: Excel文件路径
- `range_expression`: 范围表达式 (如 'A1:C10' 或 'Sheet1!A1:C10')
- `include_formatting`: 是否包含格式信息

### excel_update_range
修改Excel文件指定范围的数据
- `file_path`: Excel文件路径
- `range_expression`: 范围表达式
- `data`: 二维数据数组
- `preserve_formulas`: 是否保留现有公式

### excel_insert_rows
在Excel文件中插入空白行
- `file_path`: Excel文件路径
- `sheet_name`: 工作表名称 (可选)
- `row_index`: 插入位置（1-based）
- `count`: 插入行数 (最多1000行)

### excel_insert_columns
在Excel文件中插入空白列

- `file_path`: Excel文件路径
- `sheet_name`: 工作表名称 (可选)
- `column_index`: 插入位置（1-based）
- `count`: 插入列数 (最多100列)

## 💡 使用示例

### 1. 在 Claude Desktop 中使用

启动服务器并配置好 Claude Desktop 后，你可以直接与 Claude 对话：

```
# 对话示例
用户: "请帮我分析 D:/data/sales.xlsx 文件中包含邮箱地址的所有单元格"

Claude 会自动调用 excel_regex_search 工具来完成任务。
```

### 2. API 调用示例

#### 正则搜索示例

```python
# 搜索包含邮箱地址的单元格
result = excel_regex_search(
    file_path="example.xlsx",
    pattern=r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b",
    flags="i"
)
```

#### 范围获取示例

```python
# 获取A1:C10范围的数据
result = excel_get_range(
    file_path="example.xlsx",
    range_expression="Sheet1!A1:C10",
    include_formatting=True
)
```

#### 范围修改示例

```python
# 修改B2:D4范围的数据
result = excel_update_range(
    file_path="example.xlsx",
    range_expression="B2:D4",
    data=[
        ["新值1", "新值2", "新值3"],
        [100, 200, 300],
        ["=SUM(B2:C2)", "文本", 42]
    ]
)
```

### 3. 实际业务场景

#### 场景1：数据清理
```
使用场景：清理Excel文件中的重复数据和格式问题
1. 使用 excel_regex_search 查找格式异常的数据
2. 使用 excel_update_range 批量修正数据
3. 使用 excel_get_range 验证修改结果
```

#### 场景2：报表生成
```
使用场景：自动生成月度销售报表
1. 使用 excel_get_range 提取原始销售数据
2. 使用 excel_insert_rows 添加新的统计行
3. 使用 excel_update_range 填入计算结果
```

## 🛠️ 开发指南

### 运行测试

项目包含多个测试文件，用于验证功能：

```bash
# 基础 MCP 协议测试
python test_simple_mcp.py

# 完整 MCP 协议测试
python test_mcp_protocol.py

# 新功能测试
python test_new_features.py

# 新 API 测试
python test_new_apis.py
```

### 开发新功能

1. **添加新工具**：
   - 在 `server.py` 中定义新的工具函数
   - 使用 `@mcp.tool()` 装饰器注册
   - 添加适当的类型注解和文档字符串

2. **测试新功能**：
   - 创建对应的测试文件
   - 编写单元测试和集成测试

3. **更新文档**：
   - 更新 README 中的 API 参考部分
   - 添加使用示例

## ❓ 故障排除

### 常见问题

#### 1. 服务器无法启动
**症状**：运行 `python server.py` 时出现错误

**解决方案**：
```bash
# 检查 Python 版本
python --version  # 确保 >= 3.10

# 检查依赖安装
pip list | grep -E "(fastmcp|openpyxl|mcp)"

# 重新安装依赖
pip install --upgrade fastmcp openpyxl mcp
```

#### 2. Claude Desktop 无法连接
**症状**：Claude Desktop 中看不到 Excel 相关功能

**解决方案**：
1. 检查配置文件路径是否正确
2. 验证 JSON 配置语法
3. 重启 Claude Desktop
4. 检查服务器进程是否运行

#### 3. Excel 文件操作失败
**症状**：提示文件不存在或权限问题

**解决方案**：
- 确保文件路径使用绝对路径
- 检查文件是否被其他程序占用
- 验证文件格式是否为支持的类型（.xlsx, .xlsm, .xls）

#### 4. 虚拟环境问题
**症状**：依赖包找不到或版本冲突

**解决方案**：
```bash
# 删除虚拟环境重新创建
rm -rf venv
python -m venv venv
venv\Scripts\activate
pip install fastmcp openpyxl mcp
```

### 日志调试

启用详细日志输出：

```python
# 在 server.py 中修改日志级别
logging.basicConfig(level=logging.DEBUG)
```

### 获取帮助

- **GitHub Issues**: 报告 bug 或功能请求
- **文档**: 查看项目 README 和代码注释
- **测试文件**: 参考测试用例了解使用方法

## 📄 许可证

本项目采用 MIT 许可证，详见 LICENSE 文件。
    ]
)
```

## 技术实现

- 基于 **FastMCP** 框架，使用 `@mcp.tool()` 装饰器定义工具
- 使用 **openpyxl** 进行Excel文件操作，支持 .xlsx/.xlsm 格式
- 支持公式保护和格式保持
- 完整的错误处理和输入验证

## 配置

服务器运行在标准输入输出模式，可通过Claude Desktop或其他MCP客户端连接。
