# Excel MCP Server (FastMCP)

基于 FastMCP 和 openpyxl 实现的 Excel 操作 MCP 服务器

## 功能特性

- **🔍 正则搜索**: 在Excel文件中使用正则表达式搜索单元格内容
- **📊 范围获取**: 读取指定范围的Excel数据，支持格式信息
- **✏️ 范围修改**: 修改指定范围的数据，保持公式和格式完整性
- **➕ 行列插入**: 在指定位置插入空白行或列
- **📋 工作表管理**: 列出所有工作表和相关信息

## 工具列表

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

## 安装

```bash
# 创建虚拟环境
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

# 安装依赖
pip install fastmcp openpyxl mcp

# 运行服务器
python server.py
```

## 使用示例

### 1. 正则搜索示例
```python
# 搜索包含邮箱地址的单元格
result = excel_regex_search(
    file_path="example.xlsx",
    pattern=r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b",
    flags="i"
)
```

### 2. 范围获取示例
```python
# 获取A1:C10范围的数据
result = excel_get_range(
    file_path="example.xlsx",
    range_expression="Sheet1!A1:C10",
    include_formatting=True
)
```

### 3. 范围修改示例
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

## 技术实现

- 基于 **FastMCP** 框架，使用 `@mcp.tool()` 装饰器定义工具
- 使用 **openpyxl** 进行Excel文件操作，支持 .xlsx/.xlsm 格式
- 支持公式保护和格式保持
- 完整的错误处理和输入验证

## 配置

服务器运行在标准输入输出模式，可通过Claude Desktop或其他MCP客户端连接。
