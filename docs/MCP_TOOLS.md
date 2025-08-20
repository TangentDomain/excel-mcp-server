# Excel MCP Server - 工具列表

## 📋 可用的MCP工具 (12个)

### 1. **excel_list_sheets** 📝
**功能**: 列出Excel文件中的所有工作表
```python
excel_list_sheets(file_path: str) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径

### 2. **excel_regex_search** 🔍
**功能**: 在工作表中使用正则表达式搜索内容
```python
excel_regex_search(
    file_path: str,
    pattern: str,
    sheet_name: Optional[str] = None,
    case_sensitive: bool = False
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `pattern`: 正则表达式模式
- `sheet_name`: 工作表名称（可选）
- `case_sensitive`: 是否区分大小写

### 3. **excel_get_range** 📊
**功能**: 获取指定范围的数据
```python
excel_get_range(
    file_path: str,
    range_expr: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `range_expr`: 范围表达式（如"A1:C10", "1:5", "A:C"）
- `sheet_name`: 工作表名称（可选）

### 4. **excel_update_range** ✏️
**功能**: 更新指定范围的数据
```python
excel_update_range(
    file_path: str,
    range_expr: str,
    data: List[List[Any]],
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `range_expr`: 范围表达式
- `data`: 要写入的数据（二维列表）
- `sheet_name`: 工作表名称（可选）

### 5. **excel_insert_rows** ➕
**功能**: 在指定位置插入行
```python
excel_insert_rows(
    file_path: str,
    row_index: int,
    count: int = 1,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `row_index`: 插入位置（1-based）
- `count`: 插入行数
- `sheet_name`: 工作表名称（可选）

### 6. **excel_insert_columns** ➕
**功能**: 在指定位置插入列
```python
excel_insert_columns(
    file_path: str,
    col_index: int,
    count: int = 1,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `col_index`: 插入位置（1-based）
- `count`: 插入列数
- `sheet_name`: 工作表名称（可选）

### 7. **excel_create_file** 🆕
**功能**: 创建新的Excel文件
```python
excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: 新文件路径
- `sheet_names`: 工作表名称列表（可选）

### 8. **excel_create_sheet** 📄
**功能**: 在现有文件中创建新工作表
```python
excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `sheet_name`: 新工作表名称
- `index`: 插入位置（可选）

### 9. **excel_delete_sheet** 🗑️
**功能**: 删除工作表
```python
excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `sheet_name`: 要删除的工作表名称

### 10. **excel_rename_sheet** 📝
**功能**: 重命名工作表
```python
excel_rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `old_name`: 原工作表名称
- `new_name`: 新工作表名称

### 11. **excel_delete_rows** ➖
**功能**: 删除指定行
```python
excel_delete_rows(
    file_path: str,
    row_index: int,
    count: int = 1,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `row_index`: 起始行索引（1-based）
- `count`: 删除行数
- `sheet_name`: 工作表名称（可选）

### 12. **excel_delete_columns** ➖
**功能**: 删除指定列
```python
excel_delete_columns(
    file_path: str,
    col_index: int,
    count: int = 1,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]
```
**参数**:
- `file_path`: Excel文件路径
- `col_index`: 起始列索引（1-based）
- `count`: 删除列数
- `sheet_name`: 工作表名称（可选）

## 🔧 使用示例

```bash
# 启动MCP服务器
uv run python src/excel_mcp/server_new.py

# 然后可以通过MCP客户端调用这些工具
```

## 📋 支持的范围表达式格式

- **单元格范围**: `A1:C10`, `Sheet1!A1:C10`
- **整行**: `1:5` (第1到5行)
- **整列**: `A:C` (A到C列)
- **单行**: `3` (第3行)
- **单列**: `B` (B列)

## ⚠️ 注意事项

1. 所有索引都是基于1的（1-based）
2. 文件路径支持相对路径和绝对路径
3. 支持的文件格式：`.xlsx`, `.xlsm`
4. 所有工具都包含完整的错误处理和验证
5. 返回结果统一为字典格式，包含`success`和`message`字段
