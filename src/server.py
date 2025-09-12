#!/usr/bin/env python3
"""
Excel MCP Server - 基于 FastMCP 和 openpyxl 实现

重构后的服务器文件，只包含MCP接口定义，具体实现委托给核心模块

主要功能：
1. 正则搜索：在Excel文件中搜索符合正则表达式的单元格
2. 范围获取：读取指定范围的Excel数据
3. 范围修改：修改指定范围的Excel数据
4. 工作表管理：创建、删除、重命名工作表
5. 行列操作：插入、删除行列

技术栈：
- FastMCP: 用于MCP服务器框架
- openpyxl: 用于Excel文件操作
"""

import logging
from enum import Enum
from typing import Optional, List, Dict, Any, Union

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    print(f"Error: 缺少必要的依赖包: {e}")
    print("请运行: pip install fastmcp openpyxl")
    exit(1)

# 导入API模块
from .api.excel_operations import ExcelOperations

# ==================== 配置和初始化 ====================
# 开启详细日志用于调试
logging.basicConfig(
    level=logging.DEBUG,  # 改为DEBUG级别获取更多信息
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger(__name__)

# 创建FastMCP服务器实例，开启调试模式和详细日志
mcp = FastMCP(
    name="excel-mcp",
    instructions="""🎮 游戏开发Excel配置表专家 - 28个专业工具

## 🎯 核心设计原则
• **搜索优先**：任何查找、定位、分析操作都优先使用 `excel_search`
• **1-based索引**：第1行=1, 第1列=1 (匹配Excel惯例)
• **范围格式**：必须包含工作表名 `"技能配置表!A1:Z100"` `"装备配置表!B2:F50"`
• **ID驱动**：所有配置表以ID为主键，支持ID对象跟踪
• **中文友好**：完全支持中文工作表名和游戏术语

## ⚠️ 核心注意事项
🔴 **默认覆盖**：`excel_update_range`默认覆盖模式，需保留数据时用`insert_mode=True`
🔴 **操作验证**：更新前用`excel_get_range`预览，确保目标正确

## 🎮 游戏配置表专项操作

### 技能配置表常用操作
```
📋 技能表结构: ID|技能名|类型|等级|消耗|冷却|伤害|描述
🔍 查找技能: excel_search("skills.xlsx", r"火球|冰冻", "技能配置表")
📊 批量更新: excel_update_range("skills.xlsx", "技能配置表!G2:G100", damage_data)
🆚 版本对比: excel_compare_sheets("v1.xlsx", "技能配置表", "v2.xlsx", "技能配置表")
```

### 装备配置表操作模式
```
📦 装备配置: ID|名称|类型|品质|属性|套装|获取方式
🔧 属性调整: excel_get_range("items.xlsx", "装备配置表!E2:E200") → 分析 → 批量调整
🎨 品质标记: excel_format_cells("items.xlsx", "装备配置表", "D2:D200", preset="highlight")
```

### 怪物配置表管理
```
👹 怪物数据: ID|名称|等级|血量|攻击|防御|技能|掉落
📈 数值平衡: 使用excel_find_last_row定位 → 渐进式调整数值
🔄 AI行为: excel_search搜索特定AI模式进行批量调整
```

## 🚀 高效工作流程

### 标准配置表更新流程
1. **🔍 搜索定位**：`excel_search` → 了解数据分布和结构
2. **📏 确定边界**：`excel_find_last_row` → 确认数据范围
3. **📊 读取现状**：`excel_get_range` → 获取当前配置
4. **✏️ 更新数据**：`excel_update_range` → 覆盖写入（默认）
5. **🎨 美化显示**：`excel_format_cells` → 标记重要数据
6. **✅ 验证结果**：重新读取确认更新成功

### 版本对比工作流
```
🆚 配置对比流程:
excel_compare_sheets("old_config.xlsx", "技能配置表", "new_config.xlsx", "技能配置表")
↓ 分析差异报告
🆕 新增技能: 直接添加到新版本
🗑️ 删除技能: 检查依赖关系后移除
🔄 修改技能: 重点测试数值平衡
```

## 🛠️ 错误处理专家指南

### 常见问题快速解决
```
❌ 文件被锁定 → 检查Excel是否打开，关闭后重试
❌ 权限不足 → 使用管理员权限或检查文件属性
❌ 范围超界 → 先用excel_find_last_row确认实际数据范围
❌ 中文乱码 → 确认编码格式，使用utf-8
❌ 公式错误 → 设置preserve_formulas=False强制覆盖
❌ 内存不足 → 分批处理大文件，限制读取范围
```

### 复杂范围操作示例
```
📐 复杂范围支持:
单元格: "技能配置表!A1:Z100"    # 标准矩形范围
整行:   "装备配置表!5:10"        # 第5-10行
整列:   "怪物配置表!C:F"         # C到F列
单行:   "技能配置表!1"           # 仅第1行
单列:   "道具配置表!D"           # 仅D列
```

## ⚡ 性能优化要点
- **分批处理**：大文件分段操作，避免内存溢出
- **精确范围**：指定具体单元格范围，避免全表读取
- **批量操作**：一次性更新优于逐行处理

## 🎨 格式化预设

| 预设 | 用途 | 效果 |
|------|------|------|
| `"title"` | 标题行 | 粗体+居中 |
| `"header"` | 表头行 | 粗体+边框 |
| `"highlight"` | 重要数据 | 黄色高亮 |

## 🔍 智能搜索与分析

### 配置表数据挖掘
```
🔎 强大搜索能力:
excel_search("all_configs.xlsx", r"攻击力\s*[\d+]", regex_flags="i")     # 搜索攻击力数值
excel_search_directory("./configs", r"火|冰|雷", recursive=True)         # 批量搜索元素技能
excel_search("skills.xlsx", r"冷却.*[5-9]", include_formulas=True)      # 搜索长冷却技能
```

🚀 **游戏开发专家模式**: 搜索定位→数据分析→安全更新→视觉优化→版本对比→性能监控""",
    debug=True,                    # 开启调试模式
    log_level="DEBUG"              # 设置日志级别为DEBUG
)


# ==================== MCP 工具定义 ====================

@mcp.tool()
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """
    列出Excel文件中所有工作表名称

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)

    Returns:
        Dict: 包含success、sheets、total_sheets

    Example:
        # 列出工作表名称
        result = excel_list_sheets("data.xlsx")
        # 返回: {
        #   'success': True,
        #   'sheets': ['Sheet1', 'Sheet2'],
        #   'total_sheets': 2
        # }
    """
    return ExcelOperations.list_sheets(file_path)


@mcp.tool()
def excel_get_sheet_headers(file_path: str) -> Dict[str, Any]:
    """
    获取Excel文件中所有工作表的双行表头信息（游戏开发专用）

    这是 excel_get_headers 的便捷封装，用于批量获取所有工作表的双行表头。
    专为游戏配置表设计，同时获取字段描述和字段名。

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)

    Returns:
        Dict: 包含所有工作表的双行表头信息
        {
            'success': bool,
            'sheets_with_headers': [
                {
                    'name': str,
                    'headers': List[str],       # 字段名（兼容性）
                    'descriptions': List[str],  # 字段描述（第1行）
                    'field_names': List[str],   # 字段名（第2行）
                    'header_count': int
                }
            ],
            'file_path': str,
            'total_sheets': int
        }

    游戏配置表批量分析:
        一次性获取所有配置表的结构信息，包括字段描述和字段名，便于快速了解整个配置文件的结构。

    Example:
        # 获取游戏配置文件中所有表的双行表头
        result = excel_get_sheet_headers("game_config.xlsx")
        for sheet in result['sheets_with_headers']:
            print(f"表名: {sheet['name']}")
            print(f"字段描述: {sheet['descriptions']}")
            print(f"字段名: {sheet['field_names']}")
            print("---")

        # 返回示例: {
        #   'success': True,
        #   'sheets_with_headers': [
        #     {
        #       'name': '技能配置表',
        #       'headers': ['skill_id', 'skill_name', 'skill_type'],
        #       'descriptions': ['技能ID描述', '技能名称描述', '技能类型描述'],
        #       'field_names': ['skill_id', 'skill_name', 'skill_type'],
        #       'header_count': 3
        #     },
        #     {
        #       'name': '装备配置表',
        #       'headers': ['item_id', 'item_name', 'item_quality'],
        #       'descriptions': ['装备ID描述', '装备名称描述', '装备品质描述'],
        #       'field_names': ['item_id', 'item_name', 'item_quality'],
        #       'header_count': 3
        #     }
        #   ],
        #   'total_sheets': 2
        # }
    """
    return ExcelOperations.get_sheet_headers(file_path)


@mcp.tool()
def excel_search(
    file_path: str,
    pattern: str,
    sheet_name: Optional[str] = None,
    regex_flags: str = "",
    include_values: bool = True,
    include_formulas: bool = False,
    range: Optional[str] = None
) -> Dict[str, Any]:
    """
    在Excel文件中使用正则表达式搜索单元格内容

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        pattern: 正则表达式模式，支持常用格式：
            - r'\\d+': 匹配数字
            - r'\\w+@\\w+\\.\\w+': 匹配邮箱
            - r'^总计|合计$': 匹配特定文本
        sheet_name: 工作表名称 (可选，不指定时搜索所有工作表)
        regex_flags: 正则修饰符 ("i"忽略大小写, "m"多行, "s"点号匹配换行)
        include_values: 是否搜索单元格值
        include_formulas: 是否搜索公式内容
        range: 搜索范围表达式，支持多种格式：
            - 单元格范围: "A1:C10" 或 "Sheet1!A1:C10"
            - 行范围: "3:5" 或 "Sheet1!3:5" (第3行到第5行)
            - 列范围: "B:D" 或 "Sheet1!B:D" (B列到D列)
            - 单行: "7" 或 "Sheet1!7" (仅第7行)
            - 单列: "C" 或 "Sheet1!C" (仅C列)

    Returns:
        Dict: 包含 success、matches(List[Dict])、match_count、searched_sheets

    Example:
        # 搜索所有工作表中的邮箱格式
        result = excel_search("data.xlsx", r'\\w+@\\w+\\.\\w+', regex_flags="i")
        # 搜索指定工作表中的数字
        result = excel_search("data.xlsx", r'\\d+', sheet_name="Sheet1")
        # 搜索指定单元格范围内的数字
        result = excel_search("data.xlsx", r'\\d+', range="Sheet1!A1:C10")
        # 搜索第3-5行中的邮箱
        result = excel_search("data.xlsx", r'@', range="3:5", sheet_name="Sheet1")
        # 搜索B列到D列中的内容
        result = excel_search("data.xlsx", r'关键词', range="B:D", sheet_name="Sheet1")
        # 搜索单行或单列
        result = excel_search("data.xlsx", r'总计', range="10", sheet_name="Sheet1")  # 仅第10行
        result = excel_search("data.xlsx", r'金额', range="E", sheet_name="Sheet1")   # 仅E列
        # 搜索数字并包含公式
        result = excel_search("data.xlsx", r'\\d+', include_formulas=True)
    """
    return ExcelOperations.search(file_path, pattern, sheet_name, regex_flags, include_values, include_formulas, range)


@mcp.tool()
def excel_search_directory(
    directory_path: str,
    pattern: str,
    regex_flags: str = "",
    include_values: bool = True,
    include_formulas: bool = False,
    recursive: bool = True,
    file_extensions: Optional[List[str]] = None,
    file_pattern: Optional[str] = None,
    max_files: int = 100
) -> Dict[str, Any]:
    """
    在目录下的所有Excel文件中使用正则表达式搜索单元格内容

    Args:
        directory_path: 目录路径
        pattern: 正则表达式模式，支持常用格式：
            - r'\\d+': 匹配数字
            - r'\\w+@\\w+\\.\\w+': 匹配邮箱
            - r'^总计|合计$': 匹配特定文本
        regex_flags: 正则修饰符 ("i"忽略大小写, "m"多行, "s"点号匹配换行)
        include_values: 是否搜索单元格值
        include_formulas: 是否搜索公式内容
        recursive: 是否递归搜索子目录
        file_extensions: 文件扩展名过滤，如[".xlsx", ".xlsm"]
        file_pattern: 文件名正则模式过滤
        max_files: 最大搜索文件数限制

    Returns:
        Dict: 包含 success、matches(List[Dict])、total_matches、searched_files

    Example:
        # 搜索目录中的邮箱格式
        result = excel_search_directory("./data", r'\\w+@\\w+\\.\\w+', "i")
        # 搜索特定文件名模式
        result = excel_search_directory("./reports", r'\\d+', file_pattern=r'.*销售.*')
    """
    return ExcelOperations.search_directory(directory_path, pattern, regex_flags, include_values, include_formulas, recursive, file_extensions, file_pattern, max_files)


@mcp.tool()
def excel_get_range(
    file_path: str,
    range: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
    读取Excel指定范围的数据

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        range: 范围表达式，必须包含工作表名，支持格式：
            - 标准单元格范围: "Sheet1!A1:C10"、"TrSkill!A1:Z100"
            - 行范围: "Sheet1!1:1"、"数据!5:10"
            - 列范围: "Sheet1!A:C"、"统计!B:E"
            - 单行/单列: "Sheet1!5"、"数据!C"
        include_formatting: 是否包含单元格格式

    Returns:
        Dict: 包含 success、data(List[List])、range_info

    注意:
        为保持API一致性和清晰度，range必须包含工作表名。
        这消除了参数间的条件依赖，提高了可预测性。

    Example:
        # 读取单元格范围
        result = excel_get_range("data.xlsx", "Sheet1!A1:C10")
        # 读取整行
        result = excel_get_range("data.xlsx", "Sheet1!1:1")
        # 读取列范围
        result = excel_get_range("data.xlsx", "数据!A:C")
    """
    return ExcelOperations.get_range(file_path, range, include_formatting)


@mcp.tool()
def excel_get_headers(
    file_path: str,
    sheet_name: str,
    header_row: int = 1,
    max_columns: Optional[int] = None
) -> Dict[str, Any]:
    """
    获取Excel工作表的双行表头信息（游戏开发专用）

    专为游戏配置表设计，同时获取字段描述（第1行）和字段名（第2行）

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        header_row: 表头起始行号 (1-based，默认从第1行开始获取两行)
        max_columns: 最大读取列数限制 (可选)
            - 指定数值: 精确读取指定列数，如 max_columns=10 读取A-J列
            - None(默认): 读取前100列范围 (A-CV列)，然后截取到第一个空列

    Returns:
        Dict: 包含双行表头信息
        {
            'success': bool,
            'data': List[str],          # 字段名列表（兼容性）
            'headers': List[str],       # 字段名列表（兼容性）
            'descriptions': List[str],  # 字段描述列表（第1行）
            'field_names': List[str],   # 字段名列表（第2行）
            'header_count': int,
            'sheet_name': str,
            'header_row': int,
            'message': str
        }

    游戏配置表标准格式:
        第1行（descriptions）: ['技能ID描述', '技能名称描述', '技能类型描述', '技能等级描述']
        第2行（field_names）:   ['skill_id', 'skill_name', 'skill_type', 'skill_level']

    Example:
        # 获取技能配置表的双行表头
        result = excel_get_headers("skills.xlsx", "技能配置表")
        print(result['descriptions'])  # ['技能ID描述', '技能名称描述', ...]
        print(result['field_names'])   # ['skill_id', 'skill_name', ...]

        # 获取装备表第3-4行作为表头，精确读取8列
        result = excel_get_headers("items.xlsx", "装备配置表", header_row=3, max_columns=8)
    """
    return ExcelOperations.get_headers(file_path, sheet_name, header_row, max_columns)


@mcp.tool()
def excel_update_range(
    file_path: str,
    range: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = False
) -> Dict[str, Any]:
    """
更新Excel指定范围的数据。操作会覆盖目标范围内的现有数据。

Args:
    file_path: Excel文件路径 (.xlsx/.xlsm)
    range: 范围表达式，必须包含工作表名，支持格式：
        - 标准单元格范围: "Sheet1!A1:C10"、"TrSkill!A1:Z100"
        - 不支持行范围格式，必须使用明确单元格范围
    data: 二维数组数据 [[row1], [row2], ...]
    preserve_formulas: 保留已有公式 (默认值: True)
        - True: 如果目标单元格包含公式，则保留公式不覆盖
        - False: 覆盖所有内容，包括公式
    insert_mode: 数据写入模式 (默认值: False)
        - False: 覆盖模式，直接覆盖目标范围的现有数据（默认推荐）
        - True: 插入模式，在指定位置插入新行然后写入数据（更安全）

Returns:
    Dict: 包含 success、updated_cells(int)、message

注意:
    为保持API一致性和清晰度，range必须包含工作表名。
    这消除了参数间的条件依赖，提高了可预测性。

Example:
    data = [["姓名", "年龄"], ["张三", 25]]
    # 覆盖模式（默认）
    result = excel_update_range("test.xlsx", "Sheet1!A1:B2", data)
    # 插入模式（显式指定）
    result = excel_update_range("test.xlsx", "Sheet1!A1:B2", data, insert_mode=True)
    """
    return ExcelOperations.update_range(file_path, range, data, preserve_formulas, insert_mode)
@mcp.tool()
def excel_insert_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    在指定位置插入空行

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        row_index: 插入位置 (1-based，即第1行对应Excel中的第1行)
        count: 插入行数 (默认值: 1，即插入1行)

    Returns:
        Dict: 包含 success、inserted_rows、message

    Example:
        # 在第3行插入1行（使用默认count=1）
        result = excel_insert_rows("data.xlsx", "Sheet1", 3)
        # 在第5行插入3行（明确指定count）
        result = excel_insert_rows("data.xlsx", "Sheet1", 5, 3)
    """
    return ExcelOperations.insert_rows(file_path, sheet_name, row_index, count)


@mcp.tool()
def excel_insert_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    在指定位置插入空列

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        column_index: 插入位置 (1-based，即第1列对应Excel中的A列)
        count: 插入列数 (默认值: 1，即插入1列)

    Returns:
        Dict: 包含 success、inserted_columns、message

    Example:
        # 在第2列插入1列（使用默认count=1，即在B列前插入1列）
        result = excel_insert_columns("data.xlsx", "Sheet1", 2)
        # 在第1列插入2列（明确指定count，即在A列前插入2列）
        result = excel_insert_columns("data.xlsx", "Sheet1", 1, 2)
    """
    return ExcelOperations.insert_columns(file_path, sheet_name, column_index, count)


@mcp.tool()
def excel_find_last_row(
    file_path: str,
    sheet_name: str,
    column: Optional[Union[str, int]] = None
) -> Dict[str, Any]:
    """
    查找表格中最后一行有数据的位置

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        column: 指定列来查找最后一行（可选）
            - None: 查找整个工作表的最后一行
            - 整数: 列索引 (1-based，1=A列)
            - 字符串: 列名 (A, B, C...)

    Returns:
        Dict: 包含 success、last_row、message 等信息

    Example:
        # 查找整个工作表的最后一行
        result = excel_find_last_row("data.xlsx", "Sheet1")
        # 查找A列的最后一行有数据的位置
        result = excel_find_last_row("data.xlsx", "Sheet1", "A")
        # 查找第3列的最后一行有数据的位置
        result = excel_find_last_row("data.xlsx", "Sheet1", 3)
    """
    return ExcelOperations.find_last_row(file_path, sheet_name, column)


@mcp.tool()
def excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    创建新的Excel文件

    Args:
        file_path: 新文件路径 (必须以.xlsx或.xlsm结尾)
        sheet_names: 工作表名称列表 (默认值: None)
            - None: 创建包含一个默认工作表"Sheet1"的文件
            - []: 创建空的工作簿
            - ["名称1", "名称2"]: 创建包含指定名称工作表的文件

    Returns:
        Dict: 包含 success、file_path、sheets

    Example:
        # 创建简单文件（使用默认sheet_names=None，会有一个"Sheet1"）
        result = excel_create_file("new_file.xlsx")
        # 创建包含多个工作表的文件
        result = excel_create_file("report.xlsx", ["数据", "图表", "汇总"])
    """
    return ExcelOperations.create_file(file_path, sheet_names)


@mcp.tool()
def excel_export_to_csv(
    file_path: str,
    output_path: str,
    sheet_name: Optional[str] = None,
    encoding: str = "utf-8"
) -> Dict[str, Any]:
    """
    将Excel工作表导出为CSV文件

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        output_path: 输出CSV文件路径
        sheet_name: 工作表名称 (默认使用活动工作表)
        encoding: 文件编码 (默认: utf-8，可选: gbk)

    Returns:
        Dict: 包含 success、output_path、row_count、message

    Example:
        # 导出活动工作表为CSV
        result = excel_export_to_csv("data.xlsx", "output.csv")
        # 导出指定工作表
        result = excel_export_to_csv("report.xlsx", "summary.csv", "汇总", "gbk")
    """
    return ExcelOperations.export_to_csv(file_path, output_path, sheet_name, encoding)


@mcp.tool()
def excel_import_from_csv(
    csv_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    encoding: str = "utf-8",
    has_header: bool = True
) -> Dict[str, Any]:
    """
    从CSV文件导入数据创建Excel文件

    Args:
        csv_path: CSV文件路径
        output_path: 输出Excel文件路径
        sheet_name: 工作表名称 (默认: Sheet1)
        encoding: CSV文件编码 (默认: utf-8，可选: gbk)
        has_header: 是否包含表头行

    Returns:
        Dict: 包含 success、output_path、row_count、sheet_name

    Example:
        # 从CSV创建Excel文件
        result = excel_import_from_csv("data.csv", "output.xlsx")
        # 指定编码和工作表名
        result = excel_import_from_csv("sales.csv", "report.xlsx", "销售数据", "gbk")
    """
    return ExcelOperations.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header)


@mcp.tool()
def excel_convert_format(
    input_path: str,
    output_path: str,
    target_format: str = "xlsx"
) -> Dict[str, Any]:
    """
    转换Excel文件格式

    Args:
        input_path: 输入文件路径
        output_path: 输出文件路径
        target_format: 目标格式，可选值: "xlsx", "xlsm", "csv", "json"

    Returns:
        Dict: 包含 success、input_format、output_format、file_size

    Example:
        # 将xlsm转换为xlsx
        result = excel_convert_format("macro.xlsm", "data.xlsx", "xlsx")
        # 转换为JSON格式
        result = excel_convert_format("data.xlsx", "data.json", "json")
    """
    return ExcelOperations.convert_format(input_path, output_path, target_format)


@mcp.tool()
def excel_merge_files(
    input_files: List[str],
    output_path: str,
    merge_mode: str = "sheets"
) -> Dict[str, Any]:
    """
    合并多个Excel文件

    Args:
        input_files: 输入文件路径列表
        output_path: 输出文件路径
        merge_mode: 合并模式，可选值:
            - "sheets": 将每个文件作为独立工作表
            - "append": 将数据追加到单个工作表中
            - "horizontal": 水平合并（按列）

    Returns:
        Dict: 包含 success、merged_files、total_sheets、output_path

    Example:
        # 将多个文件合并为多个工作表
        files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
        result = excel_merge_files(files, "merged.xlsx", "sheets")

        # 将数据追加合并
        result = excel_merge_files(files, "combined.xlsx", "append")
    """
    return ExcelOperations.merge_files(input_files, output_path, merge_mode)


@mcp.tool()
def excel_get_file_info(file_path: str) -> Dict[str, Any]:
    """
    获取Excel文件的详细信息

    Args:
        file_path: Excel文件路径

    Returns:
        Dict: 包含文件信息，如大小、创建时间、工作表数量、格式等

    Example:
        # 获取文件详细信息
        result = excel_get_file_info("data.xlsx")
        # 返回: {
        #   'success': True,
        #   'file_size': 12345,
        #   'created_time': '2025-01-01 10:00:00',
        #   'modified_time': '2025-01-02 15:30:00',
        #   'format': 'xlsx',
        #   'sheet_count': 3,
        #   'has_macros': False
        # }
    """
    return ExcelOperations.get_file_info(file_path)


@mcp.tool()
def excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
    在文件中创建新工作表

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 新工作表名称 (不能与现有工作表重复)
        index: 插入位置 (0-based，默认值: None)
            - None: 在所有工作表的最后位置创建
            - 0: 在第一个位置创建
            - 1: 在第二个位置创建，以此类推

    Returns:
        Dict: 包含 success、sheet_name、total_sheets

    Example:
        # 创建新工作表到末尾（使用默认index=None）
        result = excel_create_sheet("data.xlsx", "新数据")
        # 创建新工作表到第一个位置（index=0）
        result = excel_create_sheet("data.xlsx", "首页", 0)
    """
    return ExcelOperations.create_sheet(file_path, sheet_name, index)


@mcp.tool()
def excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    删除指定工作表

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 要删除的工作表名称

    Returns:
        Dict: 包含 success、deleted_sheet、remaining_sheets

    Example:
        # 删除指定工作表
        result = excel_delete_sheet("data.xlsx", "临时数据")
    """
    return ExcelOperations.delete_sheet(file_path, sheet_name)


@mcp.tool()
def excel_rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]:
    """
    重命名工作表

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        old_name: 当前工作表名称
        new_name: 新工作表名称 (不能与现有重复)

    Returns:
        Dict: 包含 success、old_name、new_name

    Example:
        # 重命名工作表
        result = excel_rename_sheet("data.xlsx", "Sheet1", "主数据")
    """
    return ExcelOperations.rename_sheet(file_path, old_name, new_name)


@mcp.tool()
def excel_delete_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    删除指定行

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        row_index: 起始行号 (1-based，即第1行对应Excel中的第1行)
        count: 删除行数 (默认值: 1，即删除1行)

    Returns:
        Dict: 包含 success、deleted_rows、message

    Example:
        # 删除第5行（使用默认count=1）
        result = excel_delete_rows("data.xlsx", "Sheet1", 5)
        # 删除第3-5行（删除3行，从第3行开始）
        result = excel_delete_rows("data.xlsx", "Sheet1", 3, 3)
    """
    return ExcelOperations.delete_rows(file_path, sheet_name, row_index, count)


@mcp.tool()
def excel_delete_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
    删除指定列

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        column_index: 起始列号 (1-based，即第1列对应Excel中的A列)
        count: 删除列数 (默认值: 1，即删除1列)

    Returns:
        Dict: 包含 success、deleted_columns、message

    Example:
        # 删除第2列（使用默认count=1，即删除B列）
        result = excel_delete_columns("data.xlsx", "Sheet1", 2)
        # 删除第1-3列（删除3列，从A列开始删除A、B、C列）
        result = excel_delete_columns("data.xlsx", "Sheet1", 1, 3)
    """
    return ExcelOperations.delete_columns(file_path, sheet_name, column_index, count)

# 暂时注释掉, 以后可能会用到
# @mcp.tool()
def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """
    设置单元格公式

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        cell_address: 单元格地址 (如"A1")
        formula: Excel公式 (不包含等号)

    Returns:
        Dict: 包含 success、formula、calculated_value

    Example:
        # 设置求和公式
        result = excel_set_formula("data.xlsx", "Sheet1", "D10", "SUM(D1:D9)")
        # 设置平均值公式
        result = excel_set_formula("data.xlsx", "Sheet1", "E1", "AVERAGE(A1:A10)")
    """
    return ExcelOperations.set_formula(file_path, sheet_name, cell_address, formula)

# 暂时注释掉, 以后可能会用到
# @mcp.tool()
def excel_evaluate_formula(
    file_path: str,
    formula: str,
    context_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
    临时执行Excel公式并返回计算结果，不修改文件

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        formula: Excel公式 (不包含等号，如"SUM(A1:A10)")
        context_sheet: 公式执行的上下文工作表名称

    Returns:
        Dict: 包含 success、formula、result、result_type

    Example:
        # 计算A1:A10的和
        result = excel_evaluate_formula("data.xlsx", "SUM(A1:A10)")
        # 计算特定工作表的平均值
        result = excel_evaluate_formula("data.xlsx", "AVERAGE(B:B)", "Sheet1")
    """
    return ExcelOperations.evaluate_formula(formula, context_sheet)


@mcp.tool()
def excel_format_cells(
    file_path: str,
    sheet_name: str,
    range: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None
) -> Dict[str, Any]:
    """
    设置单元格格式（字体、颜色、对齐等）- 支持自定义和预设两种模式

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range: 目标范围 (如"A1:C10")
        formatting: 自定义格式配置字典（可选）：
            - font: {'name': '宋体', 'size': 12, 'bold': True, 'color': 'FF0000'}
            - fill: {'color': 'FFFF00'}
            - alignment: {'horizontal': 'center', 'vertical': 'center'}
        preset: 预设样式（可选），可选值: "title", "header", "data", "highlight", "currency"

    注意: formatting 和 preset 必须指定其中一个，如果同时指定，preset 优先

    Returns:
        Dict: 包含 success、formatted_count、message

    Example:
        # 使用预设样式
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1", preset="title")

        # 使用自定义格式
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1",
            formatting={'font': {'bold': True, 'color': 'FF0000'}})
    """
    return ExcelOperations.format_cells(file_path, sheet_name, range, formatting, preset)


@mcp.tool()
def excel_merge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
    合并指定范围的单元格

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range: 要合并的范围 (如"A1:C3")

    Returns:
        Dict: 包含 success、message、merged_range

    Example:
        # 合并A1:C3范围的单元格
        result = excel_merge_cells("data.xlsx", "Sheet1", "A1:C3")
        # 合并标题行
        result = excel_merge_cells("report.xlsx", "Summary", "A1:E1")
    """
    return ExcelOperations.merge_cells(file_path, sheet_name, range)


@mcp.tool()
def excel_unmerge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
    取消合并指定范围的单元格

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range: 要取消合并的范围 (如"A1:C3")

    Returns:
        Dict: 包含 success、message、unmerged_range

    Example:
        # 取消合并A1:C3范围的单元格
        result = excel_unmerge_cells("data.xlsx", "Sheet1", "A1:C3")
    """
    return ExcelOperations.unmerge_cells(file_path, sheet_name, range)


@mcp.tool()
def excel_set_borders(
    file_path: str,
    sheet_name: str,
    range: str,
    border_style: str = "thin"
) -> Dict[str, Any]:
    """
    为指定范围设置边框样式

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range: 目标范围 (如"A1:C10")
        border_style: 边框样式，可选值: "thin", "thick", "medium", "double", "dotted", "dashed"

    Returns:
        Dict: 包含 success、message、styled_range

    Example:
        # 为表格添加细边框
        result = excel_set_borders("data.xlsx", "Sheet1", "A1:E10", "thin")
        # 为标题添加粗边框
        result = excel_set_borders("data.xlsx", "Sheet1", "A1:E1", "thick")
    """
    return ExcelOperations.set_borders(file_path, sheet_name, range, border_style)


@mcp.tool()
def excel_set_row_height(
    file_path: str,
    sheet_name: str,
    row_index: int,
    height: float,
    count: int = 1
) -> Dict[str, Any]:
    """
    调整指定行的高度

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        row_index: 起始行号 (1-based)
        height: 行高 (磅值，如15.0)
        count: 调整行数 (默认值: 1)

    Returns:
        Dict: 包含 success、message、affected_rows

    Example:
        # 调整第1行高度为25磅
        result = excel_set_row_height("data.xlsx", "Sheet1", 1, 25.0)
        # 调整第2-4行高度为18磅
        result = excel_set_row_height("data.xlsx", "Sheet1", 2, 18.0, 3)
    """
    return ExcelOperations.set_row_height(file_path, sheet_name, row_index, height, count)


@mcp.tool()
def excel_set_column_width(
    file_path: str,
    sheet_name: str,
    column_index: int,
    width: float,
    count: int = 1
) -> Dict[str, Any]:
    """
    调整指定列的宽度

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        column_index: 起始列号 (1-based，1=A列)
        width: 列宽 (字符单位，如12.0)
        count: 调整列数 (默认值: 1)

    Returns:
        Dict: 包含 success、message、affected_columns

    Example:
        # 调整A列宽度为15字符
        result = excel_set_column_width("data.xlsx", "Sheet1", 1, 15.0)
        # 调整B-D列宽度为12字符
        result = excel_set_column_width("data.xlsx", "Sheet1", 2, 12.0, 3)
    """
    return ExcelOperations.set_column_width(file_path, sheet_name, column_index, width, count)


# ==================== Excel比较功能 ====================

# @mcp.tool()
def excel_compare_files(
    file1_path: str,
    file2_path: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """
    比较两个Excel文件 - 游戏开发专用版

    专注于ID对象的新增、删除、修改检测，自动识别配置表变化。

    Args:
        file1_path: 第一个Excel文件路径
        file2_path: 第二个Excel文件路径
        id_column: ID列位置（1-based数字或列名），默认第一列
        header_row: 表头行号（1-based），默认第一行

    Returns:
        Dict: 比较结果，包含新增、删除、修改的ID对象信息
        - 🆕 新增对象：ID在文件2中新出现
        - 🗑️ 删除对象：ID在文件1中存在但文件2中消失
        - 🔄 修改对象：ID存在于两文件中但属性发生变化
    """
    return ExcelOperations.compare_files(file1_path, file2_path)


@mcp.tool()
def excel_check_duplicate_ids(
    file_path: str,
    sheet_name: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """
    检查Excel工作表中ID列的重复值

    专为游戏配置表设计，快速识别ID重复问题，确保配置数据的唯一性。

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        id_column: ID列位置（1-based数字或列名），默认第一列
        header_row: 表头行号（1-based），默认第一行

    Returns:
        Dict: 查重结果
        {
            "success": true,
            "has_duplicates": true,
            "duplicate_count": 2,
            "total_ids": 100,
            "unique_ids": 98,
            "duplicates": [
                {
                    "id_value": "100001",
                    "count": 3,
                    "rows": [5, 15, 25]
                },
                {
                    "id_value": "100002",
                    "count": 2,
                    "rows": [8, 18]
                }
            ],
            "message": "发现2个重复ID，涉及5行数据"
        }

    Example:
        # 检查技能配置表ID重复
        result = excel_check_duplicate_ids("skills.xlsx", "技能配置表")
        # 检查装备表第2列ID重复
        result = excel_check_duplicate_ids("items.xlsx", "装备配置表", id_column=2)
    """
    return ExcelOperations.check_duplicate_ids(file_path, sheet_name, id_column, header_row)


@mcp.tool()
def excel_compare_sheets(
    file1_path: str,
    sheet1_name: str,
    file2_path: str,
    sheet2_name: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """
    比较两个Excel工作表，识别ID对象的新增、删除、修改。

    专为游戏配置表设计，使用紧凑数组格式提高传输效率。

    Args:
        file1_path: 第一个Excel文件路径
        sheet1_name: 第一个工作表名称
        file2_path: 第二个Excel文件路径
        sheet2_name: 第二个工作表名称
        id_column: ID列位置（1-based数字或列名），默认第一列
        header_row: 表头行号（1-based），默认第一行

    Returns:
        Dict: 比较结果
        {
            "success": true,
            "message": "成功比较工作表，发现3处差异",
            "data": {
                "sheet_name": "TrSkill vs TrSkill",
                "total_differences": 3,
                "row_differences": [
                    // 字段定义
                    ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"],

                    // 新增行
                    ["100001", "row_added", 0, 5, "TrSkill", null],

                    // 删除行
                    ["100002", "row_removed", 8, 0, "TrSkill", null],

                    // 修改行 - 包含变化的字段
                    ["100003", "row_modified", 10, 10, "TrSkill",
                        // field_differences: 变化的字段数组，每个元素格式 [字段名, 旧值, 新值, 变化类型]
                        [["技能名称", "火球术", "冰球术", "text_change"]]
                    ]
                ],
                "structural_changes": {
                    "max_row": {"sheet1": 100, "sheet2": 101, "difference": 1}
                }
            }
        }

    数据解析：
        row_differences[0] = 字段定义（索引说明）
        row_differences[1+] = 实际数据行

        对于row_modified类型：
        - field_differences: 变化的字段数组
          格式：[[字段名, 旧值, 新值, 变化类型], ...]
          变化类型："text_change" | "numeric_change" | "formula_change"

        对于row_added/row_removed类型：
        - field_differences为null，因为整行都是变化

    Example:
        result = excel_compare_sheets("old.xlsx", "Sheet1", "new.xlsx", "Sheet1")
        differences = result['data']['row_differences']
        for row in differences[1:]:  # 跳过字段定义行
            row_id, diff_type = row[0], row[1]
            print(f"{diff_type}: {row_id}")
    """
    return ExcelOperations.compare_sheets(file1_path, sheet1_name, file2_path, sheet2_name, id_column, header_row)
# ==================== 主程序 ====================
if __name__ == "__main__":
    # 运行FastMCP服务器
    mcp.run()
