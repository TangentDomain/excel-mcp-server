#!/usr/bin/env python3
"""
Excel MCP Server - 基于标准 MCP Python SDK 和 openpyxl 实现

使用低级MCP API实现的服务器，支持Excel文件操作

主要功能：
1. 正则搜索：在Excel文件中搜索符合正则表达式的单元格
2. 范围获取：读取指定范围的Excel数据
3. 范围修改：修改指定范围的Excel数据
4. 工作表管理：创建、删除、重命名工作表
5. 行列操作：插入、删除行列

技术栈：
- mcp: 标准MCP Python SDK
- openpyxl: 用于Excel文件操作
"""

import asyncio
import json
import logging
from typing import Optional, List, Dict, Any, Union

try:
    import mcp.server.stdio
    import mcp.types as types
    from mcp.server.lowlevel import NotificationOptions, Server
    from mcp.server.models import InitializationOptions
except ImportError as e:
    print(f"Error: 缺少必要的依赖包: {e}")
    print("请运行: pip install mcp openpyxl")
    exit(1)

# 导入核心模块
from .core.excel_reader import ExcelReader
from .core.excel_writer import ExcelWriter
from .core.excel_manager import ExcelManager
from .core.excel_search import ExcelSearcher
from .core.excel_compare import ExcelComparer

# 导入统一错误处理
from .utils.error_handler import unified_error_handler, extract_file_context, extract_formula_context

# 导入结果格式化工具
from .utils.formatter import format_operation_result

# ==================== 配置和初始化 ====================
# 开启详细日志用于调试
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger(__name__)

# 创建MCP服务器实例
server = Server("excel-mcp-server")


# ==================== 工具定义 ====================

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    """列出所有可用的Excel操作工具"""
    return [
        types.Tool(
            name="excel_list_sheets",
            description="列出Excel文件中所有工作表名称",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Excel文件路径 (.xlsx/.xlsm)"}
                },
                "required": ["file_path"]
            }
        ),
        types.Tool(
            name="excel_regex_search",
            description="在Excel文件中使用正则表达式搜索单元格内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Excel文件路径 (.xlsx/.xlsm)"},
                    "pattern": {"type": "string", "description": "正则表达式模式"},
                    "sheet_name": {"type": "string", "description": "工作表名称（可选）"},
                    "flags": {"type": "string", "description": "正则修饰符", "default": ""},
                    "search_values": {"type": "boolean", "description": "是否搜索单元格值", "default": True},
                    "search_formulas": {"type": "boolean", "description": "是否搜索公式内容", "default": False}
                },
                "required": ["file_path", "pattern"]
            }
        ),
        types.Tool(
            name="excel_get_range",
            description="读取Excel指定范围的数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Excel文件路径 (.xlsx/.xlsm)"},
                    "range_expression": {"type": "string", "description": "范围表达式，如 'Sheet1!A1:C10'"},
                    "include_formatting": {"type": "boolean", "description": "是否包含格式信息", "default": False}
                },
                "required": ["file_path", "range_expression"]
            }
        ),
        types.Tool(
            name="excel_update_range",
            description="更新Excel指定范围的数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Excel文件路径 (.xlsx/.xlsm)"},
                    "range_expression": {"type": "string", "description": "范围表达式，如 'Sheet1!A1:C10'"},
                    "data": {"type": "array", "description": "二维数组数据", "items": {"type": "array"}},
                    "preserve_formulas": {"type": "boolean", "description": "保留已有公式", "default": True}
                },
                "required": ["file_path", "range_expression", "data"]
            }
        ),
        types.Tool(
            name="excel_create_file",
            description="创建新的Excel文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "新文件路径 (.xlsx/.xlsm)"},
                    "sheet_names": {"type": "array", "description": "工作表名称列表", "items": {"type": "string"}}
                },
                "required": ["file_path"]
            }
        )
        # 这里可以继续添加更多工具...
    ]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict[str, Any]) -> list[types.TextContent]:
    """处理工具调用请求"""
    try:
        if name == "excel_list_sheets":
            result = _excel_list_sheets(arguments["file_path"])
        elif name == "excel_regex_search":
            result = _excel_regex_search(
                file_path=arguments["file_path"],
                pattern=arguments["pattern"],
                sheet_name=arguments.get("sheet_name"),
                flags=arguments.get("flags", ""),
                search_values=arguments.get("search_values", True),
                search_formulas=arguments.get("search_formulas", False)
            )
        elif name == "excel_get_range":
            result = _excel_get_range(
                file_path=arguments["file_path"],
                range_expression=arguments["range_expression"],
                include_formatting=arguments.get("include_formatting", False)
            )
        elif name == "excel_update_range":
            result = _excel_update_range(
                file_path=arguments["file_path"],
                range_expression=arguments["range_expression"],
                data=arguments["data"],
                preserve_formulas=arguments.get("preserve_formulas", True)
            )
        elif name == "excel_create_file":
            result = _excel_create_file(
                file_path=arguments["file_path"],
                sheet_names=arguments.get("sheet_names")
            )
        else:
            raise ValueError(f"Unknown tool: {name}")

        return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

    except Exception as e:
        logger.error(f"Tool {name} failed: {e}")
        return [types.TextContent(type="text", text=json.dumps({"error": str(e)}, ensure_ascii=False))]


# ==================== 工具实现函数 ====================

@unified_error_handler("列出工作表", extract_file_context, return_dict=True)
def _excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """
    列出Excel文件中所有工作表名称

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)

    Returns:
        Dict: 包含 success、sheets、active_sheet

    Example:
        # 列出工作表
        result = excel_list_sheets("data.xlsx")
        # 返回: {'success': True, 'sheets': ['Sheet1', 'Sheet2'], 'active_sheet': 'Sheet1'}
    """
    reader = ExcelReader(file_path)
    result = reader.list_sheets()

    # 提取工作表名称列表
    sheets = [sheet.name for sheet in result.data] if result.data else []

    return {
        'success': True,
        'sheets': sheets,
        'file_path': file_path,
        'total_sheets': result.metadata.get('total_sheets', len(sheets)) if result.metadata else len(sheets),
        'active_sheet': result.metadata.get('active_sheet', '') if result.metadata else ''
    }



@unified_error_handler("正则搜索", extract_file_context, return_dict=True)
def _excel_regex_search(
    file_path: str,
    pattern: str,
    sheet_name: Optional[str] = None,
    flags: str = "",
    search_values: bool = True,
    search_formulas: bool = False
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
        flags: 正则修饰符 ("i"忽略大小写, "m"多行, "s"点号匹配换行)
        search_values: 是否搜索单元格值
        search_formulas: 是否搜索公式内容

    Returns:
        Dict: 包含 success、matches(List[Dict])、match_count、searched_sheets

    Example:
        # 搜索所有工作表中的邮箱格式
        result = excel_regex_search("data.xlsx", r'\\w+@\\w+\\.\\w+', flags="i")
        # 搜索指定工作表中的数字
        result = excel_regex_search("data.xlsx", r'\\d+', sheet_name="Sheet1")
        # 搜索数字并包含公式
        result = excel_regex_search("data.xlsx", r'\\d+', search_formulas=True)
    """
    searcher = ExcelSearcher(file_path)
    result = searcher.regex_search(pattern, flags, search_values, search_formulas, sheet_name)
    return format_operation_result(result)



@unified_error_handler("目录正则搜索", extract_file_context, return_dict=True)
def _excel_regex_search_directory(
    directory_path: str,
    pattern: str,
    flags: str = "",
    search_values: bool = True,
    search_formulas: bool = False,
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
        flags: 正则修饰符 ("i"忽略大小写, "m"多行, "s"点号匹配换行)
        search_values: 是否搜索单元格值
        search_formulas: 是否搜索公式内容
        recursive: 是否递归搜索子目录
        file_extensions: 文件扩展名过滤，如[".xlsx", ".xlsm"]
        file_pattern: 文件名正则模式过滤
        max_files: 最大搜索文件数限制

    Returns:
        Dict: 包含 success、matches(List[Dict])、total_matches、searched_files

    Example:
        # 搜索目录中的邮箱格式
        result = excel_regex_search_directory("./data", r'\\w+@\\w+\\.\\w+', "i")
        # 搜索特定文件名模式
        result = excel_regex_search_directory("./reports", r'\\d+', file_pattern=r'.*销售.*')
    """
    # 直接调用ExcelSearcher的静态方法，避免创建需要文件路径的实例
    from .core.excel_search import ExcelSearcher
    result = ExcelSearcher.search_directory_static(
        directory_path, pattern, flags, search_values, search_formulas,
        recursive, file_extensions, file_pattern, max_files
    )
    return format_operation_result(result)



@unified_error_handler("范围数据读取", extract_file_context, return_dict=True)
def _excel_get_range(
    file_path: str,
    range_expression: str,
    sheet_name: Optional[str] = None,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
    读取Excel指定范围的数据

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        range_expression: 范围表达式，支持两种格式：
            - 包含工作表名: "Sheet1!A1:C10"、"TrSkill!A1:Z100"
            - 不包含工作表名: "A1:C10" (需要同时指定sheet_name参数)
        sheet_name: 工作表名称 (可选，当range_expression不包含工作表名时必需)
        include_formatting: 是否包含单元格格式

    Returns:
        Dict: 包含 success、data(List[List])、range_info

    Example:
        # 使用包含工作表名的范围表达式
        result = excel_get_range("data.xlsx", "Sheet1!A1:C10")
        # 使用分离的参数
        result = excel_get_range("data.xlsx", "A1:C10", sheet_name="Sheet1")
    """
    reader = ExcelReader(file_path)

    # 检查range_expression是否已包含工作表名
    if '!' in range_expression:
        # 已包含工作表名，直接使用
        result = reader.get_range(range_expression, include_formatting)
    else:
        # 不包含工作表名，需要sheet_name参数
        if not sheet_name:
            return {"success": False, "error": "当range_expression不包含工作表名时，必须提供sheet_name参数"}
        full_range_expression = f"{sheet_name}!{range_expression}"
        result = reader.get_range(full_range_expression, include_formatting)

    return format_operation_result(result)



@unified_error_handler("范围数据更新", extract_file_context, return_dict=True)
def _excel_update_range(
    file_path: str,
    range_expression: str,
    data: List[List[Any]],
    sheet_name: Optional[str] = None,
    preserve_formulas: bool = True
) -> Dict[str, Any]:
    """
    更新Excel指定范围的数据。操作会覆盖目标范围内的现有数据。

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        range_expression: 范围表达式，支持两种格式：
            - 包含工作表名: "Sheet1!A1:C10"、"TrSkill!A1:Z100"
            - 不包含工作表名: "A1:C10" (需要同时指定sheet_name参数)
        data: 二维数组数据 [[row1], [row2], ...]
        sheet_name: 工作表名称 (可选，当range_expression不包含工作表名时必需)
        preserve_formulas: 保留已有公式 (默认值: True)
            - True: 如果目标单元格包含公式，则保留公式不覆盖
            - False: 覆盖所有内容，包括公式

    Returns:
        Dict: 包含 success、updated_cells(int)、message

    Example:
        data = [["姓名", "年龄"], ["张三", 25]]
        result = excel_update_range("test.xlsx", "Sheet1!A1:B2", data)
        result = excel_update_range("test.xlsx", "A1:B2", data, sheet_name="Sheet1", preserve_formulas=False)
    """
    writer = ExcelWriter(file_path)

    # 检查range_expression是否已包含工作表名
    if '!' in range_expression:
        # 已包含工作表名，直接使用
        full_range_expression = range_expression
    else:
        # 不包含工作表名，需要sheet_name参数
        if not sheet_name:
            raise ValueError("当range_expression不包含工作表名时，必须提供sheet_name参数")
        full_range_expression = f"{sheet_name}!{range_expression}"

    result = writer.update_range(full_range_expression, data, preserve_formulas)
    return format_operation_result(result)



@unified_error_handler("插入行操作", extract_file_context, return_dict=True)
def _excel_insert_rows(
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
    writer = ExcelWriter(file_path)
    result = writer.insert_rows(sheet_name, row_index, count)
    return format_operation_result(result)



@unified_error_handler("插入列操作", extract_file_context, return_dict=True)
def _excel_insert_columns(
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
    writer = ExcelWriter(file_path)
    result = writer.insert_columns(sheet_name, column_index, count)
    return format_operation_result(result)



@unified_error_handler("文件创建", extract_file_context, return_dict=True)
def _excel_create_file(
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
    result = ExcelManager.create_file(file_path, sheet_names)
    return format_operation_result(result)



@unified_error_handler("创建工作表", extract_file_context, return_dict=True)
def _excel_create_sheet(
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
    manager = ExcelManager(file_path)
    result = manager.create_sheet(sheet_name, index)
    return format_operation_result(result)



@unified_error_handler("删除工作表", extract_file_context, return_dict=True)
def _excel_delete_sheet(
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
    manager = ExcelManager(file_path)
    result = manager.delete_sheet(sheet_name)
    return format_operation_result(result)



@unified_error_handler("重命名工作表", extract_file_context, return_dict=True)
def _excel_rename_sheet(
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
    manager = ExcelManager(file_path)
    result = manager.rename_sheet(old_name, new_name)
    return format_operation_result(result)



@unified_error_handler("删除行操作", extract_file_context, return_dict=True)
def _excel_delete_rows(
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
    writer = ExcelWriter(file_path)
    result = writer.delete_rows(sheet_name, row_index, count)
    return format_operation_result(result)



@unified_error_handler("删除列操作", extract_file_context, return_dict=True)
def _excel_delete_columns(
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
    writer = ExcelWriter(file_path)
    result = writer.delete_columns(sheet_name, column_index, count)
    return format_operation_result(result)


#
@unified_error_handler("设置公式", extract_file_context, return_dict=True)
def _excel_set_formula(
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
    writer = ExcelWriter(file_path)
    result = writer.set_formula(cell_address, formula, sheet_name)
    return format_operation_result(result)


#
@unified_error_handler("公式计算", extract_formula_context, return_dict=True)
def _excel_evaluate_formula(
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
    writer = ExcelWriter(file_path)
    result = writer.evaluate_formula(formula, context_sheet)
    return format_operation_result(result)



@unified_error_handler("单元格格式化", extract_file_context, return_dict=True)
def _excel_format_cells(
    file_path: str,
    sheet_name: str,
    range_expression: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None
) -> Dict[str, Any]:
    """
    设置单元格格式（字体、颜色、对齐等）。formatting 和 preset 必须至少指定一个。

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range_expression: 目标范围 (如"A1:C10")
        formatting: 自定义格式配置字典：
            - font: {'name': '宋体', 'size': 12, 'bold': True, 'color': 'FF0000'}
            - fill: {'color': 'FFFF00'}
            - alignment: {'horizontal': 'center', 'vertical': 'center'}
        preset: 预设样式，可选值: "title", "header", "data", "highlight", "currency"

    Returns:
        Dict: 包含 success、formatted_count、message

    Example:
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1", preset="title")
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1", formatting={'font': {'bold': True}})
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1", preset="header", formatting={'font': {'size': 14}})
    """
    # 预设样式模板
    PRESETS = {
        "title": {
            'font': {'name': '微软雅黑', 'size': 16, 'bold': True, 'color': 'FFFFFF'},
            'fill': {'color': '4472C4'},
            'alignment': {'horizontal': 'center', 'vertical': 'center'}
        },
        "header": {
            'font': {'name': '微软雅黑', 'size': 12, 'bold': True, 'color': '000000'},
            'fill': {'color': 'D9E1F2'},
            'alignment': {'horizontal': 'center', 'vertical': 'center'}
        },
        "data": {
            'font': {'name': '宋体', 'size': 11, 'color': '000000'},
            'alignment': {'horizontal': 'left', 'vertical': 'center'}
        },
        "highlight": {
            'font': {'bold': True, 'color': '000000'},
            'fill': {'color': 'FFFF00'}
        },
        "currency": {
            'font': {'name': '宋体', 'size': 11, 'color': '000000'},
            'alignment': {'horizontal': 'right', 'vertical': 'center'}
        }
    }

    # 构建最终格式配置
    final_formatting = {}

    # 1. 如果有预设，先应用预设
    if preset:
        if preset not in PRESETS:
            return {"success": False, "error": f"未知的预设样式: {preset}。可选值: {list(PRESETS.keys())}"}
        final_formatting = PRESETS[preset].copy()

    # 2. 如果有自定义格式，合并到最终配置（覆盖预设）
    if formatting:
        for key, value in formatting.items():
            if key in final_formatting and isinstance(final_formatting[key], dict) and isinstance(value, dict):
                # 深度合并字典类型的格式设置
                final_formatting[key].update(value)
            else:
                final_formatting[key] = value

    # 3. 如果既没有预设也没有自定义格式，返回错误
    if not final_formatting:
        return {"success": False, "error": "必须指定 formatting 或 preset 参数中的至少一个"}

    writer = ExcelWriter(file_path)
    result = writer.format_cells(range_expression, final_formatting, sheet_name)
    return format_operation_result(result)


# ==================== Excel比较功能 ====================

#
@unified_error_handler("Excel文件比较", extract_file_context, return_dict=True)
def _excel_compare_files(
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
    # 游戏开发专用配置 - 直接创建固定配置
    from .models.types import ComparisonOptions
    from .core.excel_compare import ExcelComparer

    options = ComparisonOptions(
        compare_values=True,
        compare_formulas=False,
        compare_formats=False,
        ignore_empty_cells=True,
        case_sensitive=True,
        structured_comparison=True,
        header_row=header_row,
        id_column=id_column,
        show_numeric_changes=True,
        game_friendly_format=True,
        focus_on_id_changes=True
    )

    comparer = ExcelComparer(options)
    result = comparer.compare_files(file1_path, file2_path)
    return format_operation_result(result)

@unified_error_handler("Excel工作表比较", extract_file_context, return_dict=True)
def _excel_compare_sheets(
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
    # 游戏开发专用配置 - 直接创建固定配置
    from .models.types import ComparisonOptions
    from .core.excel_compare import ExcelComparer

    options = ComparisonOptions(
        compare_values=True,
        compare_formulas=False,
        compare_formats=False,
        ignore_empty_cells=True,
        case_sensitive=True,
        structured_comparison=True,
        header_row=header_row,
        id_column=id_column,
        show_numeric_changes=True,
        game_friendly_format=True,
        focus_on_id_changes=True
    )

    comparer = ExcelComparer(options)
    result = comparer.compare_sheets(file1_path, sheet1_name, file2_path, sheet2_name)
    return format_operation_result(result)
# ==================== 主程序 ====================
async def run_server():
    """运行MCP服务器"""
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="excel-mcp-server",
                server_version="1.0.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    # 运行MCP服务器
    asyncio.run(run_server())
