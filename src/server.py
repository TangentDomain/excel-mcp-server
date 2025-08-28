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

# 导入核心模块
from .core.excel_reader import ExcelReader
from .core.excel_writer import ExcelWriter
from .core.excel_manager import ExcelManager
from .core.excel_search import ExcelSearcher
from .core.excel_compare import ExcelComparer

# 导入API模块
from .api.excel_operations import ExcelOperations

# 导入统一错误处理
from .utils.error_handler import unified_error_handler, extract_file_context, extract_formula_context

# 导入结果格式化工具
from .utils.formatter import format_operation_result

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
        Dict: 包含success、sheets、total_sheets、active_sheet

    Example:
        # 列出工作表名称
        result = excel_list_sheets("data.xlsx")
        # 返回: {
        #   'success': True,
        #   'sheets': ['Sheet1', 'Sheet2'],
        #   'total_sheets': 2,
        #   'active_sheet': 'Sheet1'
        # }
    """
    return ExcelOperations.list_sheets(file_path)


@mcp.tool()
@unified_error_handler("获取工作表表头", extract_file_context, return_dict=True)
def excel_get_sheet_headers(file_path: str) -> Dict[str, Any]:
    """
    获取Excel文件中所有工作表的表头信息

    这是 excel_get_headers 的便捷封装，用于批量获取所有工作表的表头。

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)

    Returns:
        Dict: 包含success、sheets_with_headers等信息

    Example:
        # 获取所有工作表的表头
        result = excel_get_sheet_headers("data.xlsx")
        # 返回: {
        #   'success': True,
        #   'sheets_with_headers': [
        #     {'name': 'Sheet1', 'headers': ['列1', '列2'], 'header_count': 2},
        #     {'name': 'Sheet2', 'headers': ['ID', '名称'], 'header_count': 2}
        #   ]
        # }
    """
    # 先获取所有工作表列表
    sheets_result = excel_list_sheets(file_path)
    if not sheets_result.get('success'):
        return sheets_result

    sheets_with_headers = []
    sheets = sheets_result.get('sheets', [])  # 修正字段名

    for sheet_name in sheets:
        try:
            # 使用统一的 excel_get_headers 方法获取每个工作表的表头
            header_result = excel_get_headers(file_path, sheet_name, header_row=1)

            if header_result.get('success'):
                # 兼容两种可能的数据格式
                headers = header_result.get('headers', [])
                if not headers and 'data' in header_result:
                    # 如果headers字段为空，尝试从data字段获取
                    headers = header_result.get('data', [])
                
                sheets_with_headers.append({
                    'name': sheet_name,
                    'headers': headers,
                    'header_count': len(headers)
                })
            else:
                # 如果读取某个工作表失败，记录错误但继续处理其他工作表
                sheets_with_headers.append({
                    'name': sheet_name,
                    'headers': [],
                    'header_count': 0,
                    'error': header_result.get('error', '未知错误')
                })

        except Exception as e:
            sheets_with_headers.append({
                'name': sheet_name,
                'headers': [],
                'header_count': 0,
                'error': str(e)
            })

    return format_operation_result({
        'success': True,
        'sheets_with_headers': sheets_with_headers,
        'file_path': file_path,
        'total_sheets': len(sheets)
    })


@mcp.tool()
@unified_error_handler("正则搜索", extract_file_context, return_dict=True)
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
    searcher = ExcelSearcher(file_path)
    result = searcher.regex_search(pattern, regex_flags, include_values, include_formulas, sheet_name, range)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("目录搜索", extract_file_context, return_dict=True)
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
        directory_path, pattern, regex_flags, include_values, include_formulas,
        recursive, file_extensions, file_pattern, max_files
    )
    return format_operation_result(result)


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
    获取Excel工作表的表头信息

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        header_row: 表头行号 (1-based，默认第1行)
        max_columns: 最大读取列数限制 (可选)
            - 指定数值: 精确读取指定列数，如 max_columns=10 读取A-J列
            - None(默认): 读取前100列范围 (A-CV列)，然后截取到第一个空列

    Returns:
        Dict: 包含 success、headers(List[str])、header_count、sheet_name

    注意:
        为保持与范围更新操作的一致性，方法内部使用明确的单元格范围而非行范围格式。
        当 max_columns=None 时，实际读取 A1:CV1 范围，然后自动截取到第一个空列。

    Example:
        # 获取第1行作为表头（自动截取到空列）
        result = excel_get_headers("data.xlsx", "Sheet1")
        # 获取第2行作为表头，精确读取10列
        result = excel_get_headers("data.xlsx", "Sheet1", header_row=2, max_columns=10)
        # 返回格式:
        # {
        #   'success': True,
        #   'headers': ['ID', '名称', '类型', '数量'],
        #   'header_count': 4,
        #   'sheet_name': 'Sheet1',
        #   'header_row': 1
        # }
    """
    return ExcelOperations.get_headers(file_path, sheet_name, header_row, max_columns)


@mcp.tool()
def excel_update_range(
    file_path: str,
    range: str,
    data: List[List[Any]],
    preserve_formulas: bool = True
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

    Returns:
        Dict: 包含 success、updated_cells(int)、message

    注意:
        为保持API一致性和清晰度，range必须包含工作表名。
        这消除了参数间的条件依赖，提高了可预测性。

    Example:
        data = [["姓名", "年龄"], ["张三", 25]]
        # 正确用法
        result = excel_update_range("test.xlsx", "Sheet1!A1:B2", data)
    """
    return ExcelOperations.update_range(file_path, range, data, preserve_formulas)


@mcp.tool()
@unified_error_handler("插入行操作", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.insert_rows(sheet_name, row_index, count)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("插入列操作", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.insert_columns(sheet_name, column_index, count)
    return format_operation_result(result)


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
@unified_error_handler("导出为CSV", extract_file_context, return_dict=True)
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
    from .core.excel_converter import ExcelConverter
    converter = ExcelConverter(file_path)
    result = converter.export_to_csv(output_path, sheet_name, encoding)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("从CSV导入", extract_file_context, return_dict=True)
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
    from .core.excel_converter import ExcelConverter
    result = ExcelConverter.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("文件格式转换", extract_file_context, return_dict=True)
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
    from .core.excel_converter import ExcelConverter
    result = ExcelConverter.convert_format(input_path, output_path, target_format)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("合并Excel文件", extract_file_context, return_dict=True)
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
    from .core.excel_converter import ExcelConverter
    result = ExcelConverter.merge_files(input_files, output_path, merge_mode)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("获取文件信息", extract_file_context, return_dict=True)
def excel_get_file_info(
    file_path: str
) -> Dict[str, Any]:
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
    from .core.excel_manager import ExcelManager
    result = ExcelManager.get_file_info(file_path)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("创建工作表", extract_file_context, return_dict=True)
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
    manager = ExcelManager(file_path)
    result = manager.create_sheet(sheet_name, index)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("删除工作表", extract_file_context, return_dict=True)
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
    manager = ExcelManager(file_path)
    result = manager.delete_sheet(sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("重命名工作表", extract_file_context, return_dict=True)
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
    manager = ExcelManager(file_path)
    result = manager.rename_sheet(old_name, new_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("删除行操作", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.delete_rows(sheet_name, row_index, count)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("删除列操作", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.delete_columns(sheet_name, column_index, count)
    return format_operation_result(result)


# @mcp.tool()
@unified_error_handler("设置公式", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.set_formula(cell_address, formula, sheet_name)
    return format_operation_result(result)


# @mcp.tool()
@unified_error_handler("公式计算", extract_formula_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.evaluate_formula(formula, context_sheet)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("单元格格式化", extract_file_context, return_dict=True)
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
    # 参数验证
    if not formatting and not preset:
        return format_operation_result({
            "success": False,
            "error": "必须指定 formatting（自定义格式）或 preset（预设样式）其中之一"
        })

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

    # 确定最终格式配置
    if preset:
        if preset not in PRESETS:
            return format_operation_result({
                "success": False,
                "error": f"未知的预设样式: {preset}。可选值: {list(PRESETS.keys())}"
            })
        final_formatting = PRESETS[preset]
    else:
        final_formatting = formatting

    writer = ExcelWriter(file_path)
    result = writer.format_cells(range, final_formatting, sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("合并单元格", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.merge_cells(range, sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("取消合并单元格", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.unmerge_cells(range, sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("设置边框样式", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.set_borders(range, border_style, sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("调整行高", extract_file_context, return_dict=True)
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
    writer = ExcelWriter(file_path)
    result = writer.set_row_height(row_index, height, sheet_name)
    return format_operation_result(result)


@mcp.tool()
@unified_error_handler("调整列宽", extract_file_context, return_dict=True)
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
    # 将列索引转换为列字母（1->A, 2->B, etc）
    from openpyxl.utils import get_column_letter
    column_letter = get_column_letter(column_index)

    writer = ExcelWriter(file_path)
    result = writer.set_column_width(column_letter, width, sheet_name)
    return format_operation_result(result)


# ==================== Excel比较功能 ====================

# @mcp.tool()
@unified_error_handler("Excel文件比较", extract_file_context, return_dict=True)
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
@mcp.tool()
@unified_error_handler("Excel工作表比较", extract_file_context, return_dict=True)
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
if __name__ == "__main__":
    # 运行FastMCP服务器
    mcp.run()
