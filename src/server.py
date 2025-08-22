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

# 导入统一错误处理
from .utils.error_handler import unified_error_handler, extract_file_context, extract_formula_context

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

# 创建FastMCP服务器实例
mcp = FastMCP("excel-mcp-server")


# ==================== 辅助函数 ====================
def _format_result(result) -> Dict[str, Any]:
    """
    格式化操作结果为MCP响应格式，使用JSON序列化简化方案

    Args:
        result: OperationResult对象

    Returns:
        格式化后的字典，已清理null值，并转换为紧凑数组格式
    """
    import json

    def _convert_to_compact_array_format(data):
        """
        将结构化比较结果转换为紧凑的数组格式
        
        Args:
            data: StructuredDataComparison 数据对象
            
        Returns:
            转换后的紧凑格式数据
        """
        if not isinstance(data, dict) or 'row_differences' not in data:
            return data
            
        row_differences = data.get('row_differences', [])
        if not row_differences:
            return data
            
        # 检查是否已经是数组格式（避免重复转换）
        if (isinstance(row_differences, list) and 
            len(row_differences) > 0 and 
            isinstance(row_differences[0], list)):
            return data
            
        # 转换为紧凑数组格式
        compact_differences = []
        
        # 第一行：字段定义
        field_definitions = ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"]
        compact_differences.append(field_definitions)
        
        # 后续行：实际数据
        for diff in row_differences:
            if isinstance(diff, dict):
                # 转换字段级差异为数组格式
                field_diffs = diff.get('detailed_field_differences', [])
                compact_field_diffs = None
                
                if field_diffs:
                    compact_field_diffs = []
                    for field_diff in field_diffs:
                        if isinstance(field_diff, dict):
                            # 数组格式：[field_name, old_value, new_value, change_type]
                            compact_field_diffs.append([
                                field_diff.get('field_name', ''),
                                field_diff.get('old_value', ''),
                                field_diff.get('new_value', ''), 
                                field_diff.get('change_type', '')
                            ])
                
                # 主要差异数据数组：按字段定义顺序
                compact_row = [
                    diff.get('row_id', ''),
                    diff.get('difference_type', ''),
                    diff.get('row_index1', 0),
                    diff.get('row_index2', 0),
                    diff.get('sheet_name', ''),
                    compact_field_diffs
                ]
                compact_differences.append(compact_row)
        
        # 创建新的数据副本，替换row_differences
        new_data = data.copy()
        new_data['row_differences'] = compact_differences
        
        return new_data

    def _deep_clean_nulls(obj):
        """递归深度清理对象中的null/None值"""
        if isinstance(obj, dict):
            cleaned = {}
            for key, value in obj.items():
                if value is not None:
                    cleaned_value = _deep_clean_nulls(value)
                    if cleaned_value is not None and cleaned_value != {} and cleaned_value != []:
                        cleaned[key] = cleaned_value
            return cleaned
        elif isinstance(obj, list):
            cleaned = []
            for item in obj:
                if item is not None:
                    cleaned_item = _deep_clean_nulls(item)
                    if cleaned_item is not None and cleaned_item != {} and cleaned_item != []:
                        cleaned.append(cleaned_item)
            return cleaned
        else:
            return obj

    # 步骤1: 先转成JSON字符串（自动处理dataclass）
    try:
        def json_serializer(obj):
            """自定义JSON序列化器，专门处理dataclass和枚举"""
            if isinstance(obj, Enum):
                return obj.value
            elif hasattr(obj, '__dict__'):
                return obj.__dict__
            else:
                return str(obj)

        json_str = json.dumps(result, default=json_serializer, ensure_ascii=False)
        # 步骤2: 再转回字典
        result_dict = json.loads(json_str)
        
        # 步骤3: 转换为紧凑数组格式（仅用于结构化比较结果）
        if result_dict.get('data'):
            result_dict['data'] = _convert_to_compact_array_format(result_dict['data'])
        
        # 步骤4: 应用null清理
        cleaned_dict = _deep_clean_nulls(result_dict)
        return cleaned_dict
    except Exception as e:
        # 如果JSON方案失败，回退到原始方案
        response = {
            'success': result.success,
        }

        if result.success:
            if result.data is not None:
                # 处理数据类型转换
                if hasattr(result.data, '__dict__'):
                    # 如果是数据类，转换为字典
                    response.update(result.data.__dict__)
                elif isinstance(result.data, list):
                    # 如果是列表，处理每个元素
                    response['data'] = [
                        item.__dict__ if hasattr(item, '__dict__') else item
                        for item in result.data
                    ]
                else:
                    response['data'] = result.data

            if result.metadata:
                response.update(result.metadata)

            if result.message:
                response['message'] = result.message
        else:
            response['error'] = result.error

        return response


# ==================== MCP 工具定义 ====================

@mcp.tool()
@unified_error_handler("列出工作表", extract_file_context, return_dict=True)
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
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


@mcp.tool()
@unified_error_handler("正则搜索", extract_file_context, return_dict=True)
def excel_regex_search(
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
    return _format_result(result)


@mcp.tool()
@unified_error_handler("目录正则搜索", extract_file_context, return_dict=True)
def excel_regex_search_directory(
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
    return _format_result(result)


@mcp.tool()
@unified_error_handler("范围数据读取", extract_file_context, return_dict=True)
def excel_get_range(
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

    return _format_result(result)


@mcp.tool()
@unified_error_handler("范围数据更新", extract_file_context, return_dict=True)
def excel_update_range(
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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


@mcp.tool()
@unified_error_handler("文件创建", extract_file_context, return_dict=True)
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
    result = ExcelManager.create_file(file_path, sheet_names)
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


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
    return _format_result(result)


@mcp.tool()
@unified_error_handler("单元格格式化", extract_file_context, return_dict=True)
def excel_format_cells(
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
    return _format_result(result)


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
    return _format_result(result)
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
    比较两个Excel工作表 - 游戏开发专用版（紧凑数组格式）

    专注于ID对象的新增、删除、修改检测，自动识别配置表变化。
    
    ⚡ 优化特性：
    - 使用紧凑数组格式，减少60-80%的JSON大小
    - 避免大量重复的键名，提升传输和解析效率
    - 保持完整的比较信息，无数据丢失

    Args:
        file1_path: 第一个Excel文件路径
        sheet1_name: 第一个工作表名称
        file2_path: 第二个Excel文件路径
        sheet2_name: 第二个工作表名称
        id_column: ID列位置（1-based数字或列名），默认第一列
        header_row: 表头行号（1-based），默认第一行

    Returns:
        Dict: 比较结果（紧凑数组格式）
        {
            "success": bool,
            "message": str,
            "data": {
                "sheet_name": "Sheet1 vs Sheet2",
                "exists_in_file1": bool,
                "exists_in_file2": bool,
                "total_differences": int,
                
                // 🔥 核心优化：数组格式的差异数据
                "row_differences": [
                    // 第一行：字段定义（索引说明）
                    ["row_id", "difference_type", "row_index1", "row_index2", "sheet_name", "field_differences"],
                    
                    // 后续行：实际数据（按索引顺序）
                    ["18300504", "row_added", 0, 663, "TrSkillEffect", null],
                    ["11002101", "row_removed", 979, 0, "TrSkillEffect", null],
                    ["100000101", "row_modified", 987, 1000, "TrSkillEffect", [
                        // 字段差异也使用数组格式：[field_name, old_value, new_value, change_type]
                        ["初始技能增强ID列表", "", 183002041, "text_change"]
                    ]]
                ],
                
                "structural_changes": {
                    "max_row": {"sheet1": 988, "sheet2": 1001, "difference": 13},
                    "max_column": {"sheet1": 45, "sheet2": 41, "difference": -4}
                }
            },
            "metadata": {
                "file1": str,
                "sheet1": str, 
                "file2": str,
                "sheet2": str,
                "total_differences": int,
                "comparison_type": "structured"
            }
        }
        
        � 数据解析说明：
        - row_differences[0] 是字段定义，说明每列的含义
        - row_differences[1+] 是实际数据，按字段定义顺序排列
        - difference_type 值："row_added" | "row_removed" | "row_modified"
        - field_differences 格式：[[field_name, old_value, new_value, change_type], ...]
        - change_type 值："text_change" | "numeric_change" | "formula_change"
        
        🎯 优势对比：
        - 传统格式：每个差异约150-200字符的键名开销
        - 数组格式：仅需要6个数组索引，减少80%空间占用
        - 特别适合大型配置表比较（1000+行差异时效果显著）
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
    return _format_result(result)
# ==================== 主程序 ====================
if __name__ == "__main__":
    # 运行FastMCP服务器
    mcp.run()
