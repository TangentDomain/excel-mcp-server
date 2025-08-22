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
    格式化操作结果为MCP响应格式

    Args:
        result: OperationResult对象

    Returns:
        格式化后的字典
    """
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
    sheet_name: str,
    range_expression: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
    读取Excel指定范围的数据

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称 (必选)
        range_expression: 范围表达式 (不需要包含工作表名)
            - 单元格: "A1:C10"
            - 整行: "1:5", "3" (单行)
            - 整列: "A:C", "B" (单列)
        include_formatting: 是否包含单元格格式

    Returns:
        Dict: 包含 success、data(List[List])、range_info

    Example:
        # 读取Sheet1中A1:C10范围数据
        result = excel_get_range("data.xlsx", "Sheet1", "A1:C10")
        # 包含格式的读取
        result = excel_get_range("data.xlsx", "Sheet1", "A1:C10", include_formatting=True)
    """
    reader = ExcelReader(file_path)
    # 组合sheet_name和range_expression
    full_range_expression = f"{sheet_name}!{range_expression}"
    result = reader.get_range(full_range_expression, include_formatting)
    return _format_result(result)


@mcp.tool()
@unified_error_handler("范围数据更新", extract_file_context, return_dict=True)
def excel_update_range(
    file_path: str,
    sheet_name: str,
    range_expression: str,
    data: List[List[Any]],
    preserve_formulas: bool = True
) -> Dict[str, Any]:
    """
    更新Excel指定范围的数据

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称 (必选)
        range_expression: 目标范围 (如"A1:C10", 不需要包含工作表名)
        data: 二维数组数据 [[row1], [row2], ...]
        preserve_formulas: 保留已有公式 (默认True)

    Returns:
        Dict: 包含 success、updated_cells(int)、message

    Example:
        # 更新Sheet1中A1:B2范围的数据
        data = [["姓名", "年龄"], ["张三", 25]]
        result = excel_update_range("test.xlsx", "Sheet1", "A1:B2", data)
    """
    writer = ExcelWriter(file_path)
    # 组合sheet_name和range_expression
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
        row_index: 插入位置 (1-based)
        count: 插入行数

    Returns:
        Dict: 包含 success、inserted_rows、message

    Example:
        # 在第3行插入1行
        result = excel_insert_rows("data.xlsx", "Sheet1", 3)
        # 在第5行插入3行
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
        column_index: 插入位置 (1-based)
        count: 插入列数

    Returns:
        Dict: 包含 success、inserted_columns、message

    Example:
        # 在第2列插入1列
        result = excel_insert_columns("data.xlsx", "Sheet1", 2)
        # 在第1列插入2列
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
        sheet_names: 工作表名称列表

    Returns:
        Dict: 包含 success、file_path、sheets

    Example:
        # 创建简单文件
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
        sheet_name: 新工作表名称 (不能与现有重复)
        index: 插入位置 (0-based，默认末尾)

    Returns:
        Dict: 包含 success、sheet_name、total_sheets

    Example:
        # 创建新工作表到末尾
        result = excel_create_sheet("data.xlsx", "新数据")
        # 创建新工作表到指定位置
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
        row_index: 起始行号 (1-based)
        count: 删除行数

    Returns:
        Dict: 包含 success、deleted_rows、message

    Example:
        # 删除第5行
        result = excel_delete_rows("data.xlsx", "Sheet1", 5)
        # 删除第3-5行(3行)
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
        column_index: 起始列号 (1-based)
        count: 删除列数

    Returns:
        Dict: 包含 success、deleted_columns、message

    Example:
        # 删除第2列
        result = excel_delete_columns("data.xlsx", "Sheet1", 2)
        # 删除第1-3列(3列)
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
    设置单元格格式（字体、颜色、对齐等）

    Args:
        file_path: Excel文件路径 (.xlsx/.xlsm)
        sheet_name: 工作表名称
        range_expression: 目标范围 (如"A1:C10")
        formatting: 自定义格式配置字典：
            - font: {'name': '宋体', 'size': 12, 'bold': True, 'color': 'FF0000'}
            - fill: {'color': 'FFFF00'}
            - alignment: {'horizontal': 'center', 'vertical': 'center'}
        preset: 预设样式 ("title", "header", "data", "highlight", "currency")

    Returns:
        Dict: 包含 success、formatted_count、message

    Example:
        # 使用预设样式
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1", preset="title")
        # 使用自定义格式
        result = excel_format_cells("data.xlsx", "Sheet1", "A1:D1",
                                  formatting={'font': {'bold': True, 'color': '000080'}})
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


# ==================== 主程序 ====================
if __name__ == "__main__":
    # 运行FastMCP服务器
    mcp.run()
