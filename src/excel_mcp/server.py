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

# ==================== 配置和初始化 ====================
logging.basicConfig(level=logging.WARNING)
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
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """
    获取Excel文件中所有工作表的名称列表

    Args:
        file_path: Excel文件路径

    Returns:
        包含所有工作表信息的字典
    """
    try:
        reader = ExcelReader(file_path)
        result = reader.list_sheets()
        return _format_result(result)
    except Exception as e:
        logger.error(f"列出工作表失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path
        }


@mcp.tool()
def excel_regex_search(
    file_path: str,
    pattern: str,
    flags: str = "",
    search_values: bool = True,
    search_formulas: bool = False
) -> Dict[str, Any]:
    """
    在Excel文件中使用正则表达式搜索单元格内容

    支持跨工作表搜索，可以同时搜索单元格值和公式内容。
    常用正则模式：
    - r'\\d+': 匹配数字
    - r'[A-Za-z]+': 匹配字母
    - r'\\b\\w+@\\w+\\.\\w+\\b': 匹配邮箱格式

    Args:
        file_path: Excel文件的绝对或相对路径，支持.xlsx和.xlsm格式
        pattern: 正则表达式模式字符串，使用Python re模块语法
        flags: 正则表达式修饰符，组合使用：
            - "i": 忽略大小写匹配
            - "m": 多行模式，^和$匹配每行的开始和结束
            - "s": 单行模式，点(.)匹配包括换行符的任意字符
            - 示例: "im" 表示忽略大小写且多行模式
        search_values: 是否在单元格的显示值中搜索（默认True）
        search_formulas: 是否在单元格的公式中搜索（默认False）

    Returns:
        搜索结果字典：
        - success (bool): 搜索是否成功完成
        - matches (List[Dict]): 匹配结果列表，每个匹配项包含：
            - coordinate (str): 单元格坐标，如"A1", "B5"
            - sheet_name (str): 所在工作表名称
            - value (Any): 单元格显示值
            - formula (str): 单元格公式（如果有）
            - matched_text (str): 实际匹配的文本
        - match_count (int): 总匹配数量
        - searched_sheets (List[str]): 已搜索的工作表名称列表
        - message (str): 操作成功信息
        - error (str): 错误描述（仅当success=False时存在）

    Raises:
        FileNotFoundError: Excel文件不存在或路径无效
        PermissionError: 文件被占用或无读取权限
        InvalidPatternError: 正则表达式语法错误
        UnsupportedFormatError: 不支持的文件格式

    Example:
        # 搜索所有包含邮箱的单元格
        result = excel_regex_search(
            file_path="contacts.xlsx",
            pattern=r"\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b",
            flags="i"
        )
    """
    try:
        searcher = ExcelSearcher(file_path)
        result = searcher.regex_search(pattern, flags, search_values, search_formulas)
        return _format_result(result)
    except Exception as e:
        logger.error(f"正则搜索失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'pattern': pattern
        }


@mcp.tool()
def excel_get_range(
    file_path: str,
    range_expression: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
    获取Excel文件中指定范围的数据
    支持多种范围模式：
    - 单元格范围: 'A1:C10' 或 'Sheet1!A1:C10'
    - 整行访问: '1:1' (第1行), '3:5' (第3-5行), '2' (仅第2行)
    - 整列访问: 'A:A' (A列), 'B:D' (B-D列), 'C' (仅C列)

    Args:
        file_path: Excel文件路径
        range_expression: 范围表达式
        include_formatting: 是否包含格式信息

    Returns:
        包含范围数据的字典
    """
    try:
        reader = ExcelReader(file_path)
        result = reader.get_range(range_expression, include_formatting)
        return _format_result(result)
    except Exception as e:
        logger.error(f"获取范围数据失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'range': range_expression
        }


@mcp.tool()
def excel_update_range(
    file_path: str,
    range_expression: str,
    data: List[List[Any]],
    preserve_formulas: bool = True
) -> Dict[str, Any]:
    """
    修改Excel文件中指定范围的数据

    Args:
        file_path: Excel文件路径
        range_expression: 范围表达式 (如 'A1:C10' 或 'Sheet1!A1:C10')
        data: 要写入的二维数据数组
        preserve_formulas: 是否保留现有的公式

    Returns:
        修改操作的结果信息
    """
    try:
        writer = ExcelWriter(file_path)
        result = writer.update_range(range_expression, data, preserve_formulas)
        return _format_result(result)
    except Exception as e:
        logger.error(f"更新范围数据失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'range': range_expression
        }


@mcp.tool()
def excel_insert_rows(
    file_path: str,
    sheet_name: Optional[str] = None,
    row_index: int = 1,
    count: int = 1
) -> Dict[str, Any]:
    """
    在Excel文件中插入空白行

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称 (可选，默认使用活动工作表)
        row_index: 插入行的位置（1-based），新行将插入到此位置之前
        count: 要插入的行数

    Returns:
        插入操作的结果信息
    """
    try:
        writer = ExcelWriter(file_path)
        result = writer.insert_rows(sheet_name, row_index, count)
        return _format_result(result)
    except Exception as e:
        logger.error(f"插入行失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name,
            'row_index': row_index,
            'count': count
        }


@mcp.tool()
def excel_insert_columns(
    file_path: str,
    sheet_name: Optional[str] = None,
    column_index: int = 1,
    count: int = 1
) -> Dict[str, Any]:
    """
    在Excel文件中插入空白列

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称 (可选，默认使用活动工作表)
        column_index: 插入列的位置（1-based），新列将插入到此位置之前
        count: 要插入的列数

    Returns:
        插入操作的结果信息
    """
    try:
        writer = ExcelWriter(file_path)
        result = writer.insert_columns(sheet_name, column_index, count)
        return _format_result(result)
    except Exception as e:
        logger.error(f"插入列失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name,
            'column_index': column_index,
            'count': count
        }


@mcp.tool()
def excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    创建新的Excel文件

    Args:
        file_path: 要创建的Excel文件路径
        sheet_names: 工作表名称列表 (可选，默认创建"Sheet1")

    Returns:
        创建操作的结果信息
    """
    try:
        result = ExcelManager.create_file(file_path, sheet_names)
        return _format_result(result)
    except Exception as e:
        logger.error(f"创建Excel文件失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_names': sheet_names
        }


@mcp.tool()
def excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
    在Excel文件中创建新工作表

    Args:
        file_path: Excel文件路径
        sheet_name: 新工作表名称
        index: 插入位置索引 (可选，默认在最后添加)

    Returns:
        创建操作的结果信息
    """
    try:
        manager = ExcelManager(file_path)
        result = manager.create_sheet(sheet_name, index)
        return _format_result(result)
    except Exception as e:
        logger.error(f"创建工作表失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name,
            'index': index
        }


@mcp.tool()
def excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    删除Excel文件中的工作表

    Args:
        file_path: Excel文件路径
        sheet_name: 要删除的工作表名称

    Returns:
        删除操作的结果信息
    """
    try:
        manager = ExcelManager(file_path)
        result = manager.delete_sheet(sheet_name)
        return _format_result(result)
    except Exception as e:
        logger.error(f"删除工作表失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name
        }


@mcp.tool()
def excel_rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]:
    """
    重命名Excel文件中的工作表

    Args:
        file_path: Excel文件路径
        old_name: 原工作表名称
        new_name: 新工作表名称

    Returns:
        重命名操作的结果信息
    """
    try:
        manager = ExcelManager(file_path)
        result = manager.rename_sheet(old_name, new_name)
        return _format_result(result)
    except Exception as e:
        logger.error(f"重命名工作表失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'old_name': old_name,
            'new_name': new_name
        }


@mcp.tool()
def excel_delete_rows(
    file_path: str,
    sheet_name: Optional[str] = None,
    start_row: int = 1,
    count: int = 1
) -> Dict[str, Any]:
    """
    在Excel文件中删除行

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称 (可选，默认使用活动工作表)
        start_row: 开始删除的行号（1-based）
        count: 要删除的行数

    Returns:
        删除操作的结果信息
    """
    try:
        writer = ExcelWriter(file_path)
        result = writer.delete_rows(sheet_name, start_row, count)
        return _format_result(result)
    except Exception as e:
        logger.error(f"删除行失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name,
            'start_row': start_row,
            'count': count
        }


@mcp.tool()
def excel_delete_columns(
    file_path: str,
    sheet_name: Optional[str] = None,
    start_column: int = 1,
    count: int = 1
) -> Dict[str, Any]:
    """
    在Excel文件中删除列

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称 (可选，默认使用活动工作表)
        start_column: 开始删除的列号（1-based）
        count: 要删除的列数

    Returns:
        删除操作的结果信息
    """
    try:
        writer = ExcelWriter(file_path)
        result = writer.delete_columns(sheet_name, start_column, count)
        return _format_result(result)
    except Exception as e:
        logger.error(f"删除列失败: {e}")
        return {
            'success': False,
            'error': str(e),
            'file_path': file_path,
            'sheet_name': sheet_name,
            'start_column': start_column,
            'count': count
        }


# ==================== 主程序 ====================
if __name__ == "__main__":
    # 运行FastMCP服务器
    mcp.run()
