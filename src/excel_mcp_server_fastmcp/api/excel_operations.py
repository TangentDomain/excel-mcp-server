"""
Excel MCP Server - Excel操作API模块

提供高内聚的Excel业务操作功能，包含完整的参数验证、业务逻辑、错误处理和结果格式化

@intention: 将Excel操作的具体实现从server.py中分离，提高代码内聚性和可维护性
"""

import logging
import re
from collections import Counter
from typing import Dict, Any, List, Optional, Union
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

from ..core.excel_reader import ExcelReader
from ..core.excel_writer import ExcelWriter
from ..core.excel_manager import ExcelManager
from ..core.excel_search import ExcelSearcher
from ..core.excel_compare import ExcelComparer
from ..core.excel_converter import ExcelConverter
from ..models.types import ComparisonOptions
from ..utils.formatter import format_operation_result
from ..utils.exceptions import (
    ExcelException,
    SheetNotFoundError,
    InvalidRangeError,
    DataValidationError,
    InvalidFormatError,
    OperationLimitError,
    ExcelFileNotFoundError,
    ExcelMCPError
)

logger = logging.getLogger(__name__)


class ExcelOperations:
    """
    @class ExcelOperations
    @brief Excel业务操作的高内聚封装
    @intention 提供完整的Excel操作功能，包含参数验证、错误处理、结果格式化
    """

    # ==================== 日志系统 ====================
    DEBUG_LOG_ENABLED: bool = False
    _LOG_PREFIX = '[API][ExcelOperations]'

    # ==================== 主干API ====================

    @classmethod
    def query(cls, file_path: str, sql: str, include_headers: bool = True) -> Dict[str, Any]:
        """
        @intention 执行SQL查询Excel数据

        Args:
            file_path: Excel文件路径
            sql: SQL查询语句
            include_headers: 是否包含表头

        Returns:
            Dict: 查询结果
        """
        try:
            from .advanced_sql_query import AdvancedSQLQueryEngine
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_sql_query(file_path, sql, include_headers=include_headers)
            return result
        except InvalidFormatError as e:
            return cls._format_error_result(f"无效的文件格式: {e.message}")
        except ExcelFileNotFoundError as e:
            return cls._format_error_result(f"文件不存在: {e.message}")
        except DataValidationError as e:
            return cls._format_error_result(f"数据验证失败: {e.message}")
        except ExcelMCPError as e:
            error_msg = f"SQL查询失败: {e.message}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def get_range(
        cls,
        file_path: str,
        range_expression: str,
        include_formatting: bool = False
    ) -> Dict[str, Any]:
        """
        @intention 获取Excel文件中指定范围的数据，提供完整的业务逻辑处理

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            range_expression: 范围表达式，必须包含工作表名
            include_formatting: 是否包含格式信息

        Returns:
            Dict: 标准化的操作结果

        Example:
            result = ExcelOperations.get_range("data.xlsx", "Sheet1!A1:C10")
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取范围数据: {range_expression}")

        try:
            # 步骤1: 验证参数格式
            validation_result = cls._validate_range_format(range_expression)
            if not validation_result['valid']:
                return cls._format_error_result(validation_result['error'])

            # 步骤2: 执行数据读取
            reader = ExcelReader(file_path)
            result = reader.get_range(range_expression, include_formatting)
            reader.close()

            # 步骤3: 格式化结果
            return format_operation_result(result)

        except InvalidFormatError as e:
            return cls._format_error_result(f"无效的文件格式: {e.message}")
        except ExcelFileNotFoundError as e:
            return cls._format_error_result(f"文件不存在: {e.message}")
        except SheetNotFoundError as e:
            return cls._format_error_result(f"工作表不存在: {e.message}")
        except InvalidRangeError as e:
            return cls._format_error_result(f"无效的范围表达式: {e.message}")
        except DataValidationError as e:
            return cls._format_error_result(f"数据验证失败: {e.message}")
        except ExcelMCPError as e:
            error_msg = f"获取范围数据失败: {e.message}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def update_range(
        cls,
        file_path: str,
        range_expression: str,
        data: List[List[Any]],
        preserve_formulas: bool = True,
        insert_mode: bool = True,
        streaming: bool = True
    ) -> Dict[str, Any]:
        """
        @intention 更新Excel文件中指定范围的数据，支持插入和覆盖模式

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            range_expression: 范围表达式，必须包含工作表名
            data: 二维数组数据 [[row1], [row2], ...]
            preserve_formulas: 是否保留现有公式
            insert_mode: 数据写入模式 (默认值: True)
                - True: 插入模式，在指定位置插入新行然后写入数据（更安全）
                - False: 覆盖模式，直接覆盖目标范围的现有数据
            streaming: 是否使用流式写入（仅覆盖模式有效）

        Returns:
            Dict: 标准化的操作结果
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始更新范围数据: {range_expression}, 模式: {'插入' if insert_mode else '覆盖'}")

        try:
            # 步骤1: 验证参数格式
            validation_result = cls._validate_range_format(range_expression)
            if not validation_result['valid']:
                return cls._format_error_result(validation_result['error'])

            # 步骤2: 扩展流式写入路径（支持覆盖+插入模式）
            if streaming:
                from excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
                if StreamingWriter.is_available():
                    try:
                        from excel_mcp_server_fastmcp.utils.parsers import RangeParser
                        range_info_parsed = RangeParser.parse_range_expression(range_expression)
                        sheet_name = range_info_parsed.sheet_name

                        if sheet_name:
                            # 解析起始行列
                            from openpyxl.utils import range_boundaries
                            min_col, min_row, max_col, max_row = range_boundaries(
                                range_info_parsed.cell_range
                            )
                            
                            # 根据模式选择不同的流式写入方法
                            if not insert_mode:
                                # 覆盖模式 - 直接使用流式写入
                                success, message, meta = StreamingWriter.update_range(
                                    file_path, sheet_name, min_row, min_col, data,
                                    preserve_formulas=preserve_formulas
                                )
                            else:
                                # 插入模式 - 使用流式插入
                                success, message, meta = StreamingWriter.insert_rows_streaming(
                                    file_path, sheet_name, min_row, data, 
                                    preserve_formulas=preserve_formulas
                                )
                            
                            if success:
                                return {
                                    'success': True,
                                    'message': message,
                                    'data': meta,
                                    'metadata': {
                                        'file_path': file_path,
                                        'sheet_name': sheet_name,
                                        'range': range_expression,
                                        'streaming_mode': 'insert' if insert_mode else 'overwrite',
                                        **meta
                                    }
                                }
                            else:
                                logger.warning(f"流式{ '插入' if insert_mode else '覆盖' }操作失败，降级到openpyxl: {message}")
                    except InvalidRangeError as parse_err:
                        logger.warning(f"流式update_range范围解析失败，降级到openpyxl: {parse_err.message}")

            # 步骤3: 传统openpyxl路径
            writer = ExcelWriter(file_path)
            result = writer.update_range(range_expression, data, preserve_formulas, insert_mode)

            # 步骤3: 格式化结果
            return format_operation_result(result)

        except InvalidFormatError as e:
            return cls._format_error_result(f"无效的文件格式: {e.message}")
        except ExcelFileNotFoundError as e:
            return cls._format_error_result(f"文件不存在: {e.message}")
        except SheetNotFoundError as e:
            return cls._format_error_result(f"工作表不存在: {e.message}")
        except InvalidRangeError as e:
            return cls._format_error_result(f"无效的范围表达式: {e.message}")
        except DataValidationError as e:
            return cls._format_error_result(f"数据验证失败: {e.message}")
        except ExcelMCPError as e:
            error_msg = f"更新范围数据失败: {e.message}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def list_sheets(cls, file_path: str) -> Dict[str, Any]:
        """
        @intention 获取Excel文件中所有工作表信息，提供完整的文件结构概览

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)

        Returns:
            Dict: 包含工作表列表、总数量、活动工作表等信息

        Example:
            result = ExcelOperations.list_sheets("data.xlsx")
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取工作表列表: {file_path}")

        reader = None
        try:
            # 步骤1: 读取工作表信息
            reader = ExcelReader(file_path)
            result = reader.list_sheets()

            # 步骤2: 提取和格式化数据
            sheets = [sheet.name for sheet in result.data] if result.data else []
            total_sheets = result.metadata.get('total_sheets', len(sheets)) if result.metadata else len(sheets)

            response = {
                'success': True,
                'message': f"获取到 {len(sheets)} 个工作表",
                'data': {                   # 统一data字段，集中返回核心数据
                    'sheets': sheets,
                    'total_sheets': total_sheets,
                    'file_path': file_path
                },
                'meta': {                   # meta字段保留扩展信息
                    'file_path': file_path,
                    'total_sheets': total_sheets
                },
                # 向后兼容快捷访问（数据来源data，避免重复存储）
                'sheets': sheets,
                'file_path': file_path,
                'total_sheets': total_sheets,
            }

            return response

        except InvalidFormatError as e:
            return cls._format_error_result(f"无效的文件格式: {e.message}")
        except ExcelFileNotFoundError as e:
            return cls._format_error_result(f"文件不存在: {e.message}")
        except ExcelMCPError as e:
            error_msg = f"获取工作表列表失败: {e.message}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)
        finally:
            # 步骤3: 确保清理资源，即使在异常情况下
            if reader is not None:
                try:
                    reader.close()
                except ExcelMCPError as cleanup_error:
                    logger.warning(f"{cls._LOG_PREFIX} 清理ExcelReader资源时发生错误: {cleanup_error.message}")

    @classmethod
    def get_headers(
        cls,
        file_path: str,
        sheet_name: str,
        header_row: int = 1,
        max_columns: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        @intention 获取指定工作表的表头信息，支持游戏开发双行模式（字段描述+字段名）

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            header_row: 表头起始行号 (1-based，默认从第1行开始获取两行)
            max_columns: 最大列数限制，None表示自动截取到空列

        Returns:
            Dict: 包含双行表头信息
            {
                'success': bool,
                'data': List[str],  # 字段名列表（兼容性）
                'headers': List[str],  # 字段名列表（兼容性）
                'descriptions': List[str],  # 字段描述列表（第1行）
                'field_names': List[str],   # 字段名列表（第2行）
                'header_count': int,
                'sheet_name': str,
                'header_row': int,
                'message': str
            }

        Example:
            result = ExcelOperations.get_headers("data.xlsx", "Sheet1")
            # 第1行：['技能ID描述', '技能名称描述', '技能类型描述']
            # 第2行：['skill_id', 'skill_name', 'skill_type']
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取双行表头: {sheet_name}")

        reader = None
        try:
            # 步骤1: 构建双行范围表达式
            range_expression = cls._build_header_range(sheet_name, header_row, max_columns, dual_row=True)

            # 步骤2: 读取表头数据（两行）
            reader = ExcelReader(file_path)
            result = reader.get_range(range_expression)

            if not result.success:
                return cls._format_error_result(f"无法读取表头数据: {result.message}")

            # 步骤3: 解析双行表头信息
            header_info = cls._parse_dual_header_data(result.data, max_columns)

            return {
                'success': True,
                # 统一data字段，集中核心数据
                'data': {
                    'field_names': header_info['field_names'],    # 字段名（第2行）
                    'descriptions': header_info['descriptions'],  # 字段描述（第1行）
                    'headers': header_info['field_names'],        # 向后兼容别名
                    'dual_rows': True  # 标识使用了双行模式
                },
                'meta': {
                    'sheet_name': sheet_name,
                    'header_row': header_row,
                    'header_count': len(header_info['field_names']),
                    'max_columns': max_columns,
                    'dual_row_mode': True
                },
                'message': f"成功获取{len(header_info['field_names'])}个表头字段（描述+字段名）",
                # 向后兼容快捷访问（数据来源data，避免重复存储）
                'headers': header_info['field_names'],
                'field_names': header_info['field_names'],
                'descriptions': header_info['descriptions'],
                'header_count': len(header_info['field_names']),
                'sheet_name': sheet_name,
                'header_row': header_row,
            }

        except ExcelMCPError as e:
            error_msg = f"获取表头失败: {e.message}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)
        finally:
            # 确保清理资源
            if reader is not None:
                try:
                    reader.close()
                except ExcelMCPError as cleanup_error:
                    logger.warning(f"{cls._LOG_PREFIX} 清理资源时发生错误: {cleanup_error.message}")

    @classmethod
    def create_file(
        cls,
        file_path: str,
        sheet_names: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        @intention 创建新的Excel文件，支持自定义工作表配置

        Args:
            file_path: 新文件路径 (必须以.xlsx或.xlsm结尾)
            sheet_names: 工作表名称列表，None表示默认工作表

        Returns:
            Dict: 包含创建结果和文件信息

        Example:
            result = ExcelOperations.create_file("new_file.xlsx", ["数据", "分析"])
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始创建文件: {file_path}")

        try:
            # 步骤1: 执行文件创建
            result = ExcelManager.create_file(file_path, sheet_names)

            # 步骤2: 格式化结果
            return format_operation_result(result)

        except ExcelMCPError as e:
            error_msg = f"创建文件失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

