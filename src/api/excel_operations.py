"""
Excel MCP Server - Excel操作API模块

提供高内聚的Excel业务操作功能，包含完整的参数验证、业务逻辑、错误处理和结果格式化

@intention: 将Excel操作的具体实现从server.py中分离，提高代码内聚性和可维护性
"""

import logging
from typing import Dict, Any, List, Optional

from ..core.excel_reader import ExcelReader
from ..core.excel_writer import ExcelWriter
from ..core.excel_manager import ExcelManager
from ..utils.formatter import format_operation_result

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

        except Exception as e:
            error_msg = f"获取范围数据失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def update_range(
        cls,
        file_path: str,
        range_expression: str,
        data: List[List[Any]],
        preserve_formulas: bool = True
    ) -> Dict[str, Any]:
        """
        @intention 更新Excel文件中指定范围的数据，确保数据完整性和公式保护

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            range_expression: 范围表达式，必须包含工作表名
            data: 二维数组数据 [[row1], [row2], ...]
            preserve_formulas: 是否保留现有公式

        Returns:
            Dict: 标准化的操作结果

        Example:
            data = [["姓名", "年龄"], ["张三", 25]]
            result = ExcelOperations.update_range("test.xlsx", "Sheet1!A1:B2", data)
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始更新范围数据: {range_expression}")

        try:
            # 步骤1: 验证参数格式
            validation_result = cls._validate_range_format(range_expression)
            if not validation_result['valid']:
                return cls._format_error_result(validation_result['error'])

            # 步骤2: 执行数据写入
            writer = ExcelWriter(file_path)
            result = writer.update_range(range_expression, data, preserve_formulas)

            # 步骤3: 格式化结果
            return format_operation_result(result)

        except Exception as e:
            error_msg = f"更新范围数据失败: {str(e)}"
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

        try:
            # 步骤1: 读取工作表信息
            reader = ExcelReader(file_path)
            result = reader.list_sheets()

            # 步骤2: 提取和格式化数据
            sheets = [sheet.name for sheet in result.data] if result.data else []

            response = {
                'success': True,
                'sheets': sheets,
                'file_path': file_path,
                'total_sheets': result.metadata.get('total_sheets', len(sheets)) if result.metadata else len(sheets),
                'active_sheet': result.metadata.get('active_sheet', '') if result.metadata else ''
            }

            # 步骤3: 清理资源
            reader.close()

            return response

        except Exception as e:
            error_msg = f"获取工作表列表失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    @classmethod
    def get_headers(
        cls,
        file_path: str,
        sheet_name: str,
        header_row: int = 1,
        max_columns: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        @intention 获取指定工作表的表头信息，支持智能截取和列数限制

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)
            sheet_name: 工作表名称
            header_row: 表头行号 (1-based)
            max_columns: 最大列数限制，None表示自动截取到空列

        Returns:
            Dict: 包含表头列表、数量等信息

        Example:
            result = ExcelOperations.get_headers("data.xlsx", "Sheet1")
        """
        if cls.DEBUG_LOG_ENABLED:
            logger.info(f"{cls._LOG_PREFIX} 开始获取表头: {sheet_name}")

        try:
            # 步骤1: 构建范围表达式
            range_expression = cls._build_header_range(sheet_name, header_row, max_columns)

            # 步骤2: 读取表头数据
            reader = ExcelReader(file_path)
            result = reader.get_range(range_expression)
            reader.close()

            if not result.success:
                return cls._format_error_result(f"无法读取表头数据: {result.message}")

            # 步骤3: 解析表头信息
            headers = cls._parse_header_data(result.data, max_columns)

            return {
                'success': True,
                'data': headers,
                'headers': headers,  # 兼容性字段
                'header_count': len(headers),
                'sheet_name': sheet_name,
                'header_row': header_row,
                'message': f"成功获取{len(headers)}个表头字段"
            }

        except Exception as e:
            error_msg = f"获取表头失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

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

        except Exception as e:
            error_msg = f"创建文件失败: {str(e)}"
            logger.error(f"{cls._LOG_PREFIX} {error_msg}")
            return cls._format_error_result(error_msg)

    # ==================== 分支实现 ====================

    # --- 参数验证 ---
    @classmethod
    def _validate_range_format(cls, range_expression: str) -> Dict[str, Any]:
        """验证范围表达式格式"""
        if not range_expression or not range_expression.strip():
            return {'valid': False, 'error': 'range参数不能为空'}

        if '!' not in range_expression:
            return {
                'valid': False,
                'error': f"range必须包含工作表名。当前格式: '{range_expression}'，正确格式示例: 'Sheet1!A1:B2'"
            }

        return {'valid': True}

    @classmethod
    def _build_header_range(cls, sheet_name: str, header_row: int, max_columns: Optional[int]) -> str:
        """构建表头范围表达式"""
        if max_columns:
            # 如果指定了最大列数，使用具体范围
            from openpyxl.utils import get_column_letter
            end_column = get_column_letter(max_columns)
            return f"{sheet_name}!A{header_row}:{end_column}{header_row}"
        else:
            # 否则使用一个合理的默认范围（读取前100列）
            return f"{sheet_name}!A{header_row}:CV{header_row}"  # CV = 第100列

    @classmethod
    def _parse_header_data(cls, data: List[List], max_columns: Optional[int]) -> List[str]:
        """解析表头数据"""
        headers = []
        if data and len(data) > 0:
            first_row = data[0]
            for i, cell_info in enumerate(first_row):
                # 处理CellInfo对象和普通值
                cell_value = getattr(cell_info, 'value', cell_info) if hasattr(cell_info, 'value') else cell_info

                # 转换为字符串并清理
                if cell_value is not None:
                    str_value = str(cell_value).strip()
                    if str_value != "":
                        headers.append(str_value)
                    else:
                        # 空字符串的处理
                        if max_columns:
                            headers.append("")  # 指定max_columns时保留空字符串
                        else:
                            break  # 否则停止
                else:
                    # None值的处理
                    if max_columns:
                        headers.append("")  # 指定max_columns时将None转为空字符串
                    else:
                        break  # 否则停止

                # 如果指定了max_columns，检查是否已达到限制
                if max_columns and len(headers) >= max_columns:
                    break

        return headers

    # --- 错误处理 ---
    @classmethod
    def _format_error_result(cls, error_message: str) -> Dict[str, Any]:
        """创建标准化的错误响应"""
        return {
            'success': False,
            'error': error_message,
            'data': None
        }
