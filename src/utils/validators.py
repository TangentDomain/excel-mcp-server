"""
Excel MCP Server - 验证工具

提供文件路径、参数等验证功能
"""

import re
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

from .exceptions import ExcelFileNotFoundError, InvalidFormatError, DataValidationError


class ExcelValidator:
    """Excel操作验证器"""

    SUPPORTED_FORMATS = ['.xlsx', '.xlsm', '.xls']
    MAX_ROWS_OPERATION = 1000
    MAX_COLUMNS_OPERATION = 100

    @classmethod
    def validate_file_path(cls, file_path: str) -> str:
        """
        验证并规范化文件路径

        Args:
            file_path: Excel文件路径

        Returns:
            规范化的绝对路径

        Raises:
            ExcelFileNotFoundError: 文件不存在
            InvalidFormatError: 不支持的文件格式
        """
        path = Path(file_path)
        if not path.exists():
            raise ExcelFileNotFoundError(f"Excel文件不存在: {file_path}")

        if path.suffix.lower() not in cls.SUPPORTED_FORMATS:
            raise InvalidFormatError(f"不支持的文件格式: {path.suffix}")

        return str(path.absolute())

    @classmethod
    def validate_sheet_name(cls, sheet_name: Optional[str]) -> None:
        """
        验证工作表名称

        Args:
            sheet_name: 工作表名称

        Raises:
            DataValidationError: 工作表名称无效
        """
        if sheet_name is not None and not sheet_name.strip():
            raise DataValidationError("工作表名称不能为空")

    @classmethod
    def validate_row_operations(cls, row_index: int, count: int) -> None:
        """
        验证行操作参数

        Args:
            row_index: 行索引
            count: 操作数量

        Raises:
            DataValidationError: 参数无效
        """
        if row_index < 1:
            raise DataValidationError("行索引必须大于等于1")
        if count < 1:
            raise DataValidationError("操作行数必须大于等于1")
        if count > cls.MAX_ROWS_OPERATION:
            raise DataValidationError(f"一次最多操作{cls.MAX_ROWS_OPERATION}行")

    @classmethod
    def validate_column_operations(cls, column_index: int, count: int) -> None:
        """
        验证列操作参数

        Args:
            column_index: 列索引
            count: 操作数量

        Raises:
            DataValidationError: 参数无效
        """
        if column_index < 1:
            raise DataValidationError("列索引必须大于等于1")
        if count < 1:
            raise DataValidationError("操作列数必须大于等于1")
        if count > cls.MAX_COLUMNS_OPERATION:
            raise DataValidationError(f"一次最多操作{cls.MAX_COLUMNS_OPERATION}列")

    @classmethod
    def validate_range_data(cls, data: list, range_rows: int, range_cols: int) -> None:
        """
        验证范围数据

        Args:
            data: 数据数组
            range_rows: 范围行数
            range_cols: 范围列数

        Raises:
            DataValidationError: 数据维度不匹配
        """
        if len(data) > range_rows:
            raise DataValidationError(f"数据行数({len(data)})超过范围行数({range_rows})")

        for row_idx, row_data in enumerate(data):
            if len(row_data) > range_cols:
                raise DataValidationError(
                    f"第{row_idx + 1}行数据列数({len(row_data)})超过范围列数({range_cols})"
                )

    @classmethod
    def validate_file_for_creation(cls, file_path: str, overwrite: bool = True) -> str:
        """
        验证新建文件路径

        Args:
            file_path: 要创建的文件路径
            overwrite: 是否允许覆盖已存在的文件

        Returns:
            规范化的绝对路径

        Raises:
            FileExistsError: 文件已存在且不允许覆盖
            InvalidFormatError: 不支持的文件格式
        """
        path = Path(file_path)
        if path.exists() and not overwrite:
            raise FileExistsError(f"文件已存在: {file_path}")

        if path.suffix.lower() not in ['.xlsx', '.xlsm']:
            raise InvalidFormatError(
                f"不支持的文件格式: {path.suffix}，请使用 .xlsx 或 .xlsm"
            )

        return str(path.absolute())

    @classmethod
    def validate_range_expression(cls, range_expr: str) -> Dict[str, Any]:
        """
        严格验证范围表达式格式

        Args:
            range_expr: 范围表达式，必须包含工作表名

        Returns:
            验证结果和解析信息

        Raises:
            DataValidationError: 范围格式无效
        """
        if not range_expr or not isinstance(range_expr, str):
            raise DataValidationError("范围表达式不能为空且必须是字符串")

        range_expr = range_expr.strip()

        # 检查是否包含工作表名（必须包含感叹号）
        if '!' not in range_expr:
            raise DataValidationError(
                f"范围表达式必须包含工作表名和感叹号，格式示例: 'Sheet1!A1:C10'，当前: '{range_expr}'"
            )

        # 分离工作表名和范围
        parts = range_expr.split('!', 1)
        if len(parts) != 2:
            raise DataValidationError(
                f"范围表达式格式错误，应该只有一个感叹号，当前: '{range_expr}'"
            )

        sheet_name = parts[0].strip()
        range_part = parts[1].strip()

        # 验证工作表名
        if not sheet_name:
            raise DataValidationError("工作表名不能为空")

        # 检查工作表名中的无效字符
        invalid_chars = ['[', ']', '*', ':', '?', '/', '\\']
        for char in invalid_chars:
            if char in sheet_name:
                raise DataValidationError(f"工作表名包含无效字符: '{char}'")

        # 工作表名长度限制
        if len(sheet_name) > 31:
            raise DataValidationError("工作表名长度不能超过31个字符")

        # 验证范围部分
        if not range_part:
            raise DataValidationError("范围部分不能为空")

        # 解析范围部分
        range_info = cls._parse_range_part(range_part)

        return {
            'success': True,
            'sheet_name': sheet_name,
            'range_part': range_part,
            'range_info': range_info,
            'normalized_range': f"{sheet_name}!{range_part}"
        }

    @classmethod
    def _parse_range_part(cls, range_part: str) -> Dict[str, Any]:
        """
        解析范围部分，支持多种格式

        Args:
            range_part: 范围部分（不包含工作表名）

        Returns:
            范围解析信息
        """
        # 标准单元格范围: A1:C10
        range_pattern = r'^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$'
        match = re.match(range_pattern, range_part.upper())
        if match:
            start_cell, end_cell = match.groups()
            return {
                'type': 'cell_range',
                'start_cell': start_cell,
                'end_cell': end_cell,
                'start_col': cls._col_to_num(re.match(r'^([A-Z]+)', start_cell).group(1)),
                'start_row': int(re.match(r'[A-Z]+([0-9]+)$', start_cell).group(1)),
                'end_col': cls._col_to_num(re.match(r'^([A-Z]+)', end_cell).group(1)),
                'end_row': int(re.match(r'[A-Z]+([0-9]+)$', end_cell).group(1))
            }

        # 行范围: 1:10 或 5:5
        row_range_pattern = r'^([0-9]+):([0-9]+)$'
        match = re.match(row_range_pattern, range_part)
        if match:
            start_row, end_row = match.groups()
            return {
                'type': 'row_range',
                'start_row': int(start_row),
                'end_row': int(end_row)
            }

        # 列范围: A:C
        col_range_pattern = r'^([A-Z]+):([A-Z]+)$'
        match = re.match(col_range_pattern, range_part.upper())
        if match:
            start_col, end_col = match.groups()
            return {
                'type': 'column_range',
                'start_col': cls._col_to_num(start_col),
                'end_col': cls._col_to_num(end_col)
            }

        # 单行: 5
        single_row_pattern = r'^([0-9]+)$'
        match = re.match(single_row_pattern, range_part)
        if match:
            row = match.group(1)
            return {
                'type': 'single_row',
                'row': int(row)
            }

        # 单列: A
        single_col_pattern = r'^([A-Z]+)$'
        match = re.match(single_col_pattern, range_part.upper())
        if match:
            col = match.group(1)
            return {
                'type': 'single_column',
                'column': cls._col_to_num(col)
            }

        # 单个单元格: A1
        single_cell_pattern = r'^([A-Z]+[0-9]+)$'
        match = re.match(single_cell_pattern, range_part.upper())
        if match:
            cell = match.group(1)
            return {
                'type': 'single_cell',
                'column': cls._col_to_num(re.match(r'^([A-Z]+)', cell).group(1)),
                'row': int(re.match(r'[A-Z]+([0-9]+)$', cell).group(1))
            }

        # 如果没有匹配任何已知格式，抛出异常
        raise DataValidationError(f"无法识别的范围格式: '{range_part}'")

    @classmethod
    def _col_to_num(cls, col: str) -> int:
        """
        将列字母转换为数字（A=1, B=2, ..., Z=26, AA=27, etc.）

        Args:
            col: 列字母

        Returns:
            列数字
        """
        result = 0
        for i, char in enumerate(col):
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result

    @classmethod
    def validate_operation_scale(cls, range_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        验证操作规模，防止过大的操作影响性能

        Args:
            range_info: 范围信息

        Returns:
            规模评估结果
        """
        if 'start_row' in range_info and 'end_row' in range_info:
            rows = range_info['end_row'] - range_info['start_row'] + 1
        elif 'row' in range_info:
            rows = 1
        elif 'start_row' in range_info:
            rows = 1
        else:
            rows = 1

        if 'start_col' in range_info and 'end_col' in range_info:
            cols = range_info['end_col'] - range_info['start_col'] + 1
        elif 'column' in range_info:
            cols = 1
        elif 'start_col' in range_info:
            cols = 1
        else:
            cols = 1

        total_cells = rows * cols

        # 风险等级评估
        if total_cells > 10000:
            risk_level = "HIGH"
            warning = f"高风险操作：将影响 {total_cells:,} 个单元格"
        elif total_cells > 1000:
            risk_level = "MEDIUM"
            warning = f"中等风险操作：将影响 {total_cells:,} 个单元格"
        else:
            risk_level = "LOW"
            warning = None

        # 检查是否超过限制
        if rows > cls.MAX_ROWS_OPERATION:
            raise DataValidationError(
                f"操作行数({rows})超过限制({cls.MAX_ROWS_OPERATION})"
            )

        if cols > cls.MAX_COLUMNS_OPERATION:
            raise DataValidationError(
                f"操作列数({cols})超过限制({cls.MAX_COLUMNS_OPERATION})"
            )

        return {
            'rows': rows,
            'columns': cols,
            'total_cells': total_cells,
            'risk_level': risk_level,
            'warning': warning,
            'within_limits': True
        }
