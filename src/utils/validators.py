"""
Excel MCP Server - 验证工具

提供文件路径、参数等验证功能
"""

from pathlib import Path
from typing import Optional

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
