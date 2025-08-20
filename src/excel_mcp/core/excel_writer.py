"""
Excel MCP Server - Excel写入模块

提供Excel文件写入和修改功能
"""

import logging
from typing import List, Any, Optional
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

from ..models.types import RangeInfo, ModifiedCell, OperationResult
from ..utils.validators import ExcelValidator
from ..utils.parsers import RangeParser
from ..utils.exceptions import SheetNotFoundError, DataValidationError

logger = logging.getLogger(__name__)


class ExcelWriter:
    """Excel文件写入器"""

    def __init__(self, file_path: str):
        """
        初始化Excel写入器

        Args:
            file_path: Excel文件路径
        """
        self.file_path = ExcelValidator.validate_file_path(file_path)

    def update_range(
        self,
        range_expression: str,
        data: List[List[Any]],
        preserve_formulas: bool = True
    ) -> OperationResult:
        """
        修改Excel文件中指定范围的数据

        Args:
            range_expression: 范围表达式
            data: 要写入的二维数据数组
            preserve_formulas: 是否保留现有的公式

        Returns:
            OperationResult: 修改操作的结果
        """
        try:
            # 解析范围表达式
            range_info = RangeParser.parse_range_expression(range_expression)

            # 加载Excel文件
            workbook = load_workbook(self.file_path)

            # 确定工作表
            sheet = self._get_worksheet(workbook, range_info.sheet_name)

            # 获取范围边界
            min_col, min_row, max_col, max_row = range_boundaries(range_info.cell_range)

            # 验证数据维度
            range_rows = max_row - min_row + 1
            range_cols = max_col - min_col + 1
            ExcelValidator.validate_range_data(data, range_rows, range_cols)

            # 写入数据
            modified_cells = self._write_data(
                sheet, data, min_row, min_col, preserve_formulas
            )

            # 保存文件
            workbook.save(self.file_path)

            return OperationResult(
                success=True,
                data=modified_cells,
                metadata={
                    'file_path': self.file_path,
                    'range': range_expression,
                    'sheet_name': sheet.title,
                    'modified_cells_count': len(modified_cells)
                }
            )

        except Exception as e:
            logger.error(f"更新范围数据失败: {e}")
            return OperationResult(
                success=False,
                error=str(e)
            )

    def insert_rows(
        self,
        sheet_name: Optional[str] = None,
        row_index: int = 1,
        count: int = 1
    ) -> OperationResult:
        """
        在Excel文件中插入空白行

        Args:
            sheet_name: 工作表名称
            row_index: 插入行的位置（1-based）
            count: 要插入的行数

        Returns:
            OperationResult: 插入操作的结果
        """
        try:
            # 验证参数
            ExcelValidator.validate_row_operations(row_index, count)
            ExcelValidator.validate_sheet_name(sheet_name)

            # 加载Excel文件
            workbook = load_workbook(self.file_path)

            # 确定工作表
            sheet = self._get_worksheet(workbook, sheet_name)

            # 记录操作前的信息
            original_max_row = sheet.max_row

            # 插入行
            sheet.insert_rows(row_index, count)

            # 保存文件
            workbook.save(self.file_path)

            return OperationResult(
                success=True,
                message=f"成功在第{row_index}行前插入{count}行",
                metadata={
                    'file_path': self.file_path,
                    'sheet_name': sheet.title,
                    'inserted_at_row': row_index,
                    'inserted_count': count,
                    'original_max_row': original_max_row,
                    'new_max_row': sheet.max_row
                }
            )

        except Exception as e:
            logger.error(f"插入行失败: {e}")
            return OperationResult(
                success=False,
                error=str(e)
            )

    def insert_columns(
        self,
        sheet_name: Optional[str] = None,
        column_index: int = 1,
        count: int = 1
    ) -> OperationResult:
        """
        在Excel文件中插入空白列

        Args:
            sheet_name: 工作表名称
            column_index: 插入列的位置（1-based）
            count: 要插入的列数

        Returns:
            OperationResult: 插入操作的结果
        """
        try:
            # 验证参数
            ExcelValidator.validate_column_operations(column_index, count)
            ExcelValidator.validate_sheet_name(sheet_name)

            # 加载Excel文件
            workbook = load_workbook(self.file_path)

            # 确定工作表
            sheet = self._get_worksheet(workbook, sheet_name)

            # 记录操作前的信息
            original_max_column = sheet.max_column

            # 插入列
            sheet.insert_cols(column_index, count)

            # 保存文件
            workbook.save(self.file_path)

            return OperationResult(
                success=True,
                message=f"成功插入{count}列",
                metadata={
                    'file_path': self.file_path,
                    'sheet_name': sheet.title,
                    'inserted_at_column': column_index,
                    'inserted_count': count,
                    'original_max_column': original_max_column,
                    'new_max_column': sheet.max_column
                }
            )

        except Exception as e:
            logger.error(f"插入列失败: {e}")
            return OperationResult(
                success=False,
                error=str(e)
            )

    def delete_rows(
        self,
        sheet_name: Optional[str] = None,
        start_row: int = 1,
        count: int = 1
    ) -> OperationResult:
        """
        删除Excel文件中的行

        Args:
            sheet_name: 工作表名称
            start_row: 开始删除的行号
            count: 要删除的行数

        Returns:
            OperationResult: 删除操作的结果
        """
        try:
            # 验证参数
            ExcelValidator.validate_row_operations(start_row, count)
            ExcelValidator.validate_sheet_name(sheet_name)

            # 加载Excel文件
            workbook = load_workbook(self.file_path)

            # 确定工作表
            sheet = self._get_worksheet(workbook, sheet_name)

            # 记录操作前的信息
            original_max_row = sheet.max_row

            # 验证删除范围
            if start_row > original_max_row:
                raise DataValidationError(
                    f"起始行号({start_row})超过工作表最大行数({original_max_row})"
                )

            # 计算实际删除的行数
            actual_count = min(count, original_max_row - start_row + 1)

            # 删除行
            sheet.delete_rows(start_row, actual_count)

            # 保存文件
            workbook.save(self.file_path)

            return OperationResult(
                success=True,
                message=f"成功从第{start_row}行开始删除{actual_count}行",
                metadata={
                    'file_path': self.file_path,
                    'sheet_name': sheet.title,
                    'deleted_start_row': start_row,
                    'actual_deleted_count': actual_count,
                    'original_max_row': original_max_row,
                    'new_max_row': sheet.max_row
                }
            )

        except Exception as e:
            logger.error(f"删除行失败: {e}")
            return OperationResult(
                success=False,
                error=str(e)
            )

    def delete_columns(
        self,
        sheet_name: Optional[str] = None,
        start_column: int = 1,
        count: int = 1
    ) -> OperationResult:
        """
        删除Excel文件中的列

        Args:
            sheet_name: 工作表名称
            start_column: 开始删除的列号
            count: 要删除的列数

        Returns:
            OperationResult: 删除操作的结果
        """
        try:
            # 验证参数
            ExcelValidator.validate_column_operations(start_column, count)
            ExcelValidator.validate_sheet_name(sheet_name)

            # 加载Excel文件
            workbook = load_workbook(self.file_path)

            # 确定工作表
            sheet = self._get_worksheet(workbook, sheet_name)

            # 记录操作前的信息
            original_max_column = sheet.max_column

            # 验证删除范围
            if start_column > original_max_column:
                raise DataValidationError(
                    f"起始列号({start_column})超过工作表最大列数({original_max_column})"
                )

            # 计算实际删除的列数
            actual_count = min(count, original_max_column - start_column + 1)

            # 删除列
            sheet.delete_cols(start_column, actual_count)

            # 保存文件
            workbook.save(self.file_path)

            return OperationResult(
                success=True,
                message=f"成功删除{actual_count}列",
                metadata={
                    'file_path': self.file_path,
                    'sheet_name': sheet.title,
                    'deleted_start_column': start_column,
                    'actual_deleted_count': actual_count,
                    'original_max_column': original_max_column,
                    'new_max_column': sheet.max_column
                }
            )

        except Exception as e:
            logger.error(f"删除列失败: {e}")
            return OperationResult(
                success=False,
                error=str(e)
            )

    def _get_worksheet(self, workbook, sheet_name: Optional[str]):
        """获取工作表"""
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise SheetNotFoundError(f"工作表不存在: {sheet_name}")
            return workbook[sheet_name]
        else:
            return workbook.active

    def _write_data(
        self,
        sheet,
        data: List[List[Any]],
        start_row: int,
        start_col: int,
        preserve_formulas: bool
    ) -> List[ModifiedCell]:
        """写入数据到工作表"""
        modified_cells = []

        for row_offset, row_data in enumerate(data):
            for col_offset, value in enumerate(row_data):
                row_idx = start_row + row_offset
                col_idx = start_col + col_offset
                cell = sheet.cell(row=row_idx, column=col_idx)

                # 保留公式检查
                if preserve_formulas and cell.data_type == 'f':
                    continue

                old_value = cell.value
                cell.value = value

                modified_cells.append(ModifiedCell(
                    coordinate=cell.coordinate,
                    old_value=old_value,
                    new_value=value
                ))

        return modified_cells
