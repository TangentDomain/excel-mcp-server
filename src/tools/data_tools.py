# 数据操作工具模块

from typing import Dict, Any, Optional
from ..api.excel_operations import ExcelOperations


def register_data_tools(mcp) -> None:
    """注册数据操作相关工具"""
    
    @mcp.tool()
    def excel_get_range(
        file_path: str,
        range: str,
        include_formatting: bool = False
    ) -> Dict[str, Any]:
        """读取指定范围的单元格数据
        
        Args:
            file_path: Excel文件路径
            range: 范围表达式 (如 "Sheet1!A1:C10")
            include_formatting: 是否包含格式
            
        Returns:
            {success, data, range_info}
        """
        from ..utils.validators import ExcelValidator, DataValidationError
        
        try:
            range_validation = ExcelValidator.validate_range_expression(range)
            scale_validation = ExcelValidator.validate_operation_scale(range_validation['range_info'])
        except DataValidationError as e:
            return {
                'success': False,
                'error': 'VALIDATION_FAILED',
                'message': f"范围表达式验证失败: {str(e)}"
            }
        
        return ExcelOperations.get_range(file_path, range, include_formatting)

    @mcp.tool()
    def excel_update_range(
        file_path: str,
        range: str,
        data: Any,
        preserve_formulas: bool = True
    ) -> Dict[str, Any]:
        """更新指定范围的单元格数据
        
        Args:
            file_path: Excel文件路径
            range: 范围表达式
            data: 要写入的数据（二维数组）
            preserve_formulas: 是否保留原有公式
            
        Returns:
            {success, updated_range, message}
        """
        return ExcelOperations.update_range(file_path, range, data, preserve_formulas)

    @mcp.tool()
    def excel_get_headers(
        file_path: str,
        sheet_name: str,
        row: int = 1,
        max_columns: int = 50
    ) -> Dict[str, Any]:
        """获取工作表的表头信息
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            row: 表头行号（默认第1行）
            max_columns: 最大列数
            
        Returns:
            {success, headers, row}
        """
        return ExcelOperations.get_headers(file_path, sheet_name, row, max_columns)

    @mcp.tool()
    def excel_find_last_row(
        file_path: str,
        sheet_name: Optional[str] = None,
        column: Optional[str] = None
    ) -> Dict[str, Any]:
        """查找工作表中最后一行数据的位置
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column: 指定列（字母或索引）
            
        Returns:
            {success, last_row, column}
        """
        return ExcelOperations.find_last_row(file_path, sheet_name, column)

    @mcp.tool()
    def excel_insert_rows(
        file_path: str,
        sheet_name: str,
        row: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """插入行
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            row: 插入位置
            count: 插入行数
            
        Returns:
            {success, inserted_at, count}
        """
        return ExcelOperations.insert_rows(file_path, sheet_name, row, count)

    @mcp.tool()
    def excel_delete_rows(
        file_path: str,
        sheet_name: str,
        row: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """删除行
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            row: 删除起始位置
            count: 删除行数
            
        Returns:
            {success, deleted_from, count}
        """
        return ExcelOperations.delete_rows(file_path, sheet_name, row, count)

    @mcp.tool()
    def excel_insert_columns(
        file_path: str,
        sheet_name: str,
        column: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """插入列
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column: 插入位置（列索引）
            count: 插入列数
            
        Returns:
            {success, inserted_at, count}
        """
        return ExcelOperations.insert_columns(file_path, sheet_name, column, count)

    @mcp.tool()
    def excel_delete_columns(
        file_path: str,
        sheet_name: str,
        column: int,
        count: int = 1
    ) -> Dict[str, Any]:
        """删除列
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column: 删除起始列
            count: 删除列数
            
        Returns:
            {success, deleted_from, count}
        """
        return ExcelOperations.delete_columns(file_path, sheet_name, column, count)
