# 对比工具模块

from typing import Dict, Any, Optional, Union
from ..api.excel_operations import ExcelOperations


def register_compare_tools(mcp) -> None:
    """注册对比相关工具"""
    
    @mcp.tool()
    def excel_check_duplicate_ids(
        file_path: str,
        sheet_name: str,
        id_column: Union[int, str] = 1,
        header_row: int = 1
    ) -> Dict[str, Any]:
        """检查 ID 重复
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            id_column: ID列位置（数字或字母）
            header_row: 表头行号
            
        Returns:
            {success, has_duplicates, duplicate_ids, count, message}
        """
        return ExcelOperations.check_duplicate_ids(file_path, sheet_name, id_column, header_row)

    @mcp.tool()
    def excel_compare_sheets(
        file1_path: str,
        sheet1_name: str,
        file2_path: str,
        sheet2_name: str,
        id_column: Union[int, str] = 1,
        header_row: int = 1
    ) -> Dict[str, Any]:
        """比较两个工作表的差异
        
        Args:
            file1_path: 第一个文件
            sheet1_name: 第一个工作表
            file2_path: 第二个文件
            sheet2_name: 第二个工作表
            id_column: ID列位置
            header_row: 表头行号
            
        Returns:
            {success, data: {added, removed, modified}, message}
        """
        return ExcelOperations.compare_sheets(
            file1_path, sheet1_name, file2_path, sheet2_name, id_column, header_row
        )

    @mcp.tool()
    def excel_convert_format(
        file_path: str,
        output_format: str,
        output_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """转换文件格式
        
        Args:
            file_path: 源文件路径
            output_format: 目标格式 (csv/json/xlsx/xlsm)
            output_path: 输出路径
            
        Returns:
            {success, output_path, format}
        """
        return ExcelOperations.convert_format(file_path, output_format, output_path)

    @mcp.tool()
    def excel_merge_files(
        file_paths: list,
        output_path: str,
        mode: str = "sheets"
    ) -> Dict[str, Any]:
        """合并多个 Excel 文件
        
        Args:
            file_paths: 文件路径列表
            output_path: 输出文件路径
            mode: 合并模式 (sheets/append)
            
        Returns:
            {success, output_path, files_merged}
        """
        return ExcelOperations.merge_files(file_paths, output_path, mode)
