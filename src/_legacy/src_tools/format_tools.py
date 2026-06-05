# 格式化工具模块

from typing import Any

from ..api.excel_operations import ExcelOperations


def register_format_tools(mcp) -> None:
    """注册格式化和样式相关工具"""

    @mcp.tool()
    def excel_format_cells(
        file_path: str,
        sheet_name: str,
        range: str,
        bold: bool | None = None,
        font_size: int | None = None,
        font_color: str | None = None,
        bg_color: str | None = None,
        alignment: str | None = None,
        number_format: str | None = None,
        preset: str | None = None,
    ) -> dict[str, Any]:
        """格式化单元格

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            range: 单元格范围
            bold: 是否粗体
            font_size: 字体大小
            font_color: 字体颜色
            bg_color: 背景颜色
            alignment: 对齐方式 (left/center/right)
            number_format: 数字格式
            preset: 预设样式 (highlight/warning/header)

        Returns:
            {success, formatted_range}
        """
        return ExcelOperations.format_cells(
            file_path,
            sheet_name,
            range,
            bold,
            font_size,
            font_color,
            bg_color,
            alignment,
            number_format,
            preset,
        )

    @mcp.tool()
    def excel_merge_cells(file_path: str, sheet_name: str, range: str) -> dict[str, Any]:
        """合并单元格

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            range: 要合并的范围

        Returns:
            {success, merged_range}
        """
        return ExcelOperations.merge_cells(file_path, sheet_name, range)

    @mcp.tool()
    def excel_unmerge_cells(file_path: str, sheet_name: str, range: str) -> dict[str, Any]:
        """取消合并单元格

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            range: 要取消合并的范围

        Returns:
            {success, unmerged_range}
        """
        return ExcelOperations.unmerge_cells(file_path, sheet_name, range)

    @mcp.tool()
    def excel_set_borders(
        file_path: str,
        sheet_name: str,
        range: str,
        border_style: str = "thin",
        border_color: str | None = None,
    ) -> dict[str, Any]:
        """设置边框

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            range: 单元格范围
            border_style: 边框样式 (thin/medium/thick/double)
            border_color: 边框颜色

        Returns:
            {success, bordered_range}
        """
        return ExcelOperations.set_borders(file_path, sheet_name, range, border_style, border_color)

    @mcp.tool()
    def excel_set_row_height(file_path: str, sheet_name: str, row: int, height: float) -> dict[str, Any]:
        """设置行高

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            row: 行号
            height: 行高

        Returns:
            {success, row, height}
        """
        return ExcelOperations.set_row_height(file_path, sheet_name, row, height)

    @mcp.tool()
    def excel_set_column_width(file_path: str, sheet_name: str, column: int, width: float) -> dict[str, Any]:
        """设置列宽

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column: 列号
            width: 列宽

        Returns:
            {success, column, width}
        """
        return ExcelOperations.set_column_width(file_path, sheet_name, column, width)

    @mcp.tool()
    def excel_export_to_csv(
        file_path: str,
        sheet_name: str | None = None,
        output_path: str | None = None,
        encoding: str = "utf-8",
    ) -> dict[str, Any]:
        """导出为 CSV

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            output_path: 输出路径
            encoding: 编码格式

        Returns:
            {success, output_path, rows_exported}
        """
        return ExcelOperations.export_to_csv(file_path, sheet_name, output_path, encoding)

    @mcp.tool()
    def excel_import_from_csv(
        csv_path: str,
        file_path: str,
        sheet_name: str | None = None,
        encoding: str = "utf-8",
    ) -> dict[str, Any]:
        """从 CSV 导入

        Args:
            csv_path: CSV文件路径
            file_path: 目标Excel文件路径
            sheet_name: 目标工作表名称
            encoding: 编码格式

        Returns:
            {success, rows_imported, file_path}
        """
        return ExcelOperations.import_from_csv(csv_path, file_path, sheet_name, encoding)
