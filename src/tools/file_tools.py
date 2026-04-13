# 文件操作工具模块 + MCP Resources
"""文件操作工具和 MCP Resources 定义"""

from typing import Any

from ..api.excel_operations import ExcelOperations

# ==================== MCP Resources ====================


def register_resources(mcp) -> None:
    """注册 MCP Resources - 将 Excel 文件/工作表暴露为可读资源"""

    @mcp.resource("excel://sheets/{file_path}")
    def list_sheets_resource(file_path: str) -> dict[str, Any]:
        """列出 Excel 文件中的所有工作表

        URI参数:
            file_path: Excel文件路径

        返回:
            工作表列表
        """
        return ExcelOperations.list_sheets(file_path)

    @mcp.resource("excel://info/{file_path}")
    def file_info_resource(file_path: str) -> dict[str, Any]:
        """获取 Excel 文件的基本信息

        URI参数:
            file_path: Excel文件路径

        返回:
            文件信息
        """
        return ExcelOperations.get_file_info(file_path)

    @mcp.resource("excel://headers/{file_path}/{sheet_name}")
    def get_headers_resource(file_path: str, sheet_name: str) -> dict[str, Any]:
        """获取工作表的表头信息

        URI参数:
            file_path: Excel文件路径
            sheet_name: 工作表名称

        返回:
            表头信息
        """
        return ExcelOperations.get_headers(file_path, sheet_name)

    @mcp.resource("excel://data/{file_path}/{sheet_name}")
    def get_data_resource(file_path: str, sheet_name: str) -> dict[str, Any]:
        """读取工作表的全部数据

        URI参数:
            file_path: Excel文件路径
            sheet_name: 工作表名称

        返回:
            工作表数据
        """
        # 获取最后一行
        last_row_result = ExcelOperations.find_last_row(file_path, sheet_name)
        last_row = last_row_result.get("last_row", 100) if last_row_result.get("success") else 100

        # 读取数据范围 A1:Z{last_row}
        range_expr = f"{sheet_name}!A1:Z{last_row}"
        return ExcelOperations.get_range(file_path, range_expr)


# ==================== 文件操作工具 ====================


def register_file_tools(mcp) -> None:
    """注册文件操作相关工具"""

    @mcp.tool()
    def excel_list_sheets(file_path: str) -> dict[str, Any]:
        """列出所有工作表名称

        Args:
            file_path: Excel文件路径

        Returns:
            {success, sheets, total_sheets}
        """
        return ExcelOperations.list_sheets(file_path)

    @mcp.tool()
    def excel_get_file_info(file_path: str) -> dict[str, Any]:
        """获取 Excel 文件的基本信息

        Args:
            file_path: Excel文件路径

        Returns:
            {success, file_info: {path, name, sheets, created, modified}}
        """
        return ExcelOperations.get_file_info(file_path)

    @mcp.tool()
    def excel_create_file(file_path: str, sheet_names: list[str] | None = None) -> dict[str, Any]:
        """创建新的 Excel 文件

        Args:
            file_path: 新文件路径
            sheet_names: 工作表名称列表，默认 ['Sheet1']

        Returns:
            {success, file_path, sheets}
        """
        return ExcelOperations.create_file(file_path, sheet_names)

    @mcp.tool()
    def excel_create_sheet(file_path: str, sheet_name: str, index: int | None = None) -> dict[str, Any]:
        """创建新的工作表

        Args:
            file_path: Excel文件路径
            sheet_name: 新工作表名称
            index: 位置索引（可选）

        Returns:
            {success, sheet_name}
        """
        return ExcelOperations.create_sheet(file_path, sheet_name, index)

    @mcp.tool()
    def excel_delete_sheet(file_path: str, sheet_name: str) -> dict[str, Any]:
        """删除工作表

        Args:
            file_path: Excel文件路径
            sheet_name: 要删除的工作表名称

        Returns:
            {success, message}
        """
        return ExcelOperations.delete_sheet(file_path, sheet_name)

    @mcp.tool()
    def excel_rename_sheet(file_path: str, old_name: str, new_name: str) -> dict[str, Any]:
        """重命名工作表

        Args:
            file_path: Excel文件路径
            old_name: 原工作表名称
            new_name: 新工作表名称

        Returns:
            {success, old_name, new_name}
        """
        return ExcelOperations.rename_sheet(file_path, old_name, new_name)

    @mcp.tool()
    def excel_get_sheet_headers(file_path: str) -> dict[str, Any]:
        """获取 Excel 文件中所有工作表的双行表头信息

        Args:
            file_path: Excel文件路径 (.xlsx/.xlsm)

        Returns:
            {success, sheets_with_headers: [{name, headers, descriptions, field_names, header_count}], total_sheets}
        """
        return ExcelOperations.get_sheet_headers(file_path)
