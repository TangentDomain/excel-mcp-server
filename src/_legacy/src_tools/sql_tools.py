# SQL 工具模块

from typing import Any

from ..api.advanced_sql_query import execute_advanced_sql_query


def register_sql_tools(mcp) -> None:
    """注册 SQL 查询相关工具"""

    @mcp.tool()
    def excel_query(file_path: str, query_expression: str, include_headers: bool = True) -> dict[str, Any]:
        """执行 SQL 查询

        Args:
            file_path: Excel文件路径
            query_expression: SQL 查询语句
            include_headers: 是否包含表头

        Returns:
            {success, data, columns, row_count}

        支持的 SQL 特性:
        - SELECT, DISTINCT, 别名 (AS)
        - WHERE, LIKE, IN, BETWEEN
        - ORDER BY, LIMIT, OFFSET
        - COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
        """
        try:
            return execute_advanced_sql_query(
                file_path=file_path,
                sql=query_expression,
                sheet_name=None,
                limit=None,
                include_headers=include_headers,
            )
        except ImportError:
            return {
                "success": False,
                "message": "SQLGlot未安装，无法使用高级SQL功能",
                "data": [],
                "query_info": {"error_type": "dependency_missing"},
            }

    @mcp.tool()
    def excel_evaluate_formula(file_path: str, formula: str, sheet_name: str | None = None) -> dict[str, Any]:
        """计算 Excel 公式

        Args:
            file_path: Excel文件路径
            formula: 公式 (如 "SUM(A1:A10)")
            sheet_name: 工作表名称

        Returns:
            {success, result, formula}
        """
        from ..api.excel_operations import ExcelOperations

        return ExcelOperations.evaluate_formula(file_path, formula, sheet_name)

    @mcp.tool()
    def excel_set_formula(file_path: str, sheet_name: str, cell: str, formula: str) -> dict[str, Any]:
        """设置单元格公式

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            cell: 单元格位置 (如 "C1")
            formula: 公式

        Returns:
            {success, cell, formula}
        """
        from ..api.excel_operations import ExcelOperations

        return ExcelOperations.set_formula(file_path, sheet_name, cell, formula)
