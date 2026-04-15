# -*- coding: utf-8 -*-
"""
SQL 校准器 MCP 工具注册
=======================
将 SQL 校准器功能注册为 MCP 工具，供 AI Agent 调用。

用法（在创建 FastMCP 实例后调用）:
    from excel_mcp_server_fastmcp.calibrator.tools import register_calibrator_tools
    register_calibrator_tools(mcp)
"""

from typing import Any

from .core import (
    cmd_import,
    cmd_query,
    cmd_tables,
    cmd_schema,
)


def register_calibrator_tools(mcp) -> None:
    """注册 SQL 校准器相关 MCP 工具"""

    @mcp.tool()
    def calibrate_import(
        xlsx_path: str,
        db_name: str = "default",
    ) -> dict[str, Any]:
        """导入 Excel 文件到 SQLite 校准数据库

        将游戏配置 Excel 文件导入到 SQLite 数据库，支持双表头智能拍平、
        多 Sheet 分别建表、自动推断列类型、添加 _rowid_ 自增主键。

        导入后可用 calibrate_query 执行纯 SQL 查询，与 excel_query 结果对比，
        用于开发阶段调试和 bug 定位。

        Args:
            xlsx_path: Excel 文件路径（绝对或相对路径）
            db_name: 数据库名称 (默认: default)，同一 db_name 可导入多个文件

        Returns:
            {success, message, db_path, tables, total_tables, total_rows}

        示例:
            calibrate_import("/data/宝箱掉落道具清单_v2.xlsx", "test_db")
            calibrate_import("/data/ChestProp.xlsx", "test_db")  # 同一数据库
        """
        result = cmd_import(xlsx_path, db_name)
        if result["success"]:
            return {
                "success": True,
                "message": result["message"],
                "db_path": result["db_path"],
                "tables": [
                    {
                        "sheet": t["sheet_name"],
                        "table": t["table_name"],
                        "rows": t["rows"],
                        "columns": t["columns"],
                    }
                    for t in result["tables"]
                ],
                "total_tables": result["total_tables"],
                "total_rows": result["total_rows"],
            }
        else:
            return {
                "success": False,
                "message": result["message"],
            }

    @mcp.tool()
    def calibrate_query(
        db_name: str,
        sql: str,
    ) -> dict[str, Any]:
        """在校准数据库中执行纯 SQL 查询

        对已导入的 SQLite 数据库执行标准 SQL 查询，返回格式化结果和执行耗时。
        用于与 excel_query 的结果做对比，定位 SQL 执行差异。

        Args:
            db_name: 数据库名称（需先通过 calibrate_import 导入）
            sql: SQL 查询语句

        Returns:
            {success, headers, rows, row_count, elapsed_ms, formatted}

        示例:
            calibrate_query("test_db", "SELECT * FROM 宝箱掉落道具清单_v2 LIMIT 5")
            calibrate_query("test_db", "SELECT ChestPropID, COUNT(*) FROM ChestProp GROUP BY ChestPropID")
        """
        result = cmd_query(db_name, sql)
        if result["success"]:
            return {
                "success": True,
                "headers": result["headers"],
                "rows": result["rows"],
                "row_count": result["row_count"],
                "elapsed_ms": result["elapsed_ms"],
                "formatted": result["formatted"],
            }
        else:
            return {
                "success": False,
                "message": result["message"],
            }

    @mcp.tool()
    def calibrate_tables(db_name: str = "default") -> dict[str, Any]:
        """列出校准数据库中的所有表及其行数

        Args:
            db_name: 数据库名称 (默认: default)

        Returns:
            {success, tables: [{name, count}], formatted}
        """
        result = cmd_tables(db_name)
        if result["success"]:
            return {
                "success": True,
                "tables": result["tables"],
                "formatted": result["formatted"],
            }
        else:
            return {
                "success": False,
                "message": result["message"],
            }

    @mcp.tool()
    def calibrate_schema(
        db_name: str,
        table_name: str,
    ) -> dict[str, Any]:
        """查看校准数据库中指定表的结构信息

        Args:
            db_name: 数据库名称
            table_name: 表名

        Returns:
            {success, table_name, columns: [{cid, name, type, notnull, pk, default}], formatted}
        """
        result = cmd_schema(db_name, table_name)
        if result["success"]:
            return {
                "success": True,
                "table_name": result["table_name"],
                "columns": result["columns"],
                "formatted": result["formatted"],
            }
        else:
            return {
                "success": False,
                "message": result["message"],
            }
