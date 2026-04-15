# -*- coding: utf-8 -*-
"""
ExcelMCP SQL 校准器 - 核心模块
================================
将游戏配置 Excel 文件导入 SQLite 数据库，支持双表头智能拍平，
提供 import / query / tables / schema 四个核心功能。

用途：在开发阶段，将同一个 xlsx 导入 SQLite 后跑同一条 SQL，
跟 excel_query 的返回结果做对比，定位 bug。
"""

from .core import (
    # 常量
    DEFAULT_DB_DIR,
    # 工具函数
    get_db_path,
    sanitize_table_name,
    sanitize_col_name,
    flatten_multiindex,
    is_likely_dual_header,
    infer_sqlite_type,
    format_table,
    # 核心命令（返回结构化数据，不直接 print）
    cmd_import,
    cmd_query,
    cmd_tables,
    cmd_schema,
)

from .tools import register_calibrator_tools

__all__ = [
    "DEFAULT_DB_DIR",
    "get_db_path",
    "sanitize_table_name",
    "sanitize_col_name",
    "flatten_multiindex",
    "is_likely_dual_header",
    "infer_sqlite_type",
    "format_table",
    "cmd_import",
    "cmd_query",
    "cmd_tables",
    "cmd_schema",
    "register_calibrator_tools",
]
