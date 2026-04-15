#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ExcelMCP SQL 校准器 - CLI 入口
===============================
用法:
    python -m excel_mcp_server_fastmcp.calibrate import  <xlsx文件路径> [数据库名]
    python -m excel_mcp_server_fastmcp.calibrate query   <数据库名> "<SQL语句>"
    python -m excel_mcp_server_fastmcp.calibrate tables  <数据库名>
    python -m excel_mcp_server_fastmcp.calibrate schema  <数据库名> <表名>

示例:
    # 导入 Excel 到 SQLite
    python -m excel_mcp_server_fastmcp.calibrate import /data/宝箱掉落道具清单_v2.xlsx test_db
    python -m excel_mcp_server_fastmcp.calibrate import /data/ChestProp.xlsx test_db

    # 查询
    python -m excel_mcp_server_fastmcp.calibrate query test_db "SELECT * FROM 宝箱掉落道具清单_v2 LIMIT 5"

    # 列出所有表
    python -m excel_mcp_server_fastmcp.calibrate tables test_db

    # 查看表结构
    python -m excel_mcp_server_fastmcp.calibrate schema test_db 宝箱掉落道具清单_v2
"""

import sys
import argparse

from .calibrator.core import (
    cmd_import,
    cmd_query,
    cmd_tables,
    cmd_schema,
    DEFAULT_DB_DIR,
)


def main():
    parser = argparse.ArgumentParser(
        description='ExcelMCP SQL 校准器 - 将游戏配置Excel导入SQLite并查询',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "使用示例:\n"
            "  %(prog)s import  data.xlsx mydb         # 导入Excel到mydb数据库\n"
            "  %(prog)s query   mydb \"SELECT * FROM t1 LIMIT 5\"\n"
            "  %(prog)s tables  mydb                   # 列出所有表\n"
            "  %(prog)s schema  mydb table1            # 查看表结构\n"
            "\n"
            f"数据库目录: {DEFAULT_DB_DIR}\n"
        ),
    )

    subparsers = parser.add_subparsers(dest='command', help='子命令')

    p_import = subparsers.add_parser('import', help='导入Excel文件到数据库')
    p_import.add_argument('xlsx_path', help='Excel文件路径')
    p_import.add_argument('db_name', nargs='?', default='default',
                          help='数据库名称 (默认: default)')

    p_query = subparsers.add_parser('query', help='执行SQL查询')
    p_query.add_argument('db_name', help='数据库名称')
    p_query.add_argument('sql', help='SQL查询语句')

    p_tables = subparsers.add_parser('tables', help='列出所有表')
    p_tables.add_argument('db_name', nargs='?', default='default',
                          help='数据库名称 (默认: default)')

    p_schema = subparsers.add_parser('schema', help='查看表结构')
    p_schema.add_argument('db_name', help='数据库名称')
    p_schema.add_argument('table_name', help='表名')

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == 'import':
        result = cmd_import(args.xlsx_path, args.db_name)
        if result['success']:
            # 打印日志信息（模拟原始 CLI 行为）
            for log in result.get('logs', []):
                print(f"  {log}")
            print(f"\n{'='*50}")
            print(result['message'])
            print(f"数据库路径: {result['db_path']}")
            sys.exit(0)
        else:
            print(f"[错误] {result['message']}")
            sys.exit(1)

    elif args.command == 'query':
        result = cmd_query(args.db_name, args.sql)
        if result['success']:
            print(result['formatted'])
            print(f"\n共 {result['row_count']} 行 | 耗时 {result['elapsed_ms']:.2f} ms")
            sys.exit(0)
        else:
            print(result['message'])
            sys.exit(1)

    elif args.command == 'tables':
        result = cmd_tables(args.db_name)
        if result['success']:
            print(result['formatted'])
            print(f"\n共 {len(result['tables'])} 张表")
            sys.exit(0)
        else:
            print(result['message'])
            sys.exit(1)

    elif args.command == 'schema':
        result = cmd_schema(args.db_name, args.table_name)
        if result['success']:
            print(f"\n表: {result['table_name']}")
            print(result['formatted'])
            print(f"共 {len(result['columns'])} 列")
            sys.exit(0)
        else:
            print(result['message'])
            sys.exit(1)


if __name__ == '__main__':
    main()
