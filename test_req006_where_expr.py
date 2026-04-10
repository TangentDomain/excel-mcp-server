#!/usr/bin/env python3
"""测试 REQ-EXCEL-006: WHERE 子句算术表达式"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_where_arithmetic_expression():
    """测试 WHERE 子句中的算术表达式"""
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    data = pd.DataFrame({
        '角色': ['战士', '法师', '射手', '刺客'],
        '力量': [80, 30, 50, 60],
        '敏捷': [40, 70, 90, 80],
        '智力': [20, 100, 50, 70]
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='角色属性', index=False)

    try:
        engine = AdvancedSQLQueryEngine()

        # 测试 WHERE 中的算术表达式
        sql = "SELECT 角色, 力量, 敏捷, 智力 FROM 角色属性 WHERE 力量 + 敏捷 + 智力 > 180"
        result = engine.execute_sql_query(file_path, sql)

        print("=== WHERE 算术表达式测试 ===")
        print(f"SQL: {sql}")
        print(f"Success: {result['success']}")
        if result['success']:
            print(f"Data: {result['data']}")
        else:
            print(f"Error: {result.get('message', 'Unknown error')}")

        return result['success']

    finally:
        Path(file_path).unlink(missing_ok=True)

if __name__ == '__main__':
    success = test_where_arithmetic_expression()
    print(f"\n测试结果: {'通过' if success else '失败'}")
