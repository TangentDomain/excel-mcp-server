#!/usr/bin/env python3
"""测试 REQ-EXCEL-003: IN/NOT IN 操作符"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_in_operator():
    """测试 IN 操作符"""
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    data = pd.DataFrame({
        '技能ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        '技能名称': ['火球术', '冰箭', '闪电', '毒雾', '治愈', '复活', '护盾', '狂暴', '隐身', '传送'],
        '伤害': [100, 80, 120, 50, 0, 0, 0, 150, 0, 0]
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='技能配置', index=False)

    try:
        engine = AdvancedSQLQueryEngine()

        # 测试 IN 操作符
        sql = "SELECT 技能名称, 伤害 FROM 技能配置 WHERE 技能ID IN (1, 3, 5, 7, 9)"
        result = engine.execute_sql_query(file_path, sql)

        print("=== IN 操作符测试 ===")
        print(f"SQL: {sql}")
        print(f"Success: {result['success']}")
        if result['success']:
            print(f"Data: {result['data']}")
        else:
            print(f"Error: {result.get('message', 'Unknown error')}")

        # 测试 NOT IN 操作符
        sql2 = "SELECT 技能名称, 伤害 FROM 技能配置 WHERE 技能ID NOT IN (2, 4, 6, 8, 10)"
        result2 = engine.execute_sql_query(file_path, sql2)

        print("\n=== NOT IN 操作符测试 ===")
        print(f"SQL: {sql2}")
        print(f"Success: {result2['success']}")
        if result2['success']:
            print(f"Data: {result2['data']}")
        else:
            print(f"Error: {result2.get('message', 'Unknown error')}")

        return result['success'] and result2['success']

    finally:
        Path(file_path).unlink(missing_ok=True)

if __name__ == '__main__':
    success = test_in_operator()
    print(f"\n测试结果: {'通过' if success else '失败'}")
