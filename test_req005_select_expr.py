#!/usr/bin/env python3
"""测试 REQ-EXCEL-005: SELECT 子句计算表达式"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_select_arithmetic_expression():
    """测试 SELECT 子句中的算术表达式"""
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    data = pd.DataFrame({
        '技能名称': ['火球术', '冰箭', '闪电'],
        '伤害': [100, 80, 120],
        '倍率': [1.2, 1.5, 1.1]
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='技能配置', index=False)

    try:
        engine = AdvancedSQLQueryEngine()

        # 测试 SELECT 中的算术表达式
        sql = "SELECT 技能名称, (伤害 * 1.2) as 预期伤害 FROM 技能配置"
        result = engine.execute_sql_query(file_path, sql)

        print("=== SELECT 算术表达式测试 ===")
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
    success = test_select_arithmetic_expression()
    print(f"\n测试结果: {'通过' if success else '失败'}")
