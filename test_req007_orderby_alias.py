#!/usr/bin/env python3
"""测试 REQ-EXCEL-007: ORDER BY 使用 SELECT 别名"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_orderby_with_alias():
    """测试 ORDER BY 使用 SELECT 中定义的别名"""
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    data = pd.DataFrame({
        '技能': ['火球术', '冰箭', '闪电', '治疗术'],
        '伤害': [100, 80, 120, 0]
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='技能配置', index=False)

    try:
        engine = AdvancedSQLQueryEngine()

        # 测试 ORDER BY 使用别名
        sql = "SELECT 技能, 伤害 * 1.2 as 预期伤害 FROM 技能配置 ORDER BY 预期伤害 DESC"
        result = engine.execute_sql_query(file_path, sql)

        print("=== ORDER BY 使用别名测试 ===")
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
    success = test_orderby_with_alias()
    print(f"\n测试结果: {'通过' if success else '失败'}")
