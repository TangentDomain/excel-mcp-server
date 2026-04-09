#!/usr/bin/env python3
"""测试 REQ-EXCEL-004: EXISTS 子查询"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_exists_subquery():
    """测试 EXISTS 子查询"""
    # 创建测试数据
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    # 主表 - 技能配置
    skills_data = pd.DataFrame({
        '技能ID': [1, 2, 3, 4, 5],
        '技能名称': ['火球术', '冰箭', '闪电', '毒雾', '治愈']
    })

    # 子表 - 高伤害技能
    high_damage_data = pd.DataFrame({
        '技能ID': [1, 3, 7],
        '伤害': [200, 250, 180]
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        skills_data.to_excel(writer, sheet_name='技能配置', index=False)
        high_damage_data.to_excel(writer, sheet_name='s2', index=False)

    try:
        engine = AdvancedSQLQueryEngine()

        # 测试 EXISTS 子查询 - 应该只返回技能ID为1和3的行
        sql = """
        SELECT 技能配置.技能ID, 技能配置.技能名称
        FROM 技能配置
        WHERE EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID AND s2.伤害 > 200
        )
        """
        result = engine.execute_sql_query(file_path, sql)

        print("=== EXISTS 子查询测试 ===")
        print(f"SQL: {sql.strip()}")
        print(f"Success: {result['success']}")
        if result['success']:
            print(f"Data: {result['data']}")
            # 验证结果应该只包含技能ID为3的行（因为只有技能ID=3的伤害>200）
            data = result['data']
            if len(data) > 1:  # 有表头
                data_rows = data[1:]
                print(f"数据行数: {len(data_rows)}")
                print(f"预期: 只有技能ID=3满足条件（伤害250>200）")
        else:
            print(f"Error: {result.get('message', 'Unknown error')}")

        return result['success']

    finally:
        Path(file_path).unlink(missing_ok=True)

if __name__ == '__main__':
    success = test_exists_subquery()
    print(f"\n测试结果: {'通过' if success else '失败'}")
