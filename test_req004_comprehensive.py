#!/usr/bin/env python3
"""全面测试 REQ-EXCEL-004: EXISTS 子查询"""

import pandas as pd
import tempfile
from pathlib import Path
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def test_exists_subquery_comprehensive():
    """全面测试 EXISTS 子查询"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        file_path = f.name

    # 主表 - 技能配置
    skills_data = pd.DataFrame({
        '技能ID': [1, 2, 3, 4, 5],
        '技能名称': ['火球术', '冰箭', '闪电', '毒雾', '治愈'],
        '伤害': [200, 150, 250, 180, 100]
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
        all_passed = True

        # 测试1: EXISTS - 应该只返回技能ID为1和3的行
        print("=== 测试1: EXISTS 子查询 ===")
        sql1 = """
        SELECT 技能ID, 技能名称
        FROM 技能配置
        WHERE EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID
        )
        """
        result1 = engine.execute_sql_query(file_path, sql1)
        if result1['success']:
            data1 = result1['data']
            if len(data1) > 1:
                data_rows1 = data1[1:]
                ids = [row[0] for row in data_rows1]
                print(f"结果: {ids}")
                if set(ids) == {1, 3}:
                    print("✓ 测试1通过: 返回技能ID 1和3")
                else:
                    print(f"✗ 测试1失败: 预期{{1, 3}}, 实际{set(ids)}")
                    all_passed = False
            else:
                print("✗ 测试1失败: 没有数据")
                all_passed = False
        else:
            print(f"✗ 测试1失败: {result1.get('message', 'Unknown error')}")
            all_passed = False

        # 测试2: EXISTS with additional condition - 应该只返回技能ID为3的行
        print("\n=== 测试2: EXISTS 子查询带额外条件 ===")
        sql2 = """
        SELECT 技能ID, 技能名称
        FROM 技能配置
        WHERE EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID AND s2.伤害 > 200
        )
        """
        result2 = engine.execute_sql_query(file_path, sql2)
        if result2['success']:
            data2 = result2['data']
            if len(data2) > 1:
                data_rows2 = data2[1:]
                ids = [row[0] for row in data_rows2]
                print(f"结果: {ids}")
                if set(ids) == {3}:
                    print("✓ 测试2通过: 返回技能ID 3")
                else:
                    print(f"✗ 测试2失败: 预期{{3}}, 实际{set(ids)}")
                    all_passed = False
            else:
                print("✗ 测试2失败: 没有数据")
                all_passed = False
        else:
            print(f"✗ 测试2失败: {result2.get('message', 'Unknown error')}")
            all_passed = False

        # 测试3: NOT EXISTS - 应该返回技能ID为2, 4, 5的行
        print("\n=== 测试3: NOT EXISTS 子查询 ===")
        sql3 = """
        SELECT 技能ID, 技能名称
        FROM 技能配置
        WHERE NOT EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID
        )
        """
        result3 = engine.execute_sql_query(file_path, sql3)
        if result3['success']:
            data3 = result3['data']
            if len(data3) > 1:
                data_rows3 = data3[1:]
                ids = [row[0] for row in data_rows3]
                print(f"结果: {ids}")
                if set(ids) == {2, 4, 5}:
                    print("✓ 测试3通过: 返回技能ID 2, 4, 5")
                else:
                    print(f"✗ 测试3失败: 预期{{2, 4, 5}}, 实际{set(ids)}")
                    all_passed = False
            else:
                print("✗ 测试3失败: 没有数据")
                all_passed = False
        else:
            print(f"✗ 测试3失败: {result3.get('message', 'Unknown error')}")
            all_passed = False

        # 测试4: EXISTS with no matches - 应该返回空
        print("\n=== 测试4: EXISTS 子查询无匹配 ===")
        sql4 = """
        SELECT 技能ID, 技能名称
        FROM 技能配置
        WHERE EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID AND s2.伤害 > 300
        )
        """
        result4 = engine.execute_sql_query(file_path, sql4)
        if result4['success']:
            data4 = result4['data']
            if len(data4) > 1:
                data_rows4 = data4[1:]
                print(f"✗ 测试4失败: 应该返回空，实际返回{len(data_rows4)}行")
                all_passed = False
            else:
                print("✓ 测试4通过: 返回空结果")
        else:
            print(f"✗ 测试4失败: {result4.get('message', 'Unknown error')}")
            all_passed = False

        # 测试5: EXISTS with AND condition
        print("\n=== 测试5: EXISTS 子查询与AND条件 ===")
        sql5 = """
        SELECT 技能ID, 技能名称
        FROM 技能配置
        WHERE EXISTS (
            SELECT 1 FROM s2
            WHERE s2.技能ID = 技能配置.技能ID
        ) AND 技能ID > 2
        """
        result5 = engine.execute_sql_query(file_path, sql5)
        if result5['success']:
            data5 = result5['data']
            if len(data5) > 1:
                data_rows5 = data5[1:]
                ids = [row[0] for row in data_rows5]
                print(f"结果: {ids}")
                if set(ids) == {3}:
                    print("✓ 测试5通过: 返回技能ID 3")
                else:
                    print(f"✗ 测试5失败: 预期{{3}}, 实际{set(ids)}")
                    all_passed = False
            else:
                print("✗ 测试5失败: 没有数据")
                all_passed = False
        else:
            print(f"✗ 测试5失败: {result5.get('message', 'Unknown error')}")
            all_passed = False

        return all_passed

    finally:
        Path(file_path).unlink(missing_ok=True)

if __name__ == '__main__':
    success = test_exists_subquery_comprehensive()
    print(f"\n{'='*50}")
    print(f"全面测试结果: {'全部通过 ✓' if success else '存在失败 ✗'}")
    print(f"{'='*50}")
