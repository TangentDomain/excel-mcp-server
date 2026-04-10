#!/usr/bin/env python3
"""
测试 REQ-EXCEL-008: 非等值连接支持

测试场景:
- ON s.等级限制 <= e.等级限制
- ON s.伤害 > e.最小伤害
- ON s.等级限制 >= e.等级限制
"""
import pandas as pd
import tempfile
import os
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

def create_test_excel():
    """创建测试Excel文件"""
    df1 = pd.DataFrame({
        '技能ID': [1, 2, 3, 4, 5],
        '技能名称': ['火球术', '冰霜箭', '雷电术', '治愈术', '复活术'],
        '等级限制': [1, 5, 10, 15, 20]
    })
    df2 = pd.DataFrame({
        '角色ID': [101, 102, 103],
        '角色名称': ['战士', '法师', '牧师'],
        '等级': [5, 12, 18]
    })

    with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.xlsx') as f:
        df1.to_excel(f, sheet_name='技能表', index=False)
        df2.to_excel(f, sheet_name='角色表', index=False)
        return f.name

def test_non_equi_join_le():
    """测试 <= 非等值连接"""
    print("测试: 非等值连接 <= (等级限制 <= 角色)")
    excel_file = create_test_excel()
    try:
        engine = AdvancedSQLQueryEngine()
        sql = """
        SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
        FROM 技能表 s
        JOIN 角色表 r ON s.等级限制 <= r.等级
        ORDER BY s.等级限制, r.等级
        """
        result = engine.execute_sql_query(excel_file, sql)

        if result['success']:
            print(f"✓ 查询成功，返回 {len(result['data'])} 行")
            if result['data']:
                print(f"  数据预览:")
                for row in result['data'][:5]:
                    print(f"    {row}")
            return True
        else:
            print(f"✗ 查询失败: {result['message']}")
            return False
    finally:
        os.unlink(excel_file)

def test_non_equi_join_ge():
    """测试 >= 非等值连接"""
    print("\n测试: 非等值连接 >= (等级限制 >= 角色)")
    excel_file = create_test_excel()
    try:
        engine = AdvancedSQLQueryEngine()
        sql = """
        SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
        FROM 技能表 s
        JOIN 角色表 r ON s.等级限制 >= r.等级
        ORDER BY s.等级限制, r.等级
        """
        result = engine.execute_sql_query(excel_file, sql)

        if result['success']:
            print(f"✓ 查询成功，返回 {len(result['data'])} 行")
            if result['data']:
                print(f"  数据预览:")
                for row in result['data'][:5]:
                    print(f"    {row}")
            return True
        else:
            print(f"✗ 查询失败: {result['message']}")
            return False
    finally:
        os.unlink(excel_file)

def test_non_equi_join_lt():
    """测试 < 非等值连接"""
    print("\n测试: 非等值连接 < (等级限制 < 角色)")
    excel_file = create_test_excel()
    try:
        engine = AdvancedSQLQueryEngine()
        sql = """
        SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
        FROM 技能表 s
        JOIN 角色表 r ON s.等级限制 < r.等级
        ORDER BY s.等级限制, r.等级
        """
        result = engine.execute_sql_query(excel_file, sql)

        if result['success']:
            print(f"✓ 查询成功，返回 {len(result['data'])} 行")
            if result['data']:
                print(f"  数据预览:")
                for row in result['data'][:5]:
                    print(f"    {row}")
            return True
        else:
            print(f"✗ 查询失败: {result['message']}")
            return False
    finally:
        os.unlink(excel_file)

def test_non_equi_join_gt():
    """测试 > 非等值连接"""
    print("\n测试: 非等值连接 > (等级限制 > 角色)")
    excel_file = create_test_excel()
    try:
        engine = AdvancedSQLQueryEngine()
        sql = """
        SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
        FROM 技能表 s
        JOIN 角色表 r ON s.等级限制 > r.等级
        ORDER BY s.等级限制, r.等级
        """
        result = engine.execute_sql_query(excel_file, sql)

        if result['success']:
            print(f"✓ 查询成功，返回 {len(result['data'])} 行")
            if result['data']:
                print(f"  数据预览:")
                for row in result['data'][:5]:
                    print(f"    {row}")
            return True
        else:
            print(f"✗ 查询失败: {result['message']}")
            return False
    finally:
        os.unlink(excel_file)

def test_non_equi_join_ne():
    """测试 != 非等值连接"""
    print("\n测试: 非等值连接 != (等级限制 != 角色)")
    excel_file = create_test_excel()
    try:
        engine = AdvancedSQLQueryEngine()
        sql = """
        SELECT s.技能名称, s.等级限制, r.角色名称, r.等级
        FROM 技能表 s
        JOIN 角色表 r ON s.等级限制 != r.等级
        ORDER BY s.等级限制, r.等级
        """
        result = engine.execute_sql_query(excel_file, sql)

        if result['success']:
            print(f"✓ 查询成功，返回 {len(result['data'])} 行")
            if result['data']:
                print(f"  数据预览:")
                for row in result['data'][:5]:
                    print(f"    {row}")
            return True
        else:
            print(f"✗ 查询失败: {result['message']}")
            return False
    finally:
        os.unlink(excel_file)

if __name__ == '__main__':
    print("=" * 60)
    print("REQ-EXCEL-008: 非等值连接测试")
    print("=" * 60)

    results = []
    results.append(('<= 测试', test_non_equi_join_le()))
    results.append(('>= 测试', test_non_equi_join_ge()))
    results.append(('< 测试', test_non_equi_join_lt()))
    results.append(('> 测试', test_non_equi_join_gt()))
    results.append(('!= 测试', test_non_equi_join_ne()))

    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    for name, passed in results:
        status = "✓ 通过" if passed else "✗ 失败"
        print(f"{name}: {status}")

    all_passed = all(r[1] for r in results)
    print("\n" + ("=" * 60))
    if all_passed:
        print("✓ 所有测试通过！")
    else:
        print("✗ 部分测试失败")
    print("=" * 60)
