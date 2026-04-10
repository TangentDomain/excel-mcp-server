#!/usr/bin/env python3
"""
专门测试同文件多表 JOIN（指定 sheet_name 场景）
"""
import os
import sys
import pandas as pd
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


def main():
    print("\n" + "="*70)
    print("测试 P0: 同文件内多表 JOIN（指定 sheet_name 场景）")
    print("="*70)

    # 创建测试数据
    import tempfile
    import uuid

    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, f"test_same_file_join_{uuid.uuid4().hex[:8]}.xlsx")

    # 创建两个 sheet
    characters = pd.DataFrame({
        'CharID': [1, 2, 3, 4],
        'CharName': ['Alice', 'Bob', 'Charlie', 'David'],
        'Level': [80, 65, 70, 55],
        'Class': ['Mage', 'Warrior', 'Mage', 'Priest']
    })

    raids = pd.DataFrame({
        'RaidID': [101, 102, 103],
        'RaidName': ['Dragon', 'Demon', 'Giant'],
        'CharID': [1, 2, 5]  # 5 不在 Characters 中
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        characters.to_excel(writer, sheet_name='Characters', index=False)
        raids.to_excel(writer, sheet_name='Raids', index=False)

    print(f"\n📁 测试文件: {file_path}")
    print(f"   Sheet 'Characters': {len(characters)} 行")
    print(f"   Sheet 'Raids': {len(raids)} 行")

    engine = AdvancedSQLQueryEngine()

    # 测试 1: 从 Characters sheet JOIN Raids sheet（指定 sheet_name='Characters'）
    print("\n" + "-"*70)
    print("测试 1: 从 Characters sheet JOIN Raids sheet")
    print("   SQL: SELECT c.CharID, c.CharName, r.RaidName")
    print("        FROM Characters c")
    print("        INNER JOIN Raids r ON c.CharID = r.CharID")
    print("   sheet_name='Characters' (只加载 Characters sheet)")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT c.CharID, c.CharName, r.RaidName " +
        "FROM Characters c " +
        "INNER JOIN Raids r ON c.CharID = r.CharID",
        sheet_name='Characters'  # 只指定 Characters sheet
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果: 2行 (Alice-Dragon, Bob-Demon)")
        if len(data) - 1 == 2:
            print("   ✅ 行数正确")
        else:
            print(f"   ❌ 行数错误，期望2行，实际{len(data)-1}行")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")
        print(f"   错误类型: {result.get('query_info', {}).get('error_type')}")

    # 测试 2: 从 Raids sheet JOIN Characters sheet（指定 sheet_name='Raids'）
    print("\n" + "-"*70)
    print("测试 2: 从 Raids sheet JOIN Characters sheet")
    print("   SQL: SELECT r.RaidName, c.CharName, c.Level")
    print("        FROM Raids r")
    print("        INNER JOIN Characters c ON r.CharID = c.CharID")
    print("   sheet_name='Raids' (只加载 Raids sheet)")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT r.RaidName, c.CharName, c.Level " +
        "FROM Raids r " +
        "INNER JOIN Characters c ON r.CharID = c.CharID",
        sheet_name='Raids'  # 只指定 Raids sheet
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果: 2行 (Dragon-Alice, Demon-Bob)")
        if len(data) - 1 == 2:
            print("   ✅ 行数正确")
        else:
            print(f"   ❌ 行数错误，期望2行，实际{len(data)-1}行")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")
        print(f"   错误类型: {result.get('query_info', {}).get('error_type')}")

    # 测试 3: LEFT JOIN（应该保留 Characters 所有行）
    print("\n" + "-"*70)
    print("测试 3: LEFT JOIN (保留 Characters 所有行)")
    print("   SQL: SELECT c.CharName, r.RaidName")
    print("        FROM Characters c")
    print("        LEFT JOIN Raids r ON c.CharID = r.CharID")
    print("   sheet_name='Characters'")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT c.CharName, r.RaidName " +
        "FROM Characters c " +
        "LEFT JOIN Raids r ON c.CharID = r.CharID",
        sheet_name='Characters'
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果: 4行 (4个角色，David无Raid)")
        if len(data) - 1 == 4:
            print("   ✅ 行数正确")
        else:
            print(f"   ❌ 行数错误，期望4行，实际{len(data)-1}行")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")

    # 清理
    import shutil
    shutil.rmtree(tmp_dir)
    print("\n" + "="*70)
    print("测试完成")
    print("="*70)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"\n❌ 测试执行出错: {e}")
        import traceback
        traceback.print_exc()
