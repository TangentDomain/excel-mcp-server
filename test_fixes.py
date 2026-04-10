#!/usr/bin/env python3
"""
临时测试脚本：验证 JOIN 和 GROUP_CONCAT 的修复效果
"""
import os
import sys
import pandas as pd
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


def test_same_file_join():
    """测试同文件多表 JOIN"""
    print("\n" + "="*60)
    print("测试 P0: 同文件内多表 JOIN")
    print("="*60)

    # 创建测试数据
    import tempfile
    import uuid

    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, f"test_join_{uuid.uuid4().hex[:8]}.xlsx")

    # 创建两个 sheet
    characters = pd.DataFrame({
        'CharID': [1, 2, 3, 4],
        'CharName': ['Alice', 'Bob', 'Charlie', 'David'],
        'Level': [80, 65, 70, 55]
    })

    raids = pd.DataFrame({
        'RaidID': [101, 102, 103],
        'RaidName': ['Dragon', 'Demon', 'Giant'],
        'CharID': [1, 2, 5]  # 5 不在 Characters 中
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        characters.to_excel(writer, sheet_name='Characters', index=False)
        raids.to_excel(writer, sheet_name='Raids', index=False)

    # 测试 JOIN
    engine = AdvancedSQLQueryEngine()

    # 测试 1: 从 Characters sheet JOIN Raids sheet
    print("\n测试 1: 从 Characters sheet JOIN Raids sheet")
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
        print(f"   结果行数: {len(data) - 1}")  # 减去表头
        print(f"   列名: {data[0]}")
        for row in data[1:]:
            print(f"   {row}")
    else:
        print(f"❌ 测试失败: {result.get('message')}")

    # 测试 2: 从 Raids sheet JOIN Characters sheet
    print("\n测试 2: 从 Raids sheet JOIN Characters sheet")
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
        print(f"   结果行数: {len(data) - 1}")
        print(f"   列名: {data[0]}")
        for row in data[1:]:
            print(f"   {row}")
    else:
        print(f"❌ 测试失败: {result.get('message')}")

    # 清理
    import shutil
    shutil.rmtree(tmp_dir)


def test_group_concat_complex_expression():
    """测试 GROUP_CONCAT 支持复杂表达式"""
    print("\n" + "="*60)
    print("测试 P1: GROUP_CONCAT 支持复杂表达式")
    print("="*60)

    # 创建测试数据
    import tempfile
    import uuid

    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, f"test_groupconcat_{uuid.uuid4().hex[:8]}.xlsx")

    characters = pd.DataFrame({
        'Class': ['Mage', 'Mage', 'Warrior', 'Warrior', 'Priest'],
        'Level': [80, 65, 70, 50, 60],
        'CharName': ['Alice', 'Bob', 'Charlie', 'David', 'Eve']
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        characters.to_excel(writer, sheet_name='Characters', index=False)

    engine = AdvancedSQLQueryEngine()

    # 测试 1: GROUP_CONCAT with CASE WHEN
    print("\n测试 1: GROUP_CONCAT(CASE WHEN ... END)")
    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(" +
        "  CASE " +
        "    WHEN Level >= 70 THEN 'Veteran' " +
        "    WHEN Level >= 50 THEN 'Mid' " +
        "    ELSE 'Junior' " +
        "  END" +
        ") as Levels " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   结果行数: {len(data) - 1}")
        print(f"   列名: {data[0]}")
        for row in data[1:]:
            print(f"   {row}")
    else:
        print(f"❌ 测试失败: {result.get('message')}")

    # 测试 2: GROUP_CONCAT with mathematical expression
    print("\n测试 2: GROUP_CONCAT(Level * 2)")
    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(Level * 2) as DoubleLevels " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   结果行数: {len(data) - 1}")
        print(f"   列名: {data[0]}")
        for row in data[1:]:
            print(f"   {row}")
    else:
        print(f"❌ 测试失败: {result.get('message')}")

    # 测试 3: GROUP_CONCAT with COALESCE
    print("\n测试 3: GROUP_CONCAT(COALESCE(Level, 0))")
    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(COALESCE(Level, 0)) as Levels " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   结果行数: {len(data) - 1}")
        print(f"   列名: {data[0]}")
        for row in data[1:]:
            print(f"   {row}")
    else:
        print(f"❌ 测试失败: {result.get('message')}")

    # 清理
    import shutil
    shutil.rmtree(tmp_dir)


if __name__ == '__main__':
    print("\n" + "="*60)
    print("ExcelMCP 修复验证测试")
    print("="*60)

    try:
        test_same_file_join()
        test_group_concat_complex_expression()

        print("\n" + "="*60)
        print("✅ 所有测试完成")
        print("="*60)

    except Exception as e:
        print(f"\n❌ 测试执行出错: {e}")
        import traceback
        traceback.print_exc()
