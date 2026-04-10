#!/usr/bin/env python3
"""
专门测试 GROUP_CONCAT 支持复杂表达式
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
    print("测试 P1: GROUP_CONCAT 支持复杂表达式")
    print("="*70)

    # 创建测试数据
    import tempfile
    import uuid

    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, f"test_groupconcat_complex_{uuid.uuid4().hex[:8]}.xlsx")

    characters = pd.DataFrame({
        'Class': ['Mage', 'Mage', 'Warrior', 'Warrior', 'Priest', 'Priest'],
        'Level': [80, 65, 70, 50, 60, 55],
        'CharName': ['Alice', 'Bob', 'Charlie', 'David', 'Eve', 'Frank']
    })

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        characters.to_excel(writer, sheet_name='Characters', index=False)

    print(f"\n📁 测试文件: {file_path}")
    print(f"   Sheet 'Characters': {len(characters)} 行")
    print("   数据预览:")
    for idx, row in characters.iterrows():
        print(f"      {row['Class']}: {row['CharName']} (Level {row['Level']})")

    engine = AdvancedSQLQueryEngine()

    # 测试 1: GROUP_CONCAT with CASE WHEN
    print("\n" + "-"*70)
    print("测试 1: GROUP_CONCAT(CASE WHEN ... END)")
    print("-"*70)
    print("   SQL: SELECT Class, GROUP_CONCAT(")
    print("           CASE")
    print("             WHEN Level >= 70 THEN 'Veteran'")
    print("             WHEN Level >= 50 THEN 'Mid'")
    print("             ELSE 'Junior'")
    print("           END")
    print("         ) as Levels")
    print("         FROM Characters GROUP BY Class")
    print("-"*70)

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
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果:")
        print("      Mage: Veteran,Mid (Alice=80, Bob=65)")
        print("      Warrior: Veteran,Mid (Charlie=70, David=50)")
        print("      Priest: Mid,Mid (Eve=60, Frank=55)")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")
        print(f"   错误类型: {result.get('query_info', {}).get('error_type')}")

    # 测试 2: GROUP_CONCAT with mathematical expression
    print("\n" + "-"*70)
    print("测试 2: GROUP_CONCAT(Level * 2)")
    print("-"*70)
    print("   SQL: SELECT Class, GROUP_CONCAT(Level * 2) as DoubleLevels")
    print("         FROM Characters GROUP BY Class")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(Level * 2) as DoubleLevels " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果:")
        print("      Mage: 160,130 (80*2, 65*2)")
        print("      Warrior: 140,100 (70*2, 50*2)")
        print("      Priest: 120,110 (60*2, 55*2)")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")
        print(f"   错误类型: {result.get('query_info', {}).get('error_type')}")

    # 测试 3: GROUP_CONCAT with COALESCE
    print("\n" + "-"*70)
    print("测试 3: GROUP_CONCAT(COALESCE(Level, 0))")
    print("-"*70)
    print("   SQL: SELECT Class, GROUP_CONCAT(COALESCE(Level, 0)) as Levels")
    print("         FROM Characters GROUP BY Class")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(COALESCE(Level, 0)) as Levels " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
        print("\n   预期结果:")
        print("      Mage: 80,65")
        print("      Warrior: 70,50")
        print("      Priest: 60,55")
    else:
        print(f"❌ 测试失败")
        print(f"   错误信息: {result.get('message')}")
        print(f"   错误类型: {result.get('query_info', {}).get('error_type')}")

    # 测试 4: GROUP_CONCAT with simple column (baseline)
    print("\n" + "-"*70)
    print("测试 4: GROUP_CONCAT(CharName) - 简单列名（基线测试）")
    print("-"*70)
    print("   SQL: SELECT Class, GROUP_CONCAT(CharName) as Names")
    print("         FROM Characters GROUP BY Class")
    print("-"*70)

    result = engine.execute_sql_query(
        file_path,
        "SELECT Class, GROUP_CONCAT(CharName) as Names " +
        "FROM Characters GROUP BY Class"
    )

    if result['success']:
        print("✅ 测试通过")
        data = result['data']
        print(f"   返回行数: {len(data) - 1} (不含表头)")
        print(f"   列名: {data[0]}")
        print("   数据行:")
        for row in data[1:]:
            print(f"      {row}")
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
