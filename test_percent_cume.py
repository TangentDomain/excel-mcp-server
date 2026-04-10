"""测试PERCENT_RANK和CUME_DIST窗口函数"""
import tempfile
import os
from openpyxl import Workbook
import sys

# 添加src到路径
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')

from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query


def create_test_excel():
    """创建测试数据"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xlsx', delete=False) as f:
        wb = Workbook()
        ws = wb.active
        ws.title = '技能配置'
        ws.append(['技能名称', 'skill_name', '技能类型', 'skill_type', '伤害', 'damage', '等级', 'level'])
        ws.append(['火球术', 'fireball', '法师', 'mage', 250, 250, 10, 10])
        ws.append(['雷击', 'lightning', '法师', 'mage', 250, 250, 10, 10])  # 同伤害
        ws.append(['火墙', 'firewall', '法师', 'mage', 220, 220, 10, 10])
        ws.append(['冰矛', 'ice_spear', '法师', 'mage', 200, 200, 8, 8])
        ws.append(['斩击', 'slash', '战士', 'warrior', 200, 200, 5, 5])
        ws.append(['旋风斩', 'whirlwind', '战士', 'warrior', 190, 190, 5, 5])
        ws.append(['冰冻术', 'ice', '法师', 'mage', 180, 180, 8, 8])
        ws.append(['治疗术', 'heal', '牧师', 'priest', 0, 0, 3, 3])
        wb.save(f.name)
        return f.name


def get_rows(result):
    """提取数据行"""
    if not result.get('success'):
        print(f"Query failed: {result.get('message')}")
        return []
    data = result.get('data', [])
    if not data or not isinstance(data[0], list):
        return []
    headers = data[0]
    return [dict(zip(headers, row)) for row in data[1:]]


def test_percent_rank():
    """测试PERCENT_RANK"""
    print("\n=== 测试 PERCENT_RANK ===")
    excel_path = create_test_excel()

    try:
        # 测试1: 基本PERCENT_RANK
        sql = "SELECT skill_name, damage, PERCENT_RANK() OVER (ORDER BY damage DESC) as pr FROM 技能配置"
        result = execute_advanced_sql_query(excel_path, sql)
        rows = get_rows(result)

        if rows:
            print(f"PERCENT_RANK基本查询: {len(rows)} 行")
            for row in rows:
                print(f"  {row['skill_name']:12} 伤害:{row['damage']:4} pr:{row['pr']:.4f}")

            # 验证第一行pr=0，最后一行pr=1
            assert abs(rows[0]['pr'] - 0.0) < 0.0001, f"第一行pr应该是0, 实际是{rows[0]['pr']}"
            assert abs(rows[-1]['pr'] - 1.0) < 0.0001, f"最后一行pr应该是1, 实际是{rows[-1]['pr']}"
            print("✓ PERCENT_RANK基本测试通过")
        else:
            print("✗ PERCENT_RANK基本测试失败")
            return False

        # 测试2: PERCENT_RANK with PARTITION BY
        sql2 = "SELECT skill_name, skill_type, damage, PERCENT_RANK() OVER (PARTITION BY skill_type ORDER BY damage DESC) as pr FROM 技能配置"
        result2 = execute_advanced_sql_query(excel_path, sql2)
        rows2 = get_rows(result2)

        if rows2:
            print(f"\nPERCENT_RANK with PARTITION BY: {len(rows2)} 行")
            mages = [r for r in rows2 if r['skill_type'] == 'mage']
            print(f"法师分组 (共{len(mages)}人):")
            for row in mages:
                print(f"  {row['skill_name']:12} pr:{row['pr']:.4f}")
            print("✓ PERCENT_RANK分区测试通过")
        else:
            print("✗ PERCENT_RANK分区测试失败")
            return False

        return True

    finally:
        os.unlink(excel_path)


def test_cume_dist():
    """测试CUME_DIST"""
    print("\n=== 测试 CUME_DIST ===")
    excel_path = create_test_excel()

    try:
        # 测试1: 基本CUME_DIST
        sql = "SELECT skill_name, damage, CUME_DIST() OVER (ORDER BY damage DESC) as cd FROM 技能配置"
        result = execute_advanced_sql_query(excel_path, sql)
        rows = get_rows(result)

        if rows:
            print(f"CUME_DIST基本查询: {len(rows)} 行")
            for row in rows:
                print(f"  {row['skill_name']:12} 伤害:{row['damage']:4} cd:{row['cd']:.4f}")

            # 验证最后一行cd=1
            assert abs(rows[-1]['cd'] - 1.0) < 0.0001, f"最后一行cd应该是1, 实际是{rows[-1]['cd']}"
            print("✓ CUME_DIST基本测试通过")
        else:
            print("✗ CUME_DIST基本测试失败")
            return False

        # 测试2: CUME_DIST with PARTITION BY
        sql2 = "SELECT skill_name, skill_type, damage, CUME_DIST() OVER (PARTITION BY skill_type ORDER BY damage DESC) as cd FROM 技能配置"
        result2 = execute_advanced_sql_query(excel_path, sql2)
        rows2 = get_rows(result2)

        if rows2:
            print(f"\nCUME_DIST with PARTITION BY: {len(rows2)} 行")
            mages = [r for r in rows2 if r['skill_type'] == 'mage']
            print(f"法师分组 (共{len(mages)}人):")
            for row in mages:
                print(f"  {row['skill_name']:12} cd:{row['cd']:.4f}")
            print("✓ CUME_DIST分区测试通过")
        else:
            print("✗ CUME_DIST分区测试失败")
            return False

        return True

    finally:
        os.unlink(excel_path)


def test_combined():
    """测试PERCENT_RANK和CUME_DIST同时使用"""
    print("\n=== 测试 PERCENT_RANK + CUME_DIST 组合 ===")
    excel_path = create_test_excel()

    try:
        sql = """
            SELECT skill_name, damage,
                   PERCENT_RANK() OVER (ORDER BY damage DESC) as pr,
                   CUME_DIST() OVER (ORDER BY damage DESC) as cd
            FROM 技能配置
        """
        result = execute_advanced_sql_query(excel_path, sql)
        rows = get_rows(result)

        if rows:
            print(f"组合查询: {len(rows)} 行")
            print(f"{'技能名称':12} {'伤害':4} {'PERCENT_RANK':14} {'CUME_DIST':10}")
            print("-" * 50)
            for row in rows:
                print(f"{row['skill_name']:12} {row['damage']:4} {row['pr']:14.4f} {row['cd']:10.4f}")
            print("✓ 组合测试通过")
            return True
        else:
            print("✗ 组合测试失败")
            return False

    finally:
        os.unlink(excel_path)


if __name__ == '__main__':
    print("开始测试PERCENT_RANK和CUME_DIST窗口函数...")

    success = True
    success = test_percent_rank() and success
    success = test_cume_dist() and success
    success = test_combined() and success

    if success:
        print("\n✅ 所有测试通过！")
    else:
        print("\n❌ 部分测试失败")
        sys.exit(1)
