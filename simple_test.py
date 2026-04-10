#!/usr/bin/env python3
"""Simple test runner without pytest"""
import sys
import os
import tempfile

# Add paths
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server')
os.chdir('/root/.openclaw/workspace/excel-mcp-server')

def test_group_concat():
    """Test GROUP_CONCAT functionality"""
    print("\n" + "="*70)
    print("TEST 1: GROUP_CONCAT Tests")
    print("="*70)

    try:
        from openpyxl import Workbook
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        # Create test data
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name

        wb = Workbook()
        ws = wb.active
        ws.title = '技能配置'
        ws.append(['部门', 'dept', '技能名称', 'skill_name'])
        ws.append(['法师', 'mage', '火球术', 'fireball'])
        ws.append(['法师', 'mage', '冰冻术', 'ice'])
        ws.append(['法师', 'mage', '火墙', 'firewall'])
        ws.append(['战士', 'warrior', '斩击', 'slash'])
        ws.append(['战士', 'warrior', '旋风斩', 'whirlwind'])
        ws.append(['牧师', 'priest', '治疗术', 'heal'])
        ws.append(['牧师', 'priest', '圣光术', 'holy'])
        wb.save(tmp_path)

        # Test 1: Basic GROUP_CONCAT
        print("\nTest 1.1: Basic GROUP_CONCAT")
        result = execute_advanced_sql_query(
            tmp_path,
            "SELECT 部门, GROUP_CONCAT(技能名称) as 技能列表 FROM 技能配置 GROUP BY 部门"
        )

        if result['success']:
            print("✓ Query executed successfully")
            data = result.get('data', [])
            if data and len(data) > 1:
                headers = data[0]
                print(f"  Headers: {headers}")
                for i, row in enumerate(data[1:], 1):
                    print(f"  Row {i}: {row}")
            else:
                print("✗ No data returned")
                return False
        else:
            print(f"✗ Query failed: {result.get('message')}")
            return False

        # Test 2: GROUP_CONCAT with separator
        print("\nTest 1.2: GROUP_CONCAT with custom separator")
        result = execute_advanced_sql_query(
            tmp_path,
            "SELECT 部门, GROUP_CONCAT(skill_name, '|') as skills FROM 技能配置 GROUP BY 部门"
        )

        if result['success']:
            print("✓ Query executed successfully")
            data = result.get('data', [])
            if data and len(data) > 1:
                for i, row in enumerate(data[1:], 1):
                    print(f"  Row {i}: {row}")
            else:
                print("✗ No data returned")
                return False
        else:
            print(f"✗ Query failed: {result.get('message')}")
            return False

        # Cleanup
        os.unlink(tmp_path)
        print("\n✓ GROUP_CONCAT tests PASSED")
        return True

    except Exception as e:
        print(f"\n✗ GROUP_CONCAT tests FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_right_join():
    """Test RIGHT JOIN functionality"""
    print("\n" + "="*70)
    print("TEST 2: RIGHT JOIN Test")
    print("="*70)

    try:
        import pandas as pd
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        engine = AdvancedSQLQueryEngine()

        # Create test data
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name

        skills = pd.DataFrame({
            'skill_id': [1, 2, 3, 4],
            'skill_name': ['火球术', '斩击', '治疗术', '冰冻术'],
            'type': ['法术', '物理', '法术', '法术'],
            'damage': [200, 150, 0, 180]
        })

        unlocks = pd.DataFrame({
            'skill_id': [1, 2, 5],
            'level_req': [5, 1, 10],
            'cost': [100, 50, 200]
        })

        with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
            skills.to_excel(writer, sheet_name='技能表', index=False)
            unlocks.to_excel(writer, sheet_name='解锁表', index=False)

        # Test basic RIGHT JOIN
        print("\nTest 2.1: Basic RIGHT JOIN")
        result = engine.execute_sql_query(
            tmp_path,
            "SELECT a.skill_id, a.skill_name, b.level_req, b.cost FROM 技能表 a RIGHT JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )

        if result['success']:
            print("✓ Query executed successfully")
            data = result.get('data', [])
            if data and len(data) > 1:
                headers = data[0]
                print(f"  Headers: {headers}")
                for i, row in enumerate(data[1:], 1):
                    print(f"  Row {i}: {row}")

                # Check if we have 3 rows (right table has 3 rows)
                if len(data) - 1 == 3:
                    print(f"✓ Correct number of rows: {len(data) - 1}")
                else:
                    print(f"✗ Expected 3 rows, got {len(data) - 1}")
                    return False
            else:
                print("✗ No data returned")
                return False
        else:
            print(f"✗ Query failed: {result.get('message')}")
            return False

        # Cleanup
        os.unlink(tmp_path)
        print("\n✓ RIGHT JOIN test PASSED")
        return True

    except Exception as e:
        print(f"\n✗ RIGHT JOIN test FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests"""
    print("\n" + "="*70)
    print("RUNNING TESTS")
    print("="*70)

    test1_passed = test_group_concat()
    test2_passed = test_right_join()

    print("\n" + "="*70)
    print("TEST SUMMARY")
    print("="*70)
    print(f"GROUP_CONCAT tests: {'PASSED ✓' if test1_passed else 'FAILED ✗'}")
    print(f"RIGHT JOIN test: {'PASSED ✓' if test2_passed else 'FAILED ✗'}")

    return 0 if (test1_passed and test2_passed) else 1


if __name__ == "__main__":
    sys.exit(main())
