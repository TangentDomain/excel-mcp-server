#!/usr/bin/env python3
"""
测试MCP服务器中的excel_evaluate_formula工具
"""

import sys
from pathlib import Path

# 添加src路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from excel_mcp.server import excel_evaluate_formula, excel_create_file, excel_update_range

def test_mcp_evaluate_formula():
    """测试MCP服务器中的公式评估工具"""

    print("🔧 测试MCP excel_evaluate_formula工具...")

    # 创建测试文件
    test_file = Path(__file__).parent / "mcp_test_formula.xlsx"

    print("📝 创建测试文件...")
    result = excel_create_file(str(test_file), ["TestSheet"])
    if not result.get('success'):
        print(f"❌ 创建文件失败: {result}")
        return

    # 添加测试数据
    print("📊 添加测试数据...")
    test_data = [
        [10, 1.5],   # A1, B1
        [20, 2.5],   # A2, B2
        [30, 3.5],   # A3, B3
        [40, 4.0],   # A4, B4
        [50, 5.5]    # A5, B5
    ]

    result = excel_update_range(str(test_file), "A1:B5", test_data)
    if not result.get('success'):
        print(f"❌ 写入数据失败: {result}")
        return

    # 测试各种公式
    print("\n🧪 开始测试MCP公式计算...")

    test_cases = [
        {
            "name": "A列求和",
            "formula": "SUM(A1:A5)",
            "expected": 150
        },
        {
            "name": "B列平均值",
            "formula": "AVERAGE(B1:B5)",
            "expected": 3.4  # (1.5+2.5+3.5+4.0+5.5)/5
        },
        {
            "name": "计算总数",
            "formula": "COUNT(A1:B5)",
            "expected": 10
        },
        {
            "name": "复杂表达式",
            "formula": "100 + 50 * 2",
            "expected": 200
        },
        {
            "name": "条件判断",
            "formula": 'IF(150>100,"大于","小于")',
            "expected": "大于"
        }
    ]

    success_count = 0

    for i, case in enumerate(test_cases, 1):
        print(f"\n📋 测试 {i}: {case['name']}")
        print(f"   公式: {case['formula']}")

        result = excel_evaluate_formula(
            file_path=str(test_file),
            formula=case['formula']
        )

        if result.get('success'):
            actual_result = result.get('result')
            result_type = result.get('result_type')
            execution_time = result.get('execution_time_ms', 0)

            print(f"   ✅ 成功")
            print(f"   📊 结果: {actual_result}")
            print(f"   📝 类型: {result_type}")
            print(f"   ⏱️  耗时: {execution_time}ms")

            # 验证结果
            if abs(float(actual_result) - case['expected']) < 0.01 if isinstance(case['expected'], (int, float)) else str(actual_result) == str(case['expected']):
                print(f"   🎯 验证: 通过")
                success_count += 1
            else:
                print(f"   ⚠️  验证: 失败 (期望: {case['expected']}, 实际: {actual_result})")
        else:
            print(f"   ❌ 失败: {result.get('error')}")

    print(f"\n🎯 MCP测试完成! 成功: {success_count}/{len(test_cases)}")

    # 清理
    print(f"\n🧹 清理测试文件...")
    test_file.unlink(missing_ok=True)
    print(f"✅ 清理完成")

    return success_count == len(test_cases)

if __name__ == "__main__":
    test_mcp_evaluate_formula()
