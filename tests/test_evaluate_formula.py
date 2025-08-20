#!/usr/bin/env python3
"""
测试excel_evaluate_formula工具
"""

import sys
from pathlib import Path

# 添加src路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from excel_mcp.core.excel_writer import ExcelWriter
from excel_mcp.models.types import OperationResult

def test_evaluate_formula():
    """测试公式计算功能"""

    # 使用现有的测试数据文件
    test_file = Path(__file__).parent / "data" / "test_all_features.xlsx"

    if not test_file.exists():
        print(f"❌ 测试文件不存在: {test_file}")
        return

    print(f"🔄 使用测试文件: {test_file}")

    writer = ExcelWriter(str(test_file))

    # 测试用例
    test_cases = [
        {
            "name": "简单求和",
            "formula": "SUM(A1:A5)",
            "context_sheet": None
        },
        {
            "name": "平均值计算",
            "formula": "AVERAGE(B1:B10)",
            "context_sheet": None
        },
        {
            "name": "数学运算",
            "formula": "10 + 20 * 3",
            "context_sheet": None
        },
        {
            "name": "逻辑判断",
            "formula": "IF(10>5,\"大于\",\"小于\")",
            "context_sheet": None
        },
        {
            "name": "文本函数",
            "formula": "CONCATENATE(\"Hello\",\" \",\"World\")",
            "context_sheet": None
        }
    ]

    print("🧪 开始测试公式计算...")

    for i, case in enumerate(test_cases, 1):
        print(f"\n📋 测试 {i}: {case['name']}")
        print(f"   公式: {case['formula']}")

        try:
            result = writer.evaluate_formula(
                formula=case['formula'],
                context_sheet=case['context_sheet']
            )

            if result.success:
                metadata = result.metadata or {}
                print(f"   ✅ 成功")
                print(f"   📊 结果: {metadata.get('result')}")
                print(f"   📝 类型: {metadata.get('result_type')}")
                print(f"   ⏱️  耗时: {metadata.get('execution_time_ms', 0)}ms")
            else:
                print(f"   ❌ 失败: {result.error}")

        except Exception as e:
            print(f"   💥 异常: {e}")

    print(f"\n🎯 测试完成!")

if __name__ == "__main__":
    test_evaluate_formula()
