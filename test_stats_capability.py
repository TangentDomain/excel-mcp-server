#!/usr/bin/env python3
"""
测试excel_evaluate_formula作为统计分析工具的能力
"""

import sys
from pathlib import Path
import time

# 添加src路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from excel_mcp.server import excel_evaluate_formula, excel_create_file, excel_update_range

def test_statistical_functions():
    """测试excel_evaluate_formula的统计函数支持"""

    print("📊 测试excel_evaluate_formula统计分析能力...")

    # 创建测试文件
    test_file = Path(__file__).parent / "stats_test.xlsx"

    print("📝 创建测试数据...")
    result = excel_create_file(str(test_file), ["Data"])
    if not result.get('success'):
        print(f"❌ 创建文件失败: {result}")
        return

    # 创建有意义的测试数据
    test_data = [
        ["销售额", "数量", "价格", "评分"],      # 标题行
        [1000, 50, 20.0, 4.5],               # 数据行1
        [1500, 75, 20.0, 4.2],               # 数据行2
        [800, 40, 20.0, 3.8],                # 数据行3
        [2000, 100, 20.0, 4.8],              # 数据行4
        [1200, 60, 20.0, 4.1],               # 数据行5
        [1800, 90, 20.0, 4.6],               # 数据行6
        [950, 45, 21.0, 3.9],                # 数据行7
        [1650, 80, 20.5, 4.4],               # 数据行8
        [1100, 55, 19.8, 4.0],               # 数据行9
        [1750, 85, 20.5, 4.7]                # 数据行10
    ]

    result = excel_update_range(str(test_file), "A1:D11", test_data)
    if not result.get('success'):
        print(f"❌ 写入数据失败: {result}")
        return

    print("\n🧪 测试各种统计函数...")

    # 定义统计测试用例
    stats_tests = [
        # 基础统计函数
        {
            "category": "基础统计",
            "tests": [
                {"name": "数据总数", "formula": "COUNT(A2:A11)", "range": "销售额"},
                {"name": "数值求和", "formula": "SUM(A2:A11)", "range": "销售额"},
                {"name": "平均值", "formula": "AVERAGE(A2:A11)", "range": "销售额"},
                {"name": "最大值", "formula": "MAX(A2:A11)", "range": "销售额"},
                {"name": "最小值", "formula": "MIN(A2:A11)", "range": "销售额"},
            ]
        },
        # 高级统计函数
        {
            "category": "高级统计",
            "tests": [
                {"name": "中位数", "formula": "MEDIAN(A2:A11)", "range": "销售额"},
                {"name": "标准差", "formula": "STDEV(A2:A11)", "range": "销售额"},
                {"name": "方差", "formula": "VAR(A2:A11)", "range": "销售额"},
                {"name": "百分位数", "formula": "PERCENTILE(A2:A11,0.9)", "range": "销售额90%"},
            ]
        },
        # 多列统计
        {
            "category": "多列分析",
            "tests": [
                {"name": "数量平均值", "formula": "AVERAGE(B2:B11)", "range": "数量"},
                {"name": "评分最高", "formula": "MAX(D2:D11)", "range": "评分"},
                {"name": "评分最低", "formula": "MIN(D2:D11)", "range": "评分"},
                {"name": "评分平均", "formula": "AVERAGE(D2:D11)", "range": "评分"},
            ]
        },
        # 条件统计
        {
            "category": "条件统计",
            "tests": [
                {"name": "计数大于1500", "formula": "COUNTIF(A2:A11,\">1500\")", "range": "销售额>1500"},
                {"name": "求和大于1500", "formula": "SUMIF(A2:A11,\">1500\")", "range": "销售额>1500"},
                {"name": "平均高评分", "formula": "AVERAGEIF(D2:D11,\">4.5\")", "range": "评分>4.5"},
            ]
        }
    ]

    total_start_time = time.time()
    success_count = 0
    total_tests = 0
    category_results = {}

    for category_data in stats_tests:
        category = category_data["category"]
        tests = category_data["tests"]

        print(f"\n📋 {category} ({len(tests)}个函数)")
        category_results[category] = {"success": 0, "total": len(tests), "time": 0}

        for test in tests:
            total_tests += 1
            print(f"   🧮 {test['name']}: {test['formula']}")

            start_time = time.time()
            result = excel_evaluate_formula(
                file_path=str(test_file),
                formula=test['formula']
            )
            exec_time = (time.time() - start_time) * 1000

            if result.get('success'):
                value = result.get('result')
                result_type = result.get('result_type', 'unknown')
                print(f"      ✅ 结果: {value} ({result_type}) - {exec_time:.1f}ms")
                success_count += 1
                category_results[category]["success"] += 1
                category_results[category]["time"] += exec_time
            else:
                error = result.get('error', 'Unknown error')
                print(f"      ❌ 失败: {error}")

    total_time = (time.time() - total_start_time) * 1000

    # 汇总结果
    print(f"\n📊 测试结果汇总:")
    print(f"   🎯 总体成功率: {success_count}/{total_tests} ({success_count/total_tests*100:.1f}%)")
    print(f"   ⏱️  总执行时间: {total_time:.1f}ms")
    print(f"   📈 平均每次调用: {total_time/total_tests:.1f}ms")

    print(f"\n📈 分类统计:")
    for category, stats in category_results.items():
        success_rate = stats["success"] / stats["total"] * 100
        avg_time = stats["time"] / stats["total"] if stats["total"] > 0 else 0
        print(f"   📋 {category}: {stats['success']}/{stats['total']} ({success_rate:.1f}%) - 平均{avg_time:.1f}ms")

    # 清理
    print(f"\n🧹 清理测试文件...")
    test_file.unlink(missing_ok=True)

    return {
        "success_rate": success_count / total_tests,
        "total_time": total_time,
        "avg_time_per_call": total_time / total_tests,
        "categories": category_results
    }

if __name__ == "__main__":
    results = test_statistical_functions()

    print(f"\n🎯 结论:")
    if results["success_rate"] >= 0.8:
        print(f"   ✅ excel_evaluate_formula具有强大的统计分析能力")
        print(f"   📊 支持大部分Excel统计函数")

        if results["avg_time_per_call"] < 100:
            print(f"   ⚡ 性能表现良好，适合频繁调用")
        else:
            print(f"   ⏱️  多次调用可能有性能开销，建议考虑批量处理")
    else:
        print(f"   ⚠️  统计函数支持有限，可能需要excel_get_statistics补充")

    print(f"   💡 建议: {'可以替代excel_get_statistics' if results['success_rate'] >= 0.8 and results['avg_time_per_call'] < 100 else '仍需要excel_get_statistics优化体验'}")
