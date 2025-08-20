#!/usr/bin/env python3
"""
测试formulas库对高级统计函数的支持
"""

import sys
from pathlib import Path
import tempfile
import os

# 添加src路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_formulas_library():
    """测试formulas库的统计函数支持"""

    print("🧪 测试formulas库的高级统计支持...")

    try:
        # 测试基本导入
        import formulas
        print("✅ formulas库导入成功")

        # 创建简单的Excel模型进行测试
        from openpyxl import Workbook

        # 创建测试工作簿
        wb = Workbook()
        ws = wb.active
        ws.title = "TestData"

        # 添加测试数据
        test_values = [10, 20, 30, 40, 50, 15, 25, 35, 45, 55]
        for i, value in enumerate(test_values, 1):
            ws[f'A{i}'] = value

        # 保存到临时文件
        temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        temp_file.close()
        wb.save(temp_file.name)
        wb.close()

        print(f"📝 创建测试数据: {test_values}")

        # 测试各种统计函数
        test_functions = [
            ("基础统计", [
                ("SUM", "SUM(TestData!A1:A10)"),
                ("AVERAGE", "AVERAGE(TestData!A1:A10)"),
                ("COUNT", "COUNT(TestData!A1:A10)"),
                ("MIN", "MIN(TestData!A1:A10)"),
                ("MAX", "MAX(TestData!A1:A10)"),
            ]),
            ("高级统计", [
                ("MEDIAN", "MEDIAN(TestData!A1:A10)"),
                ("STDEV", "STDEV(TestData!A1:A10)"),
                ("VAR", "VAR(TestData!A1:A10)"),
                ("PERCENTILE", "PERCENTILE(TestData!A1:A10,0.9)"),
                ("QUARTILE", "QUARTILE(TestData!A1:A10,1)"),
            ]),
            ("条件统计", [
                ("COUNTIF", "COUNTIF(TestData!A1:A10,\">30\")"),
                ("SUMIF", "SUMIF(TestData!A1:A10,\">30\")"),
                ("AVERAGEIF", "AVERAGEIF(TestData!A1:A10,\">30\")"),
            ])
        ]

        # 使用formulas计算
        try:
            # 创建Excel模型
            xl_model = formulas.ExcelModel().loads(temp_file.name).finish()

            total_success = 0
            total_tests = 0

            for category, functions in test_functions:
                print(f"\n📊 {category}:")
                category_success = 0

                for func_name, formula in functions:
                    total_tests += 1
                    try:
                        # 使用formulas计算
                        result = xl_model.calculate(formula)
                        print(f"   ✅ {func_name}: {formula} = {result}")
                        total_success += 1
                        category_success += 1

                    except Exception as e:
                        print(f"   ❌ {func_name}: {formula} - 错误: {e}")

                print(f"   📈 成功率: {category_success}/{len(functions)}")

            print(f"\n🎯 总体成功率: {total_success}/{total_tests} ({total_success/total_tests*100:.1f}%)")

            # 清理
            os.unlink(temp_file.name)

            return total_success / total_tests >= 0.8

        except Exception as e:
            print(f"❌ formulas计算引擎错误: {e}")
            return False

    except ImportError as e:
        print(f"❌ formulas库导入失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 测试过程中出现错误: {e}")
        return False

def test_alternative_libraries():
    """测试其他可选库"""

    print("\n🔍 测试其他统计库...")

    # 测试numpy + scipy统计
    try:
        import numpy as np
        from scipy import stats

        print("✅ numpy + scipy 可用于高级统计")

        # 示例数据
        data = [10, 20, 30, 40, 50, 15, 25, 35, 45, 55]

        # 使用numpy/scipy实现各种统计
        results = {
            "median": float(np.median(data)),
            "std": float(np.std(data, ddof=1)),  # 样本标准差
            "var": float(np.var(data, ddof=1)),   # 样本方差
            "percentile_90": float(np.percentile(data, 90)),
            "quartile_1": float(np.percentile(data, 25)),
        }

        print("📊 numpy/scipy统计结果:")
        for name, value in results.items():
            print(f"   {name}: {value}")

        return True

    except ImportError:
        print("❌ numpy/scipy 不可用")
        return False

if __name__ == "__main__":
    print("🔬 测试高级统计函数库支持...")

    formulas_ok = test_formulas_library()
    numpy_ok = test_alternative_libraries()

    print(f"\n💡 结论:")
    if formulas_ok:
        print("   ✅ formulas库可以很好地支持高级Excel统计函数")
        print("   📈 建议：升级到formulas引擎替代xlcalculator")
    elif numpy_ok:
        print("   ✅ numpy+scipy可以实现所有高级统计功能")
        print("   🧮 建议：扩展基础解析器，用numpy实现高级统计")
    else:
        print("   ⚠️  建议实现专门的excel_get_statistics工具")
