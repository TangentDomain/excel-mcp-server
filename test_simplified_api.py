#!/usr/bin/env python3
"""
测试简化的Excel比较API - 消除历史包袱版本
验证新的mode参数和简化的参数列表
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_simplified_api():
    """测试简化的Excel比较API"""
    print("🚀 测试简化的Excel比较API")
    print("="*60)

    # 导入简化后的方法
    try:
        from src.server import excel_compare_files, excel_compare_sheets
        print("✅ 导入成功")
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        return False

    # 测试文件路径
    file1 = r"D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx"
    file2 = r"D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx"

    print(f"\n📂 测试文件:")
    print(f"  - 文件1: {file1}")
    print(f"  - 文件2: {file2}")

    success_count = 0
    total_tests = 0

    # 测试1: excel_compare_files 默认模式 (game)
    total_tests += 1
    print(f"\n🔍 测试1: excel_compare_files - 默认模式 (game)")
    try:
        result = excel_compare_files(file1, file2)
        if result.get('success'):
            total_diffs = result.get('metadata', {}).get('total_differences', 0)
            print(f"  ✅ 成功! 发现 {total_diffs} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 测试2: excel_compare_files quick模式
    total_tests += 1
    print(f"\n🔍 测试2: excel_compare_files - quick模式")
    try:
        result = excel_compare_files(file1, file2, mode='quick')
        if result.get('success'):
            total_diffs = result.get('metadata', {}).get('total_differences', 0)
            print(f"  ✅ 成功! 发现 {total_diffs} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 测试3: excel_compare_files detailed模式
    total_tests += 1
    print(f"\n🔍 测试3: excel_compare_files - detailed模式")
    try:
        result = excel_compare_files(file1, file2, mode='detailed')
        if result.get('success'):
            total_diffs = result.get('metadata', {}).get('total_differences', 0)
            print(f"  ✅ 成功! 发现 {total_diffs} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 测试4: excel_compare_sheets 默认模式
    total_tests += 1
    print(f"\n🔍 测试4: excel_compare_sheets - 默认模式 (game)")
    try:
        result = excel_compare_sheets(file1, "TrSkill", file2, "TrSkill")
        if result.get('success'):
            # 检查工作表比较结果
            data = result.get('data', {})
            differences = 0
            if hasattr(data, '__dict__'):
                differences = getattr(data, 'total_differences', 0)
            elif isinstance(data, dict) and 'total_differences' in data:
                differences = data['total_differences']
            print(f"  ✅ 成功! TrSkill工作表有 {differences} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 测试5: 自定义参数
    total_tests += 1
    print(f"\n🔍 测试5: 自定义参数 (id_column=1, header_row=1)")
    try:
        result = excel_compare_files(file1, file2, id_column=1, header_row=1, mode='game')
        if result.get('success'):
            total_diffs = result.get('metadata', {}).get('total_differences', 0)
            print(f"  ✅ 成功! 发现 {total_diffs} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 汇总结果
    print(f"\n" + "="*60)
    print(f"📊 测试结果汇总:")
    print(f"  - 总测试数: {total_tests}")
    print(f"  - 成功数: {success_count}")
    print(f"  - 成功率: {success_count/total_tests*100:.1f}%")

    if success_count == total_tests:
        print(f"🎉 所有测试通过! 简化API工作完美!")
        return True
    else:
        print(f"⚠️  部分测试失败，需要检查")
        return False

def show_api_comparison():
    """展示API简化前后对比"""
    print(f"\n📋 API简化对比:")
    print(f"="*60)

    print(f"🔴 简化前 (历史包袱版):")
    print(f"  excel_compare_files(")
    print(f"    file1_path, file2_path,")
    print(f"    compare_values=True, compare_formulas=False,")
    print(f"    compare_formats=False, ignore_empty_cells=True,")
    print(f"    case_sensitive=True, structured_comparison=True,")
    print(f"    header_row=1, id_column=1,")
    print(f"    show_numeric_changes=True, game_friendly_format=True,")
    print(f"    focus_on_id_changes=True")
    print(f"  )")
    print(f"  📊 参数数量: 13个")

    print(f"\n🟢 简化后 (消除包袱版):")
    print(f"  excel_compare_files(")
    print(f"    file1_path, file2_path,")
    print(f"    id_column=1, header_row=1,")
    print(f"    mode='game'  # 'quick', 'detailed', 'game'")
    print(f"  )")
    print(f"  📊 参数数量: 5个")

    print(f"\n✨ 改进效果:")
    print(f"  - 参数减少: 13个 → 5个 (减少61%)")
    print(f"  - 复杂度降低: 用mode统一控制所有细节参数")
    print(f"  - 易用性提升: 常用场景只需要指定文件路径")
    print(f"  - 向后兼容: 通过mode参数实现所有原有功能")

if __name__ == "__main__":
    print("🎯 Excel比较API简化验证")
    print("="*60)

    # 显示API对比
    show_api_comparison()

    # 执行测试
    success = test_simplified_api()

    print(f"\n" + "="*60)
    if success:
        print("🎉 简化成功! 历史包袱已消除，API更简洁易用!")
    else:
        print("❌ 简化测试未完全通过")
    print("="*60)
