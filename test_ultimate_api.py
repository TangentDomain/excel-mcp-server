#!/usr/bin/env python3
"""
测试超级简化的Excel比较API - 只保留game模式
验证去掉mode参数后的终极简洁版本
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_ultimate_simplified_api():
    """测试终极简化的Excel比较API"""
    print("🎮 测试终极简化的Excel比较API - 游戏开发专用版")
    print("="*65)

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

    # 测试1: excel_compare_files 最简用法
    total_tests += 1
    print(f"\n🔍 测试1: excel_compare_files - 最简用法")
    print(f"  调用: excel_compare_files(file1, file2)")
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

    # 测试2: excel_compare_files 指定列和行
    total_tests += 1
    print(f"\n🔍 测试2: excel_compare_files - 指定ID列和表头行")
    print(f"  调用: excel_compare_files(file1, file2, id_column=1, header_row=1)")
    try:
        result = excel_compare_files(file1, file2, id_column=1, header_row=1)
        if result.get('success'):
            total_diffs = result.get('metadata', {}).get('total_differences', 0)
            print(f"  ✅ 成功! 发现 {total_diffs} 个差异")
            success_count += 1
        else:
            print(f"  ❌ 失败: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"  💥 异常: {str(e)}")

    # 测试3: excel_compare_sheets 最简用法
    total_tests += 1
    print(f"\n🔍 测试3: excel_compare_sheets - 最简用法")
    print(f"  调用: excel_compare_sheets(file1, 'TrSkill', file2, 'TrSkill')")
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

    # 测试4: excel_compare_sheets 指定参数
    total_tests += 1
    print(f"\n🔍 测试4: excel_compare_sheets - 指定ID列和表头行")
    print(f"  调用: excel_compare_sheets(file1, 'TrSkill', file2, 'TrSkill', id_column=1, header_row=1)")
    try:
        result = excel_compare_sheets(file1, "TrSkill", file2, "TrSkill", id_column=1, header_row=1)
        if result.get('success'):
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

    # 汇总结果
    print(f"\n" + "="*65)
    print(f"📊 测试结果汇总:")
    print(f"  - 总测试数: {total_tests}")
    print(f"  - 成功数: {success_count}")
    print(f"  - 成功率: {success_count/total_tests*100:.1f}%")

    if success_count == total_tests:
        print(f"🎉 所有测试通过! 终极简化API完美运行!")
        return True
    else:
        print(f"⚠️  部分测试失败，需要检查")
        return False

def show_ultimate_api_comparison():
    """展示终极API简化对比"""
    print(f"\n📋 终极API简化对比:")
    print(f"="*65)

    print(f"🔴 之前的简化版 (仍有mode参数):")
    print(f"  excel_compare_files(")
    print(f"    file1_path, file2_path,")
    print(f"    id_column=1, header_row=1,")
    print(f"    mode='game'  # 还需要选择模式")
    print(f"  )")
    print(f"  📊 参数数量: 5个")

    print(f"\n🟢 终极简化版 (游戏开发专用):")
    print(f"  excel_compare_files(")
    print(f"    file1_path, file2_path,")
    print(f"    id_column=1, header_row=1")
    print(f"    # 没有mode参数，直接游戏开发专用")
    print(f"  )")
    print(f"  📊 参数数量: 4个")

    print(f"\n🎯 最简用法:")
    print(f"  excel_compare_files('old.xlsx', 'new.xlsx')")
    print(f"  📊 参数数量: 仅2个!")

    print(f"\n✨ 终极改进效果:")
    print(f"  - 参数减少: 5个 → 4个 (再减少20%)")
    print(f"  - 模式选择: 无需选择，专为游戏开发优化")
    print(f"  - 使用体验: 开箱即用，零配置")
    print(f"  - 专业聚焦: 100%专注游戏配置表对比")
    print(f"  - 最简调用: 只需要文件路径，其他都有智能默认值")

def show_game_focused_features():
    """展示游戏开发专用功能特性"""
    print(f"\n🎮 游戏开发专用功能特性:")
    print(f"="*65)

    print(f"✅ 自动启用的功能:")
    print(f"  🎯 ID对象变化跟踪 - 自动识别新增、删除、修改的游戏对象")
    print(f"  📊 数值变化分析 - 显示攻击力、血量等数值的变化量和百分比")
    print(f"  🏗️ 结构化数据比较 - 按行比较，而非单元格级对比")
    print(f"  🎨 游戏友好格式 - 输出格式专为游戏策划和程序员优化")
    print(f"  🚀 性能优化 - 忽略格式和公式，专注数据内容")

    print(f"\n❌ 自动禁用的功能 (减少干扰):")
    print(f"  📝 公式比较 - 游戏配置表通常不涉及复杂公式")
    print(f"  🎨 格式比较 - 专注数据内容，忽略视觉格式")
    print(f"  📍 位置信息 - 隐藏单元格位置，专注业务对象")

    print(f"\n💡 智能默认设置:")
    print(f"  📋 表头行: 第1行 (游戏配置表的标准格式)")
    print(f"  🆔 ID列: 第1列 (游戏对象ID的标准位置)")
    print(f"  🔤 大小写敏感: 是 (游戏ID通常区分大小写)")
    print(f"  🗑️ 忽略空单元格: 是 (减少噪音)")

if __name__ == "__main__":
    print("🎯 Excel比较API - 终极简化验证")
    print("="*65)

    # 显示API对比
    show_ultimate_api_comparison()

    # 显示游戏专用功能
    show_game_focused_features()

    # 执行测试
    success = test_ultimate_simplified_api()

    print(f"\n" + "="*65)
    if success:
        print("🎉 终极简化成功! 专为游戏开发打造的完美API!")
        print("🎮 现在这是一个100%专注游戏开发的Excel比较工具!")
    else:
        print("❌ 终极简化测试未完全通过")
    print("="*65)
