#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试优化后的excel_compare_sheets紧凑数组格式API
"""

import sys
import os
import json
from pathlib import Path

# 添加src路径以导入模块
project_root = Path(__file__).parent
src_path = project_root / "src"
sys.path.insert(0, str(src_path))

def test_compact_array_format():
    """测试紧凑数组格式的效果"""
    print("=== 测试紧凑数组格式API ===")

    try:
        from src.server import excel_compare_sheets

        # 测试文件路径
        file1 = r"D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx"
        file2 = r"D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx"

        if not (os.path.exists(file1) and os.path.exists(file2)):
            print("❌ 测试文件不存在，请检查路径")
            return

        print(f"📁 文件1: {file1}")
        print(f"📁 文件2: {file2}")
        print("🔍 开始比较TrSkillEffect工作表...")

        # 执行比较
        result = excel_compare_sheets(
            file1, "TrSkillEffect",
            file2, "TrSkillEffect"
        )

        print(f"✅ 比较完成！")
        print(f"📊 成功状态: {result.get('success')}")
        print(f"📝 消息: {result.get('message')}")

        # 分析数据结构
        data = result.get('data', {})
        row_differences = data.get('row_differences', [])

        if row_differences and len(row_differences) > 0:
            print(f"\n🎯 数组格式分析:")
            print(f"   总差异数: {data.get('total_differences', 0)}")
            print(f"   数组行数: {len(row_differences)}")

            if len(row_differences) > 1:
                # 显示字段定义（第一行）
                field_definitions = row_differences[0]
                print(f"   字段定义: {field_definitions}")

                # 显示前几个实际数据行
                print(f"\n📋 前3个差异示例:")
                for i in range(1, min(4, len(row_differences))):
                    row_data = row_differences[i]
                    print(f"   行{i}: {row_data}")

                # 计算空间节省效果
                original_size = estimate_original_format_size(data.get('total_differences', 0))
                current_size = len(json.dumps(row_differences))
                savings = ((original_size - current_size) / original_size * 100) if original_size > 0 else 0

                print(f"\n💾 空间优化效果:")
                print(f"   估计原格式大小: {original_size:,} 字符")
                print(f"   当前数组格式: {current_size:,} 字符")
                print(f"   空间节省: {savings:.1f}%")
        else:
            print("📋 无差异数据")

        # 保存完整结果到文件供分析
        output_file = project_root / "compact_array_test_result.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"\n💾 完整结果已保存到: {output_file}")

        return result

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return None

def estimate_original_format_size(total_differences):
    """估算原始对象格式的大小"""
    # 每个差异对象大约包含的字符数（键名+值）
    base_overhead = 200  # 基础键名开销
    field_diff_overhead = 100  # 每个字段差异的开销

    # 假设平均每个差异有2个字段差异
    estimated_size = total_differences * (base_overhead + 2 * field_diff_overhead)
    return estimated_size

def analyze_field_definitions(row_differences):
    """分析字段定义和数据结构"""
    if not row_differences or len(row_differences) == 0:
        return

    print("\n🔍 详细结构分析:")

    # 字段定义
    if len(row_differences) > 0:
        field_definitions = row_differences[0]
        print(f"字段定义 (索引 → 含义):")
        for i, field_name in enumerate(field_definitions):
            print(f"  [{i}] → {field_name}")

    # 统计不同类型的差异
    if len(row_differences) > 1:
        diff_types = {}
        field_diff_counts = []

        for row_data in row_differences[1:]:
            if len(row_data) >= 2:
                diff_type = row_data[1]  # difference_type在索引1
                diff_types[diff_type] = diff_types.get(diff_type, 0) + 1

                # 统计字段差异数量
                if len(row_data) >= 6 and row_data[5]:  # field_differences在索引5
                    field_diff_counts.append(len(row_data[5]))

        print(f"\n📊 差异类型统计:")
        for diff_type, count in diff_types.items():
            print(f"  {diff_type}: {count} 个")

        if field_diff_counts:
            avg_field_diffs = sum(field_diff_counts) / len(field_diff_counts)
            print(f"  平均字段差异数: {avg_field_diffs:.1f}")

if __name__ == "__main__":
    result = test_compact_array_format()

    if result and result.get('data', {}).get('row_differences'):
        analyze_field_definitions(result['data']['row_differences'])
        print("\n🎉 紧凑数组格式测试完成！")
    else:
        print("\n❌ 测试未成功完成")
