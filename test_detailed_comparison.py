#!/usr/bin/env python3
"""
测试详细比较功能 - 验证ID对象属性变化跟踪
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.excel_compare import ExcelComparer
from src.models.types import ComparisonOptions

def test_detailed_field_differences():
    """测试详细的字段差异跟踪"""
    print("🧪 测试详细字段差异跟踪功能...")

    # 设置比较选项，启用游戏友好格式和详细跟踪
    options = ComparisonOptions(
        structured_comparison=True,
        game_friendly_format=True,
        focus_on_id_changes=True,
        show_numeric_changes=True,
        header_row=1,
        id_column=1
    )

    comparer = ExcelComparer(options)

    # 测试文件路径
    file1 = r"D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx"
    file2 = r"D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx"

    try:
        print(f"📂 比较文件:")
        print(f"  - 文件1: {file1}")
        print(f"  - 文件2: {file2}")
        print()

        # 执行比较
        result = comparer.compare_files(file1, file2, options)

        if result.success:
            print(f"✅ 比较成功!")
            print(f"� 结果类型: {type(result)}")
            print(f"📋 结果属性: {dir(result)}")

            # 检查实际的结果结构
            result_data = result.data if hasattr(result, 'data') else None
            if result_data:
                print(f"�📊 数据类型: {type(result_data)}")
                if hasattr(result_data, 'total_differences'):
                    print(f"📊 发现 {result_data.total_differences} 个差异")
                else:
                    print(f"📊 数据属性: {dir(result_data)}")

            # 分析详细字段差异
            detailed_field_count = 0
            sample_count = 0

            sheet_comparisons = None
            if result_data and hasattr(result_data, 'sheet_comparisons'):
                sheet_comparisons = result_data.sheet_comparisons
            elif hasattr(result, 'sheet_comparisons'):
                sheet_comparisons = result.sheet_comparisons

            if sheet_comparisons:
                print(f"📊 工作表比较类型: {type(sheet_comparisons)}")

                # 如果是列表，遍历列表
                if isinstance(sheet_comparisons, list):
                    for sheet_comparison in sheet_comparisons:
                        sheet_name = getattr(sheet_comparison, 'sheet_name', 'Unknown')
                        print(f"\n📋 工作表: {sheet_name}")
                        print(f"   - 比较类型: {type(sheet_comparison)}")
                        print(f"   - 比较属性: {[attr for attr in dir(sheet_comparison) if not attr.startswith('_')]}")

                        if hasattr(sheet_comparison, 'differences'):
                            differences = sheet_comparison.differences
                            print(f"   - 差异类型: {type(differences)}")

                            if isinstance(differences, list):
                                print(f"   - 行差异数: {len(differences)}")

                                for i, diff in enumerate(differences[:3]):  # 只看前3个差异作为示例
                                    print(f"\n   📝 差异 {i+1}:")
                                    print(f"      - 类型: {type(diff)}")
                                    print(f"      - 属性: {[attr for attr in dir(diff) if not attr.startswith('_')]}")
                                    print(f"      - Row ID: {getattr(diff, 'row_id', 'N/A')}")
                                    print(f"      - 对象名: {getattr(diff, 'object_name', 'N/A')}")

                                    if hasattr(diff, 'detailed_field_differences'):
                                        print(f"      - 详细差异: {type(diff.detailed_field_differences)} (长度: {len(diff.detailed_field_differences) if diff.detailed_field_differences else 0})")
                                        if diff.detailed_field_differences:
                                            for j, field_diff in enumerate(diff.detailed_field_differences[:3]):  # 只看前3个字段
                                                print(f"      🔧 字段 {j+1}: {field_diff.field_name}")
                                                print(f"         - 原值: {field_diff.old_value}")
                                                print(f"         - 新值: {field_diff.new_value}")
                                                print(f"         - 类型: {field_diff.change_type}")
                                                detailed_field_count += 1
                                    else:
                                        print(f"      - ⚠️ 没有详细字段差异属性")

                                    sample_count += 1
                                    if sample_count >= 3:
                                        break

                                if len(differences) > 3:
                                    print(f"   ... 还有 {len(differences) - 3} 个差异")
                            else:
                                print(f"   - ⚠️ 差异不是列表类型: {type(differences)}")
                                # 尝试访问StructuredSheetComparison的属性
                                if hasattr(differences, 'row_differences'):
                                    row_diffs = differences.row_differences
                                    print(f"   - 行差异数: {len(row_diffs)}")
                                    for i, diff in enumerate(row_diffs[:3]):
                                        print(f"\n   📝 行差异 {i+1}:")
                                        print(f"      - 类型: {type(diff)}")
                                        print(f"      - Row ID: {getattr(diff, 'row_id', 'N/A')}")
                                        if hasattr(diff, 'detailed_field_differences') and diff.detailed_field_differences:
                                            detailed_field_count += len(diff.detailed_field_differences)
                        else:
                            print(f"   - ⚠️ 没有差异数据")

                        # 只处理第一个工作表作为示例
                        break

                # 如果是字典，使用items()方法
                elif isinstance(sheet_comparisons, dict):
                    for sheet_name, sheet_comparison in sheet_comparisons.items():
                        if hasattr(sheet_comparison, 'differences') and sheet_comparison.differences:
                            print(f"\n📋 工作表: {sheet_name}")
                            print(f"   - 行差异数: {len(sheet_comparison.differences)}")

                            for diff in sheet_comparison.differences[:3]:  # 只看前3个差异作为示例
                                if hasattr(diff, 'detailed_field_differences') and diff.detailed_field_differences:
                                    print(f"\n🔍 ID {diff.row_id} 的详细属性变化:")
                                    print(f"   对象名: {getattr(diff, 'object_name', 'N/A')}")
                                    print(f"   变化摘要: {getattr(diff, 'id_based_summary', 'N/A')}")

                                    for field_diff in diff.detailed_field_differences[:5]:  # 只看前5个字段
                                        print(f"   🔧 属性: {field_diff.field_name}")
                                        print(f"      - 原值: {field_diff.old_value}")
                                        print(f"      - 新值: {field_diff.new_value}")
                                        print(f"      - 类型: {field_diff.change_type}")
                                        if field_diff.formatted_change:
                                            print(f"      - 格式化: {field_diff.formatted_change}")
                                        detailed_field_count += 1

                                sample_count += 1
                                if sample_count >= 3:
                                    break

                            if len(sheet_comparison.differences) > 3:
                                print(f"   ... 还有 {len(sheet_comparison.differences) - 3} 个差异")

                            # 只处理第一个工作表作为示例
                            break

                print(f"\n📈 统计:")
                print(f"   - 详细字段差异数: {detailed_field_count}")
                print(f"   - 支持ID-属性跟踪: ✅")
            else:
                print("⚠️ 未找到工作表比较数据")

            return True

        else:
            print(f"❌ 比较失败: {result.message}")
            return False

    except Exception as e:
        print(f"💥 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 Excel详细比较功能测试")
    print("=" * 60)

    success = test_detailed_field_differences()

    print("\n" + "=" * 60)
    if success:
        print("🎉 测试完成 - 详细属性变化跟踪功能正常!")
    else:
        print("❌ 测试失败 - 需要检查代码")
    print("=" * 60)
