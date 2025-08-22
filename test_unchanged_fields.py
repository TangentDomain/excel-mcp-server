#!/usr/bin/env python3
"""
测试unchanged_fields字段的具体内容
"""
import sys
import os
sys.path.append('.')

def test_unchanged_fields():
    """测试unchanged_fields字段的具体内容"""
    print("=== 测试unchanged_fields字段内容 ===")

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

        if result.get('success'):
            data = result.get('data', {})
            row_differences = data.get('row_differences', [])

            print(f"✅ 比较成功！共发现 {len(row_differences)-1} 处差异")
            print(f"🎯 字段定义: {row_differences[0]}")

            # 查找第一个有field_differences的修改行
            for i, row_data in enumerate(row_differences[1:], 1):
                if (row_data[1] == 'row_modified' and
                    row_data[5] is not None and
                    row_data[6] is not None):  # 有field_differences和unchanged_fields

                    print(f"\n📋 第{i}行差异详情 (ID: {row_data[0]}):")
                    print(f"   类型: {row_data[1]}")
                    print(f"   位置: 文件1第{row_data[2]}行 → 文件2第{row_data[3]}行")

                    field_differences = row_data[5]
                    unchanged_fields = row_data[6]

                    print(f"\n🔄 变化字段 ({len(field_differences)} 个):")
                    for j, field_diff in enumerate(field_differences[:3]):  # 只显示前3个
                        print(f"   [{j}] {field_diff[0]}: '{field_diff[1]}' → '{field_diff[2]}' ({field_diff[3]})")

                    print(f"\n⚪ 未变化字段 ({len(unchanged_fields)} 个):")
                    for j, unchanged_field in enumerate(unchanged_fields[:5]):  # 只显示前5个
                        print(f"   [{j}] {unchanged_field[0]}: '{unchanged_field[1]}' (unchanged)")

                    print(f"\n💡 完整对象状态: 变化字段 {len(field_differences)} 个 + 未变化字段 {len(unchanged_fields)} 个 = 总计 {len(field_differences) + len(unchanged_fields)} 个属性")
                    break

        else:
            print(f"❌ 比较失败: {result.get('error', '未知错误')}")

    except Exception as e:
        print(f"❌ 测试过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_unchanged_fields()
