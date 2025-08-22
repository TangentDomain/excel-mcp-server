#!/usr/bin/env python3
"""
直接测试excel_compare_files工具
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_excel_compare_directly():
    """直接调用excel_compare_files函数"""
    print("🧪 直接测试excel_compare_files函数...")

    # 导入服务器中的工具函数
    try:
        from src.server import excel_compare_files
    except ImportError:
        print("❌ 无法导入excel_compare_files，尝试从模块导入...")
        # 尝试直接从文件导入
        import importlib.util
        spec = importlib.util.spec_from_file_location("server", "src/server.py")
        server_module = importlib.util.module_from_spec(spec)
        sys.modules["server"] = server_module
        spec.loader.exec_module(server_module)
        excel_compare_files = getattr(server_module, 'excel_compare_files', None)
        if not excel_compare_files:
            print("❌ 找不到excel_compare_files函数")
            return False

    # 测试文件路径
    file1 = r"D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx"
    file2 = r"D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx"

    try:
        print(f"📂 比较文件:")
        print(f"  - 文件1: {file1}")
        print(f"  - 文件2: {file2}")
        print()

        # 调用比较函数
        result = excel_compare_files(
            file1_path=file1,
            file2_path=file2,
            header_row=1,
            id_column=1,
            case_sensitive=True
        )

        print(f"📋 比较结果类型: {type(result)}")

        detailed_field_count = 0  # 初始化变量

        if isinstance(result, dict):
            print(f"✅ 比较成功!")
            print(f"📊 成功状态: {result.get('success', False)}")
            print(f"📊 结果键: {list(result.keys())}")

            # 打印所有键的内容概览
            for key, value in result.items():
                if key != 'data':  # data可能太大
                    print(f"   {key}: {value}")
                else:
                    print(f"   {key}: {type(value)} (长度: {len(value) if hasattr(value, '__len__') else 'N/A'})")

            # 检查data字段
            if 'data' in result and result['data']:
                data = result['data']
                print(f"📊 数据类型: {type(data)}")

                if isinstance(data, dict):
                    print(f"📊 数据键: {list(data.keys())}")

                    # 检查sheet_comparisons
                    if 'sheet_comparisons' in data:
                        sheet_comparisons = data['sheet_comparisons']
                        print(f"📊 工作表比较数: {len(sheet_comparisons)}")

                        for i, sheet_comp in enumerate(sheet_comparisons):
                            sheet_name = sheet_comp.get('sheet_name', f'Sheet_{i+1}')
                            differences = sheet_comp.get('differences', [])

                            if differences:
                                print(f"\n📋 工作表 '{sheet_name}': {len(differences)} 个差异")
                                print(f"    差异类型: {type(differences)}")

                                # 检查前几个差异的详细字段变化
                                for j, diff in enumerate(differences[:3]):
                                    if isinstance(diff, dict):
                                        print(f"  🔍 差异 {j+1}:")
                                        print(f"    差异键: {list(diff.keys())}")

                                        row_id = diff.get('row_id', 'N/A')
                                        object_name = diff.get('object_name', 'N/A')
                                        print(f"    ID {row_id} ({object_name})")

                                        # 检查是否有详细字段差异
                                        if 'detailed_field_differences' in diff:
                                            detailed_fields = diff['detailed_field_differences']
                                            print(f"    详细字段变化数: {len(detailed_fields) if detailed_fields else 0}")

                                            if detailed_fields:
                                                for field_diff in detailed_fields[:3]:
                                                    if isinstance(field_diff, dict):
                                                        field_name = field_diff.get('field_name', 'N/A')
                                                        old_val = field_diff.get('old_value', 'N/A')
                                                        new_val = field_diff.get('new_value', 'N/A')
                                                        change_type = field_diff.get('change_type', 'N/A')
                                                        print(f"      🔧 {field_name} ({change_type}): '{old_val}' → '{new_val}'")
                                                        detailed_field_count += 1
                                        else:
                                            print(f"    ⚠️ 没有详细字段变化属性")
                                    else:
                                        print(f"  🔍 差异 {j+1}: {type(diff)}")

                                # 只看第一个有差异的工作表
                                break
                    else:
                        print("⚠️ data中没有sheet_comparisons")
                else:
                    print(f"📊 data不是字典类型: {type(data)}")
            else:
                print("⚠️ 没有数据字段或数据为空")

            print(f"\n📈 统计:")
            print(f"   - 详细字段差异数: {detailed_field_count}")
            print(f"   - 支持ID-属性跟踪: {'✅' if detailed_field_count > 0 else '⚠️'}")

            return detailed_field_count > 0

        else:
            print(f"⚠️ 结果不是字典类型: {result}")
            return False

    except Exception as e:
        print(f"💥 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 直接Excel比较功能测试")
    print("=" * 60)

    success = test_excel_compare_directly()

    print("\n" + "=" * 60)
    if success:
        print("🎉 详细属性变化跟踪功能正常!")
    else:
        print("❌ 详细属性变化跟踪功能需要检查")
    print("=" * 60)
