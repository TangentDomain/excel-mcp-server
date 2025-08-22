#!/usr/bin/env python3
"""
使用简化接口比较TrSkill配置表
"""

from src.server import excel_compare_files

def main():
    try:
        print("🚀 开始比较 TrSkill 配置表...")
        print("文件1: 测试配置/微小")
        print("文件2: 战斗环境配置")
        print()

        result = excel_compare_files(
            r'D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx',
            r'D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx'
        )

        print("✅ 比较完成")
        print(f"成功: {result.get('success', False)}")

        if result.get('success'):
            # 结果在data字段中
            data = result.get('data', {})

            print(f"\n🔍 数据字段:")
            for key in data.keys():
                print(f"  - {key}: {type(data[key])}")

            # 获取比较结果
            total_diff = data.get('total_differences', 0)
            print(f"\n差异总数: {total_diff}")

            sheets = data.get('sheet_comparisons', [])
            print(f"工作表数: {len(sheets)}")

            # 显示详细结果
            for sheet in sheets:
                sheet_name = sheet.get('sheet_name', 'Unknown')
                print(f"\n📋 工作表: {sheet_name}")

                if 'summary' in sheet:
                    summary = sheet['summary']
                    print(f"  • 新增对象: {summary.get('added_rows', 0)}")
                    print(f"  • 删除对象: {summary.get('removed_rows', 0)}")
                    print(f"  • 修改对象: {summary.get('modified_rows', 0)}")
                    print(f"  • 总差异数: {summary.get('total_differences', 0)}")

                # 显示前5个ID对象变化
                if 'row_differences' in sheet:
                    row_diffs = sheet['row_differences'][:5]
                    if row_diffs:
                        print("  前5个ID对象变化:")
                        for i, diff in enumerate(row_diffs, 1):
                            if 'id_based_summary' in diff:
                                print(f"    {i}. {diff['id_based_summary']}")
                            else:
                                # 显示其他有用信息
                                change_type = diff.get('change_type', 'unknown')
                                row_id = diff.get('row_id', 'N/A')
                                print(f"    {i}. {change_type}: ID {row_id}")

        return result

    except Exception as e:
        print(f"❌ 比较失败: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    result = main()
