#!/usr/bin/env python3
"""
测试ID对象变化专注模式
"""

from src.server import excel_compare_sheets

def test_id_focus():
    """测试ID对象变化专注模式"""
    print("=== ID对象变化专注模式测试 ===")
    
    result = excel_compare_sheets(
        'D:/tr/svn/trunk/配置表/测试配置/微小/TrSkill.xlsx',
        'TrSkill',
        'D:/tr/svn/trunk/配置表/战斗环境配置/TrSkill.xlsx',
        'TrSkill',
        focus_on_id_changes=True,
        game_friendly_format=True
    )
    
    if result['success']:
        print("✅ 比较成功")
        
        # 查看结果结构
        print(f"结果类型: {type(result)}")
        print(f"主要键: {list(result.keys())}")
        
        # 直接检查是否有row_differences（结构化比较结果）
        if 'row_differences' in result:
            row_diffs = result['row_differences']
            print(f"🔍 发现 {len(row_diffs)} 个ID对象变化")
            
            # 显示前几个变化
            for i, diff in enumerate(row_diffs[:5]):
                print(f"\n变化 {i+1}:")
                
                # 检查diff是对象还是字典
                if hasattr(diff, 'row_id'):
                    # 对象格式
                    print(f"  ID: {diff.row_id}")
                    print(f"  类型: {diff.difference_type}")
                    
                    if hasattr(diff, 'id_based_summary') and diff.id_based_summary:
                        print(f"  📝 {diff.id_based_summary}")
                    
                    if hasattr(diff, 'object_name') and diff.object_name:
                        print(f"  对象名: {diff.object_name}")
                    
                    # 显示字段差异
                    if hasattr(diff, 'field_differences') and diff.field_differences:
                        field_diffs = diff.field_differences
                        print(f"  字段变化 ({len(field_diffs)}个):")
                        for field_diff in field_diffs[:3]:  # 只显示前3个
                            print(f"    - {field_diff}")
                        if len(field_diffs) > 3:
                            print(f"    ... 还有 {len(field_diffs) - 3} 个变化")
                else:
                    # 字典格式
                    print(f"  ID: {diff.get('row_id', '?')}")
                    print(f"  类型: {diff.get('difference_type', 'unknown')}")
                    
                    if 'id_based_summary' in diff and diff['id_based_summary']:
                        print(f"  📝 {diff['id_based_summary']}")
                    
                    if 'object_name' in diff and diff['object_name']:
                        print(f"  对象名: {diff['object_name']}")
                    
                    # 显示字段差异
                    field_diffs = diff.get('field_differences', [])
                    if field_diffs:
                        print(f"  字段变化 ({len(field_diffs)}个):")
                        for field_diff in field_diffs[:3]:  # 只显示前3个
                            print(f"    - {field_diff}")
                        if len(field_diffs) > 3:
                            print(f"    ... 还有 {len(field_diffs) - 3} 个变化")
            
            if len(row_diffs) > 5:
                print(f"\n... 还有 {len(row_diffs) - 5} 个ID对象变化")
                
        else:
            print("❌ 这不是结构化比较结果")
            print("  可能是传统的单元格比较结果")
            
            # 检查传统比较结果
            if 'differences' in result:
                diffs = result['differences']
                print(f"  发现 {len(diffs)} 个单元格差异")
            else:
                print("  也没有找到传统比较结果")
            
    else:
        print(f"❌ 比较失败: {result.get('error', '未知错误')}")

if __name__ == "__main__":
    test_id_focus()
