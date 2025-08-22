#!/usr/bin/env python3
"""
演示ID对象属性变化（寻找修改的技能）
"""

from src.server import excel_compare_sheets

def demo_modified_skills():
    """演示修改技能的属性变化"""
    print("=== 寻找修改的技能对象 ===")
    
    result = excel_compare_sheets(
        'D:/tr/svn/trunk/配置表/测试配置/微小/TrSkill.xlsx',
        'TrSkill',
        'D:/tr/svn/trunk/配置表/战斗环境配置/TrSkill.xlsx',
        'TrSkill',
        focus_on_id_changes=True,
        game_friendly_format=True
    )
    
    if result['success'] and 'row_differences' in result:
        row_diffs = result['row_differences']
        
        # 筛选出修改的对象
        modified_objects = [diff for diff in row_diffs if hasattr(diff, 'difference_type') 
                           and str(diff.difference_type) == 'DifferenceType.ROW_MODIFIED']
        
        print(f"🔧 发现 {len(modified_objects)} 个修改的技能对象")
        
        if modified_objects:
            print("\n详细属性变化:")
            for i, diff in enumerate(modified_objects[:3]):  # 只显示前3个
                print(f"\n=== 修改 {i+1}: ID {diff.row_id} ({diff.object_name}) ===")
                
                if hasattr(diff, 'field_differences') and diff.field_differences:
                    for field_diff in diff.field_differences:
                        print(f"  🔄 {field_diff}")
                else:
                    print("  无具体字段差异信息")
        else:
            print("  未发现修改的技能对象")
            
        # 统计变化类型
        added_count = sum(1 for diff in row_diffs if 'ROW_ADDED' in str(diff.difference_type))
        removed_count = sum(1 for diff in row_diffs if 'ROW_REMOVED' in str(diff.difference_type))
        modified_count = len(modified_objects)
        
        print(f"\n📊 变化统计:")
        print(f"  🆕 新增: {added_count} 个")
        print(f"  🗑️ 删除: {removed_count} 个") 
        print(f"  🔧 修改: {modified_count} 个")
        print(f"  📈 总计: {len(row_diffs)} 个ID对象变化")
        
    else:
        print(f"❌ 操作失败或无结构化结果")

if __name__ == "__main__":
    demo_modified_skills()
