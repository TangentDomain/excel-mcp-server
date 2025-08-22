#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试简化后的比较输出 - 验证移除row_data后的效果
"""

import json
from src.core.excel_compare import ExcelComparer
from src.models.types import ComparisonOptions

def test_simplified_comparison():
    """测试简化的比较输出"""
    comparer = ExcelComparer()
    
    # 使用之前的测试文件
    file1 = "D:\\tr\\svn\\trunk\\配置表\\测试配置\\微小\\TrSkill.xlsx"
    file2 = "D:\\tr\\svn\\trunk\\配置表\\战斗环境配置\\TrSkill.xlsx"
    
    print("🔍 开始简化版本的Excel文件比较...")
    print(f"文件1: {file1}")
    print(f"文件2: {file2}")
    
    # 设置比较选项
    options = ComparisonOptions(
        compare_values=True,
        structured_comparison=True,
        focus_on_id_changes=True,
        game_friendly_format=True,
        show_numeric_changes=True,
        ignore_empty_cells=True
    )
    
    # 执行比较
    result = comparer.compare_files(file1, file2, options)
    
    if not result.success:
        print(f"❌ 比较失败: {result.error}")
        return
    
    print(f"✅ 比较成功！发现 {result.data.total_differences} 处差异")
    
    # 找一个有差异的工作表进行详细输出测试
    for sheet_comparison in result.data.sheet_comparisons:
        if sheet_comparison.differences:
            print(f"\n📋 工作表: {sheet_comparison.sheet_name}")
            print(f"差异数量: {len(sheet_comparison.differences)}")
            
            # 只输出前2个差异的简化结果
            for i, diff in enumerate(sheet_comparison.differences[:2]):
                print(f"\n🔸 差异 {i+1}:")
                print(f"   ID: {diff.row_id}")
                print(f"   类型: {diff.difference_type}")
                print(f"   对象: {diff.object_name}")
                print(f"   摘要: {diff.id_based_summary}")
                
                # 检查是否还有row_data（应该没有了）
                if hasattr(diff, 'row_data1') and diff.row_data1:
                    print("   ❌ 仍然包含row_data1 - 优化失败")
                else:
                    print("   ✅ 已移除row_data1")
                
                if hasattr(diff, 'row_data2') and diff.row_data2:
                    print("   ❌ 仍然包含row_data2 - 优化失败")
                else:
                    print("   ✅ 已移除row_data2")
                
                if diff.detailed_field_differences:
                    print(f"   详细字段差异: {len(diff.detailed_field_differences)}个")
                    for field_diff in diff.detailed_field_differences[:3]:  # 只显示前3个
                        print(f"     - {field_diff.field_name}: '{field_diff.old_value}' → '{field_diff.new_value}'")
            
            break  # 只测试一个工作表
    
    print(f"\n📊 数据大小优化效果:")
    # 将结果转换为字典形式进行大小测算
    result_dict = {
        "success": result.success,
        "file1_path": result.file1_path,
        "file2_path": result.file2_path,
        "summary": result.summary,
        "sheet_comparisons": []
    }
    
    for sheet_comp in result.sheet_comparisons:
        sheet_dict = {
            "sheet_name": sheet_comp.sheet_name,
            "differences": []
        }
        
        for diff in sheet_comp.differences:
            diff_dict = {
                "row_id": diff.row_id,
                "difference_type": diff.difference_type.value,
                "detailed_field_differences": [
                    {
                        "field_name": fd.field_name,
                        "old_value": fd.old_value,
                        "new_value": fd.new_value,
                        "change_type": fd.change_type.value if hasattr(fd, 'change_type') else None
                    } for fd in diff.detailed_field_differences
                ],
                "object_name": diff.object_name,
                "id_based_summary": diff.id_based_summary
            }
            sheet_dict["differences"].append(diff_dict)
        
        result_dict["sheet_comparisons"].append(sheet_dict)
    
    json_str = json.dumps(result_dict, ensure_ascii=False, indent=2)
    print(f"JSON总大小: {len(json_str):,} 字符")
    
    # 检查是否还包含row_data字段
    if 'row_data1' in json_str or 'row_data2' in json_str:
        print("❌ JSON中仍包含row_data字段")
    else:
        print("✅ JSON中已完全移除row_data字段")

if __name__ == "__main__":
    test_simplified_comparison()
