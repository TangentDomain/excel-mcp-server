#!/usr/bin/bin/env python3
"""
验证REQ-025返回值统一情况
检查核心工具是否都使用了统一的{success, data, meta, message}格式
"""

import sys
import os
sys.path.insert(0, 'src')

import tempfile
import pandas as pd
from excel_mcp_server_fastmcp.server import (
    excel_list_sheets,
    excel_get_headers,
    excel_get_range,
    excel_describe_table,
    excel_query,
    excel_search,
    excel_create_file,
    excel_insert_rows,
    excel_delete_rows,
    excel_assess_data_impact
)

def create_test_excel():
    """创建测试Excel文件"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_file = f.name
    
    # 创建测试数据
    data = {
        '技能名称': ['火球术', '冰箭', '治疗术', '闪电链'],
        '职业限制': ['法师', '法师', '牧师', '法师'],
        '伤害': [100, 80, 0, 90],
        '冷却时间': [5, 3, 8, 6],
        '类型': ['攻击', '攻击', '治疗', '攻击']
    }
    
    df = pd.DataFrame(data)
    df.to_excel(temp_file, sheet_name='技能', index=False)
    
    return temp_file

def test_return_format_unified():
    """测试返回格式统一性"""
    test_file = create_test_excel()
    print(f"测试文件: {test_file}")
    
    # 需要测试的工具列表
    tools_to_test = [
        ('excel_list_sheets', lambda: excel_list_sheets(test_file)),
        ('excel_get_headers', lambda: excel_get_headers(test_file, "技能")),
        ('excel_get_range', lambda: excel_get_range(test_file, "技能!A1:C5")),
        ('excel_describe_table', lambda: excel_describe_table(test_file, "技能")),
        ('excel_query', lambda: excel_query(test_file, "SELECT 技能名称, 伤害 FROM 技能 WHERE 类型 = '攻击'", include_headers=True)),
        ('excel_search', lambda: excel_search(test_file, "法师", "技能")),
        ('excel_create_file', lambda: excel_create_file(test_file.replace('.xlsx', '_test.xlsx'))),
        ('excel_insert_rows', lambda: excel_insert_rows(test_file, "技能", row_index=10, count=1)),
        ('excel_delete_rows', lambda: excel_delete_rows(test_file, "技能", row_index=5, count=1)),
    ]
    
    unified_count = 0
    total_count = len(tools_to_test)
    
    print("\n=== 返回格式统一性检查 ===")
    
    for tool_name, tool_func in tools_to_test:
        try:
            result = tool_func()
            is_unified = (
                isinstance(result, dict) and
                'success' in result and
                'message' in result
            )
            
            # 成功时需要完整的 {success, message, data, meta}
            if is_unified and result.get('success') is True:
                has_data = 'data' in result
                has_meta = 'meta' in result
                if has_data and has_meta:
                    pass  # 完全统一
                else:
                    is_unified = False
            
            if is_unified:
                status = "✅ 统一"
                unified_count += 1
            else:
                status = "❌ 不统一"
                
            print(f"{tool_name}: {status}")
            
            if not is_unified:
                print(f"  返回内容: {result}")
                
        except Exception as e:
            print(f"{tool_name}: ❌ 异常 - {e}")
    
    print(f"\n=== 检查结果 ===")
    print(f"统一格式工具数: {unified_count}/{total_count}")
    print(f"统一率: {unified_count/total_count*100:.1f}%")
    
    # 清理测试文件
    if os.path.exists(test_file):
        os.unlink(test_file)
    
    return unified_count == total_count

if __name__ == "__main__":
    result = test_return_format_unified()
    print(f"\nREQ-025 返回值统一检查结果: {'✅ 全部统一' if result else '❌ 需要改进'}")