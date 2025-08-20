#!/usr/bin/env python3
"""
分析TrSkill.xlsx文件中的赫卡忒信息
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from server import excel_list_sheets, excel_regex_search, excel_get_range

def analyze_trskill_file():
    """分析TrSkill.xlsx文件"""
    file_path = 'TrSkill.xlsx'
    print("📋 分析TrSkill.xlsx文件:")
    print("=" * 40)

    # 查看工作表列表
    print("📊 工作表列表:")
    sheets_result = excel_list_sheets(file_path)
    if sheets_result['success']:
        print(f"共有 {sheets_result['total_sheets']} 个工作表:")
        for sheet in sheets_result['sheets']:
            active = "🎯" if sheet['is_active'] else "📄"
            print(f"  {active} {sheet['name']} (数据范围: {sheet['max_column_letter']}{sheet['max_row']})")
    else:
        print(f"❌ 获取工作表失败: {sheets_result['error']}")
        return

    # 搜索"赫卡忒"
    print("\n🔍 搜索'赫卡忒':")
    search_result = excel_regex_search(file_path, '赫卡忒')
    if search_result['success']:
        if search_result['total_matches'] > 0:
            print(f"✅ 找到 {search_result['total_matches']} 个匹配:")
            for i, match in enumerate(search_result['matches'], 1):
                print(f"  [{i}] 工作表: {match['sheet']}")
                print(f"      单元格: {match['cell']}")
                print(f"      内容: {match['value']}")
                print("      ---")
        else:
            print("❌ 未找到'赫卡忒'相关内容")
    else:
        print(f"❌ 搜索失败: {search_result['error']}")

    # 如果找到了，获取更多上下文
    if search_result['success'] and search_result['total_matches'] > 0:
        print("\n📋 获取相关上下文数据:")
        for match in search_result['matches']:
            sheet_name = match['sheet']
            cell = match['cell']
            print(f"\n--- {sheet_name}工作表上下文 ---")

            # 获取该行的更多数据
            cell_row = int(''.join(filter(str.isdigit, cell)))
            row_result = excel_get_range(file_path, f"{sheet_name}!{cell_row}:{cell_row}")
            if row_result['success'] and row_result['data']:
                row_values = [cell_info['value'] for cell_info in row_result['data'][0] if cell_info['value']]
                print(f"第{cell_row}行数据: {row_values}")

if __name__ == "__main__":
    analyze_trskill_file()
