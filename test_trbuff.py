#!/usr/bin/env python3
"""
测试TrBuff.xlsx文件的脚本
"""

from src.server import excel_list_sheets
import json

def test_trbuff():
    file_path = r'D:\tr\svn\trunk\配置表\战斗环境配置\TrBuff.xlsx'
    print("=== TrBuff.xlsx MCP Excel 测试分析 ===")
    print(f"文件路径: {file_path}")
    print()

    try:
        result = excel_list_sheets(file_path)

        if result['success']:
            print("✅ 解析状态: 成功")
            print(f"📊 工作表总数: {result['total_sheets']}")
            print(f"🎯 活动工作表: {result['active_sheet']}")
            print()

            print("📋 工作表详细信息:")
            print("-" * 80)

            for i, sheet in enumerate(result['sheets_with_headers'], 1):
                print(f"{i:2d}. 工作表: {sheet['name']}")
                print(f"    字段数量: {sheet['header_count']}个")

                # 显示前5个表头
                headers = sheet['headers']
                if len(headers) <= 5:
                    print(f"    表头列表: {headers}")
                else:
                    print(f"    前5个表头: {headers[:5]}")
                    print(f"    ...还有{len(headers)-5}个表头")
                print()

            print("🔍 完整结果 (JSON格式):")
            print(json.dumps(result, ensure_ascii=False, indent=2))

        else:
            print(f"❌ 解析失败: {result.get('error', '未知错误')}")

    except Exception as e:
        print(f"💥 程序异常: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_trbuff()
