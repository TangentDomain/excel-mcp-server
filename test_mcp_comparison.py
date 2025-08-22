#!/usr/bin/env python3
"""
使用MCP服务器接口测试详细比较功能
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_mcp_excel_compare():
    """使用MCP服务器接口测试比较功能"""
    print("🧪 使用MCP服务器接口测试比较功能...")

    # 导入MCP服务器
    from src.server import FastMCPExcelServer
    from mcp.types import GetPromptRequest

    server = FastMCPExcelServer()

    # 测试文件路径
    file1 = r"D:\tr\svn\trunk\配置表\测试配置\微小\TrSkill.xlsx"
    file2 = r"D:\tr\svn\trunk\配置表\战斗环境配置\TrSkill.xlsx"

    try:
        print(f"📂 比较文件:")
        print(f"  - 文件1: {file1}")
        print(f"  - 文件2: {file2}")
        print()

        # 调用比较工具
        from mcp.types import CallToolRequest
        import asyncio

        async def run_comparison():
            request = CallToolRequest(
                method="call_tool",
                params={
                    "name": "excel_compare_files",
                    "arguments": {
                        "file1_path": file1,
                        "file2_path": file2,
                        "structured_comparison": True,
                        "game_friendly_format": True,
                        "focus_on_id_changes": True,
                        "show_numeric_changes": True
                    }
                }
            )

            result = await server.call_tool(request)
            return result

        # 运行异步比较
        result = asyncio.run(run_comparison())

        print(f"📋 比较结果类型: {type(result)}")
        print(f"📋 比较结果: {result}")

        if hasattr(result, 'content'):
            for content in result.content:
                if hasattr(content, 'text'):
                    import json
                    try:
                        data = json.loads(content.text)
                        print(f"✅ 比较成功!")
                        print(f"📊 发现差异: {data.get('total_differences', 0)}")

                        # 检查详细差异
                        sheet_comparisons = data.get('sheet_comparisons', [])
                        if sheet_comparisons:
                            for sheet_comp in sheet_comparisons:
                                if sheet_comp.get('differences'):
                                    sheet_name = sheet_comp.get('sheet_name', 'Unknown')
                                    differences = sheet_comp.get('differences', [])
                                    print(f"\n📋 工作表 {sheet_name}: {len(differences)} 个差异")

                                    # 检查前几个差异的详细字段变化
                                    for i, diff in enumerate(differences[:3]):
                                        if isinstance(diff, dict) and diff.get('detailed_field_differences'):
                                            print(f"  🔍 ID {diff.get('row_id', 'N/A')} 详细变化:")
                                            for field_diff in diff['detailed_field_differences'][:3]:
                                                print(f"    - {field_diff.get('field_name', 'N/A')}: {field_diff.get('old_value', 'N/A')} → {field_diff.get('new_value', 'N/A')}")
                                    break
                        return True

                    except json.JSONDecodeError:
                        print(f"📋 原始结果: {content.text[:500]}...")
                        return True
        else:
            print(f"⚠️ 结果没有内容字段")
            return False

    except Exception as e:
        print(f"💥 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 MCP Excel详细比较功能测试")
    print("=" * 60)

    success = test_mcp_excel_compare()

    print("\n" + "=" * 60)
    if success:
        print("🎉 MCP接口测试完成!")
    else:
        print("❌ MCP接口测试失败")
    print("=" * 60)
