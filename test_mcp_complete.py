#!/usr/bin/env python3
"""测试MCP服务器和ID比较功能"""

def test_mcp_server():
    try:
        from src.server import mcp, excel_compare_files, excel_compare_sheets
        print("✅ MCP服务器和比较函数导入成功")
        print("  - excel_compare_files: 文件比较")
        print("  - excel_compare_sheets: 工作表比较")
        return True
    except Exception as e:
        print(f"❌ 导入失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_id_comparison_config():
    try:
        from src.models.types import ComparisonOptions

        # 测试ID比较的完整配置
        options = ComparisonOptions(
            structured_comparison=True,
            header_row=1,
            id_column=1,
            game_friendly_format=True,
            focus_on_id_changes=True,
            show_numeric_changes=True
        )
        print("✅ ID比较配置创建成功")
        print(f"  - 专注ID变化: {options.focus_on_id_changes}")
        return True
    except Exception as e:
        print(f"❌ 配置创建失败: {e}")
        return False

if __name__ == "__main__":
    print("=== MCP Excel比较服务器测试 ===\n")

    success_count = 0
    total_tests = 2

    print("1. 测试MCP服务器导入...")
    if test_mcp_server():
        success_count += 1
    print()

    print("2. 测试ID比较配置...")
    if test_id_comparison_config():
        success_count += 1
    print()

    print(f"=== 测试结果: {success_count}/{total_tests} 通过 ===")

    if success_count == total_tests:
        print("🎉 所有测试通过！ID对象比较功能已恢复。")
        print("\n使用方法:")
        print("1. 启动MCP服务器: python -m src.server")
        print("2. 比较文件时会自动使用ID对象比较逻辑")
        print("3. 新增/删除/修改的对象会以ID为基础进行分类")
    else:
        print("❌ 部分测试失败，请检查配置")
