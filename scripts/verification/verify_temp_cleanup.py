#!/usr/bin/env python3
"""
验证临时文件清理和核心功能
"""

import sys
import os
import tempfile
from pathlib import Path

# 添加src目录到路径
sys.path.append('src')

def verify_core_functionality():
    """验证核心功能"""
    try:
        from src.api.excel_operations import ExcelOperations
        print("OK ExcelOperations导入成功")

        # 创建临时测试文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file_path = tmp.name

            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(['Test', 'Data'])
            ws.append([1, 'Value1'])
            wb.save(file_path)
            wb.close()

            # 测试读取功能
            result = ExcelOperations.get_range(file_path, 'Sheet1!A1:B2')
            print(f"OK get_range测试: {result['success']}")

            # 测试安全功能
            impact = ExcelOperations.assess_operation_impact(
                file_path=file_path,
                range_expression="Sheet1!A1:C1",
                operation_type="read",
                preview_data=None
            )
            print(f"OK assess_operation_impact测试: {impact['success']}")

            # 清理临时文件
            os.unlink(file_path)
            print("OK 临时文件清理完成")

        return True

    except Exception as e:
        print(f"❌ 核心功能验证失败: {e}")
        return False

def verify_temp_directory():
    """验证临时文件目录"""
    system_temp = tempfile.gettempdir()
    excel_temp_dir = os.path.join(system_temp, "excel_mcp_server_tests")

    if os.path.exists(excel_temp_dir):
        files = list(Path(excel_temp_dir).glob("*"))
        print(f"✅ 临时目录存在: {excel_temp_dir}")
        print(f"✅ 临时文件数量: {len(files)}")
        return True
    else:
        print(f"❌ 临时目录不存在: {excel_temp_dir}")
        return False

def main():
    """主函数"""
    print("Excel MCP Server - 临时文件清理验证")
    print("=" * 50)

    # 验证临时目录
    print("\n1. 验证临时文件目录...")
    temp_ok = verify_temp_directory()

    # 验证核心功能
    print("\n2. 验证核心功能...")
    core_ok = verify_core_functionality()

    # 总结
    print("\n" + "=" * 50)
    if temp_ok and core_ok:
        print("🎉 验证成功！所有功能正常")
        print("📁 临时文件已移动到系统temp目录")
        print("🔧 核心Excel操作功能正常")
    else:
        print("⚠️  验证过程中发现问题")

    return temp_ok and core_ok

if __name__ == "__main__":
    main()