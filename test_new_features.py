#!/usr/bin/env python3
"""
测试新增功能：公式和格式化
"""

import sys
import os
import tempfile

# 添加模块路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from excel_mcp.core.excel_writer import ExcelWriter
from excel_mcp.core.excel_manager import ExcelManager

def test_new_features():
    """测试新增的公式和格式化功能"""
    print("🧪 测试新功能...")

    try:
        # 创建临时文件路径（不预先创建文件）
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"test_excel_mcp_{os.getpid()}.xlsx")

        # 确保文件不存在
        if os.path.exists(temp_file):
            os.unlink(temp_file)

        # 1. 创建Excel文件
        result = ExcelManager.create_file(temp_file, ["测试表"])
        if not result.success:
            print(f"❌ 创建文件失败: {result.error}")
            return False

        # 2. 写入一些数据
        writer = ExcelWriter(temp_file)
        data = [[10, 20], [30, 40]]
        result = writer.update_range("A1:B2", data)
        if not result.success:
            print(f"❌ 写入数据失败: {result.error}")
            return False

        # 3. 测试公式功能
        result = writer.set_formula("C1", "A1+B1")
        if result.success:
            print(f"✅ 公式设置成功: C1 = A1+B1, 计算值: {result.metadata.get('calculated_value')}")
        else:
            print(f"❌ 设置公式失败: {result.error}")
            return False

        # 4. 测试格式化功能
        formatting = {
            'font': {'bold': True, 'size': 14},
            'fill': {'color': 'FFFF00'},  # 黄色背景
            'alignment': {'horizontal': 'center'}
        }
        result = writer.format_cells("A1:C2", formatting)
        if result.success:
            print(f"✅ 格式化成功: 格式化了 {result.metadata.get('formatted_count')} 个单元格")
        else:
            print(f"❌ 格式化失败: {result.error}")
            return False

        # 清理临时文件
        if os.path.exists(temp_file):
            os.unlink(temp_file)

        print("✅ 新功能测试全部通过！")
        return True

    except Exception as e:
        print(f"❌ 测试异常: {e}")
        return False

if __name__ == "__main__":
    print("🎯 Excel MCP新功能测试")
    print("=" * 40)

    success = test_new_features()

    print("=" * 40)
    if success:
        print("🎉 所有新功能测试通过！")
        sys.exit(0)
    else:
        print("💥 部分测试失败！")
        sys.exit(1)
