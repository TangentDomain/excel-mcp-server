"""
Excel MCP Server - 简化测试运行器

用于验证重构后的代码结构和基本功能
"""

import sys
import os
from pathlib import Path

# 添加src目录到Python路径
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

def test_imports():
    """测试所有模块是否能正确导入"""
    print("🧪 测试模块导入...")

    try:
        # 测试utils模块
        from excel_mcp.utils.validators import ExcelValidator
        from excel_mcp.utils.parsers import RangeParser
        from excel_mcp.utils.exceptions import ExcelException
        print("✅ utils模块导入成功")

        # 测试models模块
        from excel_mcp.models.types import RangeType, OperationResult
        print("✅ models模块导入成功")

        # 测试core模块
        from excel_mcp.core.excel_reader import ExcelReader
        from excel_mcp.core.excel_writer import ExcelWriter
        from excel_mcp.core.excel_manager import ExcelManager
        from excel_mcp.core.excel_search import ExcelSearcher
        print("✅ core模块导入成功")

        return True
    except Exception as e:
        print(f"❌ 导入失败: {e}")
        return False

def test_validator_basic():
    """测试验证器基本功能"""
    print("🧪 测试验证器...")

    try:
        from excel_mcp.utils.validators import ExcelValidator
        from excel_mcp.utils.exceptions import DataValidationError

        # 测试行操作验证
        try:
            ExcelValidator.validate_row_operations(0, 1)
            print("❌ 应该抛出异常但没有")
            return False
        except DataValidationError:
            print("✅ 行操作验证正常工作")

        # 测试列操作验证
        try:
            ExcelValidator.validate_column_operations(1, 101)
            print("❌ 应该抛出异常但没有")
            return False
        except DataValidationError:
            print("✅ 列操作验证正常工作")

        return True
    except Exception as e:
        print(f"❌ 验证器测试失败: {e}")
        return False

def test_parser_basic():
    """测试解析器基本功能"""
    print("🧪 测试解析器...")

    try:
        from excel_mcp.utils.parsers import RangeParser
        from excel_mcp.models.types import RangeType

        # 测试单元格范围解析
        result = RangeParser.parse_range_expression("A1:C10")
        assert result.range_type == RangeType.CELL_RANGE
        assert result.cell_range == "A1:C10"
        print("✅ 单元格范围解析正常")

        # 测试行范围解析
        result = RangeParser.parse_range_expression("1:5")
        assert result.range_type == RangeType.ROW_RANGE
        print("✅ 行范围解析正常")

        return True
    except Exception as e:
        print(f"❌ 解析器测试失败: {e}")
        return False

def test_data_types():
    """测试数据类型"""
    print("🧪 测试数据类型...")

    try:
        from excel_mcp.models.types import OperationResult, CellInfo, RangeType

        # 测试OperationResult
        result = OperationResult(success=True, message="测试")
        assert result.success is True
        assert result.message == "测试"
        print("✅ OperationResult正常")

        # 测试CellInfo
        cell = CellInfo(coordinate="A1", value="test")
        assert cell.coordinate == "A1"
        assert cell.value == "test"
        print("✅ CellInfo正常")

        return True
    except Exception as e:
        print(f"❌ 数据类型测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("🚀 开始Excel MCP服务器重构验证...\n")

    tests = [
        ("模块导入", test_imports),
        ("验证器功能", test_validator_basic),
        ("解析器功能", test_parser_basic),
        ("数据类型", test_data_types),
    ]

    passed = 0
    total = len(tests)

    for test_name, test_func in tests:
        print(f"\n📋 {test_name}:")
        if test_func():
            passed += 1
            print(f"✅ {test_name} 通过")
        else:
            print(f"❌ {test_name} 失败")

    print(f"\n📊 测试结果: {passed}/{total} 通过")

    if passed == total:
        print("🎉 所有基础测试通过！重构成功。")
        return 0
    else:
        print("😞 部分测试失败，需要检查代码。")
        return 1

if __name__ == "__main__":
    exit(main())
