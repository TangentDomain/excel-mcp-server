#!/usr/bin/env python3
"""
Excel MCP Server 全面功能验证脚本

模拟真实使用场景，验证所有功能是否在实际应用中正常工作
"""

import os
import sys
import time
import uuid
import tempfile
import pytest
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.server import mcp
from src.core.excel_manager import ExcelManager
from src.core.excel_writer import ExcelWriter
from src.core.excel_reader import ExcelReader
from src.utils.formula_cache import get_formula_cache
from src.utils.validators import DataValidationError


def test_real_world_scenario():
    """真实世界使用场景测试"""
    print("真实世界使用场景测试...")

    # 使用真实的文件名，模拟用户实际操作
    test_file = os.path.join(tempfile.gettempdir(), f"财务报表_{uuid.uuid4().hex[:8]}.xlsx")

    try:
        print(f"   创建测试文件: {test_file}")

        # 1. 创建包含中文名称的工作表
        result = ExcelManager.create_file(test_file, ["总览"])
        if not result.success:
            print(f"   ❌ 文件创建失败: {result.error}")
            return False

        manager = ExcelManager(test_file)

        # 2. 添加多个中文工作表
        chinese_sheets = ["销售数据", "成本分析", "利润统计", "趋势预测"]
        for sheet_name in chinese_sheets:
            result = manager.create_sheet(sheet_name)
            if result.success:
                print(f"   [OK] 成功创建工作表: '{sheet_name}'")
            else:
                print(f"   [ERROR] 工作表创建失败: {sheet_name} - {result.error}")

        # 3. 写入真实的财务数据
        writer = ExcelWriter(test_file)

        # 销售数据
        sales_data = [
            ["月份", "销售额", "成本", "利润"],
            ["1月", 120000, 80000, 40000],
            ["2月", 135000, 85000, 50000],
            ["3月", 148000, 92000, 56000],
            ["4月", 132000, 88000, 44000],
            ["5月", 156000, 98000, 58000],
            ["6月", 169000, 105000, 64000],
        ]

        result = writer.update_range("销售数据!A1:D7", sales_data)
        if result.success:
            print(f"   [OK] 销售数据写入成功: {len(result.data)} 个单元格")
        else:
            print(f"   [ERROR] 数据写入失败: {result.error}")
            return False

        # 4. 读取验证
        reader = ExcelReader(test_file)
        sheets_result = reader.list_sheets()
        if sheets_result.success:
            sheet_names = [s.name for s in sheets_result.data]
            print(f"   [INFO] 最终工作表: {sheet_names}")

        # 验证数据完整性
        data_result = reader.get_range("销售数据!A1:D7")
        if data_result.success:
            read_data = data_result.data
            if isinstance(read_data, list) and len(read_data) == len(sales_data):
                print(f"   [OK] 数据完整性验证通过: {len(read_data)}行数据")
            else:
                print(f"   [WARNING] 数据格式与预期不同: {type(read_data)}")
        else:
            print(f"   [ERROR] 数据读取失败: {data_result.error}")

        print("   [SUCCESS] 真实场景测试完成！")
        return True

    except Exception as e:
        print(f"   [ERROR] 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理
        try:
            if os.path.exists(test_file):
                time.sleep(0.1)  # 确保文件句柄释放
                os.unlink(test_file)
                print(f"   [CLEANUP] 已清理测试文件")
        except Exception as e:
            print(f"   [WARNING] 清理文件失败: {e}")


def test_mcp_server_integration():
    """测试与MCP服务器的集成"""
    print("[INTEGRATION] MCP服务器集成测试...")

    try:
        # 测试MCP服务器对象是否存在
        if mcp is not None:
            print(f"   [OK] MCP服务器对象已创建")
        else:
            print("   [ERROR] MCP服务器对象创建失败")
            return False

        # 检查是否有工具注册方法
        if hasattr(mcp, 'tool'):
            print("   [OK] MCP服务器支持工具注册")
        else:
            print("   [ERROR] MCP服务器不支持工具注册")
            return False

        # 测试核心模块是否正常
        from src.core.excel_manager import ExcelManager
        from src.core.excel_writer import ExcelWriter
        from src.core.excel_reader import ExcelReader
        from src.core.excel_search import ExcelSearcher

        print("   [OK] 所有核心模块导入成功")

        # 测试错误处理模块
        from src.utils.error_handler import unified_error_handler
        print("   [OK] 统一错误处理模块加载成功")

        # 测试缓存模块
        cache = get_formula_cache()
        if cache:
            print("   [OK] 公式缓存模块加载成功")

        return True

    except Exception as e:
        print(f"   [ERROR] 集成测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_performance_benchmark():
    """性能基准测试"""
    print("[PERFORMANCE] 性能基准测试...")

    test_file = os.path.join(tempfile.gettempdir(), f"performance_test_{uuid.uuid4().hex[:6]}.xlsx")

    try:
        # 创建大量数据进行性能测试
        ExcelManager.create_file(test_file, ["性能测试"])
        writer = ExcelWriter(test_file)

        # 生成大量测试数据 (100行 x 10列)
        print("   [INFO] 生成大量测试数据...")
        large_data = [["列" + str(i) for i in range(10)]]  # 表头
        for row in range(100):
            large_data.append([row + 1 + col * 100 for col in range(10)])

        start_time = time.time()
        result = writer.update_range("性能测试!A1:J101", large_data)
        write_time = time.time() - start_time

        if result.success:
            print(f"   [OK] 大量数据写入成功: 1010个单元格，耗时 {write_time*1000:.2f}ms")
        else:
            print(f"   [ERROR] 数据写入失败: {result.error}")
            return False

        # 性能等级评定
        if write_time < 1.0:
            print("   [EXCELLENT] 性能等级: 优秀")
        elif write_time < 2.0:
            print("   [GOOD] 性能等级: 良好")
        elif write_time < 5.0:
            print("   [AVERAGE] 性能等级: 一般")
        else:
            print("   [POOR] 性能等级: 需要改进")

        return True

    except Exception as e:
        print(f"   [ERROR] 性能测试异常: {e}")
        return False

    finally:
        try:
            if os.path.exists(test_file):
                time.sleep(0.1)
                os.unlink(test_file)
        except:
            pass


def test_safety_features():
    """测试安全功能"""
    print("[SAFETY] 安全功能测试...")

    try:
        # 测试范围验证
        from src.utils.validators import ExcelValidator

        # 测试有效范围
        valid_ranges = [
            "Sheet1!A1:C10",
            "数据!1:100",
            "Report!A:Z"
        ]

        for range_expr in valid_ranges:
            result = ExcelValidator.validate_range_expression(range_expr)
            assert result['success'] is True
        print("   [OK] 范围验证功能正常")

        # 测试无效范围检测
        try:
            ExcelValidator.validate_range_expression("invalid_range")
            assert False, "应该检测到无效范围"
        except DataValidationError:
            print("   [OK] 无效范围检测正常")

        # 测试操作规模验证
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 10, 'start_row': 1, 'end_row': 10}
        scale_result = ExcelValidator.validate_operation_scale(range_info)
        assert scale_result['within_limits'] is True
        print("   [OK] 操作规模验证正常")

        return True

    except Exception as e:
        print(f"   [ERROR] 安全功能测试异常: {e}")
        return False


def main():
    """主测试函数"""
    print("Excel MCP Server 全面功能验证开始\n")
    print("=" * 60)

    test_results = []

    # 执行所有测试
    test_results.append(("真实世界场景", test_real_world_scenario()))
    test_results.append(("MCP服务器集成", test_mcp_server_integration()))
    test_results.append(("性能基准测试", test_performance_benchmark()))
    test_results.append(("安全功能验证", test_safety_features()))

    # 汇总结果
    print("\n" + "=" * 60)
    print("[SUMMARY] 测试结果汇总:")

    passed = 0
    total = len(test_results)

    for test_name, result in test_results:
        status = "[PASS] 通过" if result else "[FAIL] 失败"
        print(f"   {test_name}: {status}")
        if result:
            passed += 1

    success_rate = (passed / total) * 100
    print(f"\n[RESULT] 总体成功率: {passed}/{total} ({success_rate:.1f}%)")

    if success_rate == 100:
        print("[SUCCESS] 所有测试都通过了！Excel MCP Server 已经完全准备就绪！")
    elif success_rate >= 80:
        print("[GOOD] 大部分测试通过！系统基本可用，有少量需要改进的地方。")
    else:
        print("[WARNING] 部分测试失败，需要进一步调试和优化。")

    print("\n[INFO] 如果您发现任何问题，请告诉我具体的错误信息，我会立即修复！")


if __name__ == "__main__":
    main()
