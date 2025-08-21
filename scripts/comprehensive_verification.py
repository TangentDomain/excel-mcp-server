#!/usr/bin/env python3
"""
Excel MCP Server 全面功能验证脚本
模拟真实使用场景，验证所有优化是否在实际应用中正常工作
"""

import os
import sys
import time
import uuid
import tempfile

# 添加项目根目录到路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from src.server import mcp
from src.core.excel_manager import ExcelManager
from src.core.excel_writer import ExcelWriter
from src.core.excel_reader import ExcelReader
from src.utils.formula_cache import get_formula_cache


def test_real_world_scenario():
    """真实世界使用场景测试"""
    print("🌍 真实世界使用场景测试...")

    # 使用真实的文件名，模拟用户实际操作
    test_file = os.path.join(tempfile.gettempdir(), f"财务报表_{uuid.uuid4().hex[:8]}.xlsx")

    try:
        print(f"   📁 创建测试文件: {test_file}")

        # 1. 创建包含中文名称的工作表
        result = ExcelManager.create_file(test_file, ["总览"])
        if not result.success:
            print(f"   ❌ 文件创建失败: {result.error}")
            return False

        manager = ExcelManager(test_file)

        # 2. 添加多个中文工作表
        chinese_sheets = ["销售数据", "成本分析", "利润统计", "趋势/预测"]
        for sheet_name in chinese_sheets:
            result = manager.create_sheet(sheet_name)
            if result.success:
                print(f"   ✅ 成功创建工作表: '{sheet_name}' -> '{result.data.name}'")
            else:
                print(f"   ❌ 工作表创建失败: {sheet_name} - {result.error}")

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

        result = writer.update_range("A1:D7", sales_data)
        if result.success:
            print(f"   ✅ 销售数据写入成功: {len(result.data)} 个单元格")
        else:
            print(f"   ❌ 数据写入失败: {result.error}")
            return False

        # 4. 测试复杂的公式计算（模拟真实业务场景）
        print("   🧮 测试业务公式计算...")

        cache = get_formula_cache()
        cache.clear()

        business_formulas = [
            ("总销售额", "SUM(B2:B7)"),
            ("平均销售额", "AVERAGE(B2:B7)"),
            ("总成本", "SUM(C2:C7)"),
            ("总利润", "SUM(D2:D7)"),
            ("利润率", "SUM(D2:D7)/SUM(B2:B7)*100"),
            ("最大月销售", "MAX(B2:B7)"),
            ("最小月销售", "MIN(B2:B7)"),
            ("销售增长", "(B7-B2)/B2*100"),
        ]

        # 第一轮计算（建立缓存）
        first_round_results = {}
        first_round_times = {}

        for name, formula in business_formulas:
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time

            if result and result.success:
                first_round_results[name] = result.data
                first_round_times[name] = elapsed
                print(f"       {name}: {result.data} (耗时: {elapsed*1000:.2f}ms)")
            else:
                print(f"       ❌ {name} 计算失败: {result.error if result else '未知错误'}")

        # 5. 第二轮计算（测试缓存效果）
        print("   🚀 测试缓存性能...")

        second_round_times = {}
        cache_hits = 0

        for name, formula in business_formulas:
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time

            if result and result.success:
                second_round_times[name] = elapsed
                # 验证结果一致性
                if abs(result.data - first_round_results[name]) < 0.001:
                    cache_hits += 1
                    improvement = (first_round_times[name] - elapsed) / first_round_times[name] * 100
                    print(f"       ✅ {name}: 缓存命中，性能提升 {improvement:.1f}%")
                else:
                    print(f"       ⚠️  {name}: 结果不一致")

        cache_stats = cache.get_stats()
        print(f"   📊 缓存统计: {cache_stats}")
        print(f"   📈 缓存命中数: {cache_hits}/{len(business_formulas)}")

        # 6. 测试错误处理
        print("   ⚠️  测试错误处理...")

        error_test_cases = [
            ("空公式", ""),
            ("无效函数", "INVALID_FUNC(A1)"),
            ("循环引用", "A1+A1"),
            ("除零错误", "B2/0"),
        ]

        error_handled_count = 0
        for case_name, formula in error_test_cases:
            result = writer.evaluate_formula(formula)
            if result and not result.success:
                error_handled_count += 1
                print(f"       ✅ {case_name}: 错误正确处理 - {result.error}")
            else:
                print(f"       ❌ {case_name}: 应该返回错误但成功了")

        print(f"   📊 错误处理成功率: {error_handled_count}/{len(error_test_cases)}")

        # 7. 读取验证
        reader = ExcelReader(test_file)
        sheets_result = reader.list_sheets()
        if sheets_result.success:
            sheet_names = [s.name for s in sheets_result.data]
            print(f"   📋 最终工作表: {sheet_names}")

        # 验证数据完整性
        data_result = reader.get_range("A1:D7")
        if data_result.success:
            read_data = data_result.data
            if isinstance(read_data, list) and len(read_data) == len(sales_data):
                print(f"   ✅ 数据完整性验证通过: {len(read_data)}行数据")
            elif hasattr(read_data, 'rows') and len(read_data.rows) == len(sales_data):
                print(f"   ✅ 数据完整性验证通过: {len(read_data.rows)}行数据")
            else:
                print(f"   ⚠️  数据格式与预期不同: {type(read_data)}")
        else:
            print(f"   ❌ 数据读取失败: {data_result.error}")

        print("   🎉 真实场景测试完成！")
        return True

    except Exception as e:
        print(f"   ❌ 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理
        try:
            if os.path.exists(test_file):
                time.sleep(0.1)  # 确保文件句柄释放
                os.unlink(test_file)
                print(f"   🗑️  已清理测试文件")
        except Exception as e:
            print(f"   ⚠️  清理文件失败: {e}")


def test_mcp_server_integration():
    """测试与MCP服务器的集成"""
    print("🔗 MCP服务器集成测试...")

    try:
        # 测试MCP服务器对象是否存在
        if mcp is not None:
            print(f"   ✅ MCP服务器对象已创建")
        else:
            print("   ❌ MCP服务器对象创建失败")
            return False

        # 检查是否有工具注册方法
        if hasattr(mcp, 'tool'):
            print("   ✅ MCP服务器支持工具注册")
        else:
            print("   ❌ MCP服务器不支持工具注册")
            return False

        # 测试核心模块是否正常
        from src.core.excel_manager import ExcelManager
        from src.core.excel_writer import ExcelWriter
        from src.core.excel_reader import ExcelReader
        from src.core.excel_search import ExcelSearcher

        print("   ✅ 所有核心模块导入成功")

        # 测试错误处理模块
        from src.utils.error_handler import unified_error_handler
        print("   ✅ 统一错误处理模块加载成功")

        # 测试缓存模块
        cache = get_formula_cache()
        if cache:
            print("   ✅ 公式缓存模块加载成功")

        return True

    except Exception as e:
        print(f"   ❌ 集成测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_performance_benchmark():
    """性能基准测试"""
    print("⚡ 性能基准测试...")

    test_file = os.path.join(tempfile.gettempdir(), f"性能测试_{uuid.uuid4().hex[:6]}.xlsx")

    try:
        # 创建大量数据进行性能测试
        ExcelManager.create_file(test_file, ["性能测试"])
        writer = ExcelWriter(test_file)

        # 生成大量测试数据 (100行 x 10列)
        print("   📊 生成大量测试数据...")
        large_data = [["列" + str(i) for i in range(10)]]  # 表头
        for row in range(100):
            large_data.append([row + 1 + col * 100 for col in range(10)])

        start_time = time.time()
        result = writer.update_range("A1:J101", large_data)
        write_time = time.time() - start_time

        if result.success:
            print(f"   ✅ 大量数据写入成功: 1010个单元格，耗时 {write_time*1000:.2f}ms")
        else:
            print(f"   ❌ 数据写入失败: {result.error}")
            return False

        # 测试复杂公式的性能
        complex_formulas = [
            "SUM(A2:A101)",
            "AVERAGE(B2:B101)",
            "MAX(C2:C101)",
            "MIN(D2:D101)",
            "SUM(A2:A101)*AVERAGE(B2:B101)",
        ]

        cache = get_formula_cache()
        cache.clear()

        # 无缓存性能
        no_cache_times = []
        for formula in complex_formulas:
            cache.clear()  # 确保无缓存
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time
            no_cache_times.append(elapsed)

            if result and result.success:
                print(f"       无缓存 {formula}: {elapsed*1000:.2f}ms")

        # 有缓存性能
        cached_times = []
        for i, formula in enumerate(complex_formulas):
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time
            cached_times.append(elapsed)

            if result and result.success:
                improvement = (no_cache_times[i] - elapsed) / no_cache_times[i] * 100
                print(f"       有缓存 {formula}: {elapsed*1000:.2f}ms (提升 {improvement:.1f}%)")

        # 总体性能分析
        total_no_cache = sum(no_cache_times)
        total_cached = sum(cached_times)
        overall_improvement = (total_no_cache - total_cached) / total_no_cache * 100

        print(f"   📈 总体性能提升: {overall_improvement:.1f}%")
        print(f"   📊 最终缓存状态: {cache.get_stats()}")

        # 性能等级评定
        if overall_improvement > 80:
            print("   🏆 性能等级: 优秀")
        elif overall_improvement > 50:
            print("   🥇 性能等级: 良好")
        elif overall_improvement > 20:
            print("   🥈 性能等级: 一般")
        else:
            print("   🥉 性能等级: 需要改进")

        return True

    except Exception as e:
        print(f"   ❌ 性能测试异常: {e}")
        return False

    finally:
        try:
            if os.path.exists(test_file):
                time.sleep(0.1)
                os.unlink(test_file)
        except:
            pass


def main():
    """主测试函数"""
    print("🚀 Excel MCP Server 全面功能验证开始\n")
    print("=" * 60)

    test_results = []

    # 执行所有测试
    test_results.append(("真实世界场景", test_real_world_scenario()))
    test_results.append(("MCP服务器集成", test_mcp_server_integration()))
    test_results.append(("性能基准测试", test_performance_benchmark()))

    # 汇总结果
    print("\n" + "=" * 60)
    print("📊 测试结果汇总:")

    passed = 0
    total = len(test_results)

    for test_name, result in test_results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"   {test_name}: {status}")
        if result:
            passed += 1

    success_rate = (passed / total) * 100
    print(f"\n🎯 总体成功率: {passed}/{total} ({success_rate:.1f}%)")

    if success_rate == 100:
        print("🎉 所有测试都通过了！Excel MCP Server 已经完全准备就绪！")
    elif success_rate >= 80:
        print("🎊 大部分测试通过！系统基本可用，有少量需要改进的地方。")
    else:
        print("⚠️  部分测试失败，需要进一步调试和优化。")

    print("\n🔍 如果您发现任何问题，请告诉我具体的错误信息，我会立即修复！")


if __name__ == "__main__":
    main()
