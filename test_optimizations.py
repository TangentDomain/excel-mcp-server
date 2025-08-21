#!/usr/bin/env python3
"""
Excel MCP 优化功能测试脚本
测试缓存机制、中文字符处理和统一错误处理
"""

import os
import sys
import time
import tempfile

# 添加项目根目录到路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from src.core.excel_writer import ExcelWriter
from src.core.excel_manager import ExcelManager
from src.utils.formula_cache import get_formula_cache


def safe_cleanup(file_path):
    """安全清理文件，避免文件占用问题"""
    import gc
    import time
    
    try:
        # 强制垃圾回收
        gc.collect()
        time.sleep(0.1)  # 短暂等待
        
        if os.path.exists(file_path):
            os.unlink(file_path)
    except Exception as e:
        print(f"   ⚠️  清理文件时出现问题: {e}")
def test_formula_caching():
    """测试公式计算缓存功能"""
    print("🧮 测试公式计算缓存功能...")
    
    # 创建测试文件
    test_file = os.path.join(tempfile.gettempdir(), "test_cache.xlsx")
    manager = ExcelManager.create_file(test_file, ["测试数据"])
    
    # 写入测试数据
    writer = ExcelWriter(test_file)
    test_data = [
        ["数值1", "数值2", "数值3"],
        [10, 20, 30],
        [15, 25, 35], 
        [20, 30, 40]
    ]
    writer.update_range("A1:C4", test_data)
    
    cache = get_formula_cache()
    cache.clear()  # 清空缓存开始测试
    
    print(f"   初始缓存状态: {cache.get_stats()}")
    
    # 第一次计算（应该未命中缓存）
    start_time = time.time()
    result1 = writer.evaluate_formula("SUM(A2:A4)")
    time1 = time.time() - start_time
    
    print(f"   第一次计算耗时: {time1*1000:.2f}ms")
    print(f"   计算结果: {result1.data if result1 and result1.success else '失败'}")
    print(f"   缓存状态: {cache.get_stats()}")
    
    # 第二次计算相同公式（应该命中缓存）
    start_time = time.time()
    result2 = writer.evaluate_formula("SUM(A2:A4)")
    time2 = time.time() - start_time
    
    print(f"   第二次计算耗时: {time2*1000:.2f}ms")
    print(f"   计算结果: {result2.data if result2 and result2.success else '失败'}")
    print(f"   缓存状态: {cache.get_stats()}")
    
    # 验证缓存效果
    if result1 and result1.success and result2 and result2.success:
        if time2 < time1 * 0.8:  # 缓存命中应该快至少20%
            print("   ✅ 缓存优化生效！")
        else:
            print("   ❌ 缓存优化可能未生效")
    else:
        print("   ⚠️  公式计算出现问题，无法测试缓存效果")
    
    # 清理
    del writer, manager
    safe_cleanup(test_file)
    print()


def test_chinese_sheet_names():
    """测试中文工作表名称处理"""
    print("🇨🇳 测试中文工作表名称处理...")
    
    test_file = os.path.join(tempfile.gettempdir(), "test_chinese.xlsx")
    
    # 创建基础文件
    ExcelManager.create_file(test_file, ["初始表"])
    manager = ExcelManager(test_file)
    
    # 测试各种中文工作表名称
    test_names = [
        "数据分析",           # 普通中文
        "销售报表2023",       # 中英文混合
        "测试/数据",         # 包含特殊字符
        "很长的中文工作表名称超过三十一个字符的情况测试", # 超长名称
        "   空格测试   ",     # 包含空格
        "",                 # 空名称
        "Sheet*Test",       # 包含无效字符
    ]
    
    results = []
    for name in test_names:
        try:
            result = manager.create_sheet(name)
            if result.success:
                actual_name = result.data.name
                results.append(f"   ✅ '{name}' -> '{actual_name}'")
            else:
                results.append(f"   ❌ '{name}' 失败: {result.error}")
        except Exception as e:
            results.append(f"   ❌ '{name}' 异常: {e}")
    
    for result in results:
        print(result)
    
    # 验证工作表列表
    reader = ExcelWriter(test_file)
    from src.core.excel_reader import ExcelReader
    sheet_reader = ExcelReader(test_file)
    sheets_result = sheet_reader.list_sheets()
    
    if sheets_result.success:
        print(f"   📋 最终工作表列表: {[s.name for s in sheets_result.data]}")
    
    # 清理
    del reader, sheet_reader, manager
    safe_cleanup(test_file)
    print()


def test_unified_error_handling():
    """测试统一错误处理"""
    print("⚠️  测试统一错误处理...")
    
    # 测试文件不存在的情况
    try:
        writer = ExcelWriter("不存在的文件.xlsx")
        result = writer.evaluate_formula("SUM(A1:A10)")
        
        if not result.success:
            error = result.error
            if isinstance(error, dict) and 'code' in error:
                print(f"   ✅ 统一错误格式: {error['code']} - {error['message']}")
            else:
                print(f"   ❌ 错误格式不统一: {error}")
        else:
            print("   ❌ 应该返回错误但却成功了")
            
    except Exception as e:
        print(f"   ❌ 异常未被正确处理: {e}")
    
    print()


def test_performance_comparison():
    """性能对比测试"""
    print("🏃 性能对比测试...")
    
    test_file = os.path.join(tempfile.gettempdir(), "test_performance.xlsx")
    
    # 创建大一些的测试数据
    ExcelManager.create_file(test_file, ["性能测试"])
    writer = ExcelWriter(test_file)
    
    # 生成100行测试数据
    large_data = [["数值"] + [f"列{i}" for i in range(1, 11)]]
    for i in range(100):
        row = [i + 1] + [f"数据{i}_{j}" for j in range(10)]
        large_data.append(row)
    
    writer.update_range("A1:K101", large_data)
    
    cache = get_formula_cache()
    
    # 测试多个复杂公式的缓存效果
    formulas = [
        "SUM(A2:A101)",
        "AVERAGE(A2:A101)", 
        "MAX(A2:A101)",
        "MIN(A2:A101)",
        "COUNT(A2:A101)"
    ]
    
    print("   第一轮计算（无缓存）:")
    first_round_times = []
    for formula in formulas:
        cache.clear()  # 清除缓存确保未命中
        start_time = time.time()
        result = writer.evaluate_formula(formula)
        elapsed = time.time() - start_time
        first_round_times.append(elapsed)
        print(f"     {formula}: {elapsed*1000:.2f}ms")
    
    print("   第二轮计算（有缓存）:")
    second_round_times = []
    for formula in formulas:
        start_time = time.time()
        result = writer.evaluate_formula(formula)
        elapsed = time.time() - start_time
        second_round_times.append(elapsed)
        print(f"     {formula}: {elapsed*1000:.2f}ms")
    
    # 计算总体改善
    total_first = sum(first_round_times)
    total_second = sum(second_round_times)
    improvement = ((total_first - total_second) / total_first) * 100
    
    print(f"   📊 总体性能提升: {improvement:.1f}%")
    print(f"   📊 缓存统计: {cache.get_stats()}")
    
    # 清理
    del writer
    safe_cleanup(test_file)
    print()


def main():
    """主测试函数"""
    print("🚀 Excel MCP 优化功能测试开始\n")
    
    try:
        test_formula_caching()
        test_chinese_sheet_names()
        test_unified_error_handling()
        test_performance_comparison()
        
        print("✅ 所有测试完成！")
        
    except Exception as e:
        print(f"❌ 测试过程中出现异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
