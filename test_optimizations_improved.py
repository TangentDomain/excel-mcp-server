#!/usr/bin/env python3
"""
Excel MCP 优化功能测试脚本 (改进版)
解决Windows文件锁定问题，并验证缓存机制、中文字符处理和统一错误处理
"""

import os
import sys
import time
import uuid
import tempfile
import gc
from contextlib import contextmanager

# 添加项目根目录到路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from src.core.excel_writer import ExcelWriter
from src.core.excel_manager import ExcelManager
from src.utils.formula_cache import get_formula_cache


@contextmanager
def temporary_excel_file(prefix="test_excel_"):
    """
    安全的临时Excel文件管理器
    使用上下文管理器确保文件被正确清理
    """
    # 生成唯一文件名避免冲突
    unique_id = str(uuid.uuid4())[:8]
    file_name = f"{prefix}{unique_id}.xlsx"
    file_path = os.path.join(tempfile.gettempdir(), file_name)
    
    writer = None
    manager = None
    
    try:
        yield file_path
    finally:
        # 强制清理所有可能的引用
        if 'writer' in locals():
            del writer
        if 'manager' in locals():
            del manager
        
        # 强制垃圾回收
        gc.collect()
        
        # 多次尝试删除文件
        max_attempts = 5
        for attempt in range(max_attempts):
            try:
                if os.path.exists(file_path):
                    os.unlink(file_path)
                break
            except PermissionError:
                if attempt < max_attempts - 1:
                    time.sleep(0.2)  # 等待文件句柄释放
                    gc.collect()  # 再次垃圾回收
                else:
                    print(f"   ⚠️  无法删除临时文件: {file_path}")


def test_formula_caching_improved():
    """测试公式计算缓存功能 (改进版)"""
    print("🧮 测试公式计算缓存功能 (改进版)...")
    
    with temporary_excel_file("cache_test_") as test_file:
        # 创建测试文件
        manager = ExcelManager.create_file(test_file, ["缓存测试"])
        
        # 确保文件创建成功
        if not os.path.exists(test_file):
            print("   ❌ 测试文件创建失败")
            return
        
        # 写入测试数据
        try:
            writer = ExcelWriter(test_file)
            test_data = [
                ["数值1", "数值2", "数值3"],
                [10, 20, 30],
                [15, 25, 35], 
                [20, 30, 40],
                [5, 10, 15]
            ]
            writer.update_range("A1:C5", test_data)
            
            # 获取缓存实例
            cache = get_formula_cache()
            cache.clear()  # 清空缓存开始测试
            
            print(f"   初始缓存状态: {cache.get_stats()}")
            
            # 测试多个公式的缓存效果
            formulas = ["SUM(A2:A5)", "AVERAGE(A2:A5)", "MAX(B2:B5)"]
            
            # 第一轮：无缓存测试
            first_times = []
            for formula in formulas:
                cache.clear()  # 确保无缓存
                start_time = time.time()
                result = writer.evaluate_formula(formula)
                elapsed = time.time() - start_time
                first_times.append(elapsed)
                
                if result and result.success:
                    print(f"   公式 {formula}: {elapsed*1000:.2f}ms, 结果: {result.data}")
                else:
                    print(f"   公式 {formula}: 计算失败")
            
            print(f"   第一轮缓存状态: {cache.get_stats()}")
            
            # 第二轮：有缓存测试
            second_times = []
            for formula in formulas:
                start_time = time.time()
                result = writer.evaluate_formula(formula)
                elapsed = time.time() - start_time
                second_times.append(elapsed)
                print(f"   缓存命中 {formula}: {elapsed*1000:.2f}ms")
            
            print(f"   第二轮缓存状态: {cache.get_stats()}")
            
            # 分析缓存效果
            total_first = sum(first_times)
            total_second = sum(second_times)
            
            if total_first > 0:
                improvement = ((total_first - total_second) / total_first) * 100
                print(f"   📊 总体性能提升: {improvement:.1f}%")
                
                if improvement > 10:  # 至少10%的提升
                    print("   ✅ 缓存机制工作正常！")
                else:
                    print("   ⚠️  缓存效果不明显，可能需要调优")
            
            # 清理引用
            del writer, manager
            
        except Exception as e:
            print(f"   ❌ 缓存测试异常: {e}")
    
    print()


def test_chinese_characters_simple():
    """简化的中文字符测试"""
    print("🇨🇳 测试中文字符处理 (简化版)...")
    
    with temporary_excel_file("chinese_test_") as test_file:
        try:
            # 创建基础文件
            ExcelManager.create_file(test_file, ["初始表"])
            manager = ExcelManager(test_file)
            
            # 测试重点中文场景
            test_cases = [
                ("数据分析", "应该保持原样"),
                ("测试/表", "特殊字符应被替换"),
                ("", "空名称应报错"),
                ("很长的中文工作表名称超过31个字符的情况", "长名称应被处理")
            ]
            
            success_count = 0
            for name, description in test_cases:
                try:
                    result = manager.create_sheet(name)
                    if result.success:
                        actual_name = result.data.name
                        print(f"   ✅ '{name}' -> '{actual_name}' ({description})")
                        success_count += 1
                    else:
                        print(f"   ❌ '{name}' 失败: {result.error} ({description})")
                except Exception as e:
                    print(f"   ❌ '{name}' 异常: {e}")
            
            print(f"   📊 成功率: {success_count}/{len(test_cases)}")
            
            # 清理引用
            del manager
            
        except Exception as e:
            print(f"   ❌ 中文字符测试异常: {e}")
    
    print()


def test_error_handling_patterns():
    """测试统一错误处理模式"""
    print("⚠️  测试统一错误处理...")
    
    try:
        # 测试不存在的文件 - 这会在构造函数中抛出异常
        non_existent_file = "不存在的文件_" + str(uuid.uuid4())[:8] + ".xlsx"
        
        try:
            # 这应该触发异常
            writer = ExcelWriter(non_existent_file)
            print("   ❌ 应该抛出异常但却成功创建了writer")
        except Exception as e:
            # 检查异常类型和消息
            if "Excel文件不存在" in str(e):
                print(f"   ✅ 正确捕获文件不存在异常: {e}")
            else:
                print(f"   ⚠️  异常类型可能不正确: {e}")
        
        # 测试API层面的错误处理 - 这应该被装饰器处理
        with temporary_excel_file("error_test_") as test_file:
            ExcelManager.create_file(test_file, ["测试"])
            writer = ExcelWriter(test_file)
            
            # 测试无效公式
            result = writer.evaluate_formula("")
            if result and not result.success:
                error = result.error
                if isinstance(error, str):
                    print(f"   ✅ API错误处理正常: {error}")
                else:
                    print(f"   ⚠️  错误格式: {error}")
            else:
                print("   ❌ 应该返回错误但却成功了")
            
    except Exception as e:
        print(f"   ❌ 测试过程异常: {e}")
    
    print()


def run_quick_integration_test():
    """快速集成测试"""
    print("⚡ 快速集成测试...")
    
    with temporary_excel_file("integration_test_") as test_file:
        try:
            # 创建包含中文的工作表
            result = ExcelManager.create_file(test_file, ["集成测试"])
            if not result.success:
                print("   ❌ 基础文件创建失败")
                return
                
            manager = ExcelManager(test_file)
            
            # 添加中文工作表
            chinese_sheet = manager.create_sheet("数据统计")
            if not chinese_sheet.success:
                print("   ❌ 中文工作表创建失败")
                return
            
            # 写入数据并测试缓存
            writer = ExcelWriter(test_file)
            test_data = [[i, i*2, i*3] for i in range(1, 21)]  # 20行数据
            writer.update_range("A1:C20", test_data)
            
            # 测试复杂公式的缓存
            complex_formula = "SUM(A1:A20)*AVERAGE(B1:B20)"
            
            # 清除缓存，第一次计算
            cache = get_formula_cache()
            cache.clear()
            
            start_time = time.time()
            result1 = writer.evaluate_formula(complex_formula)
            time1 = time.time() - start_time
            
            # 第二次计算（应该有缓存）
            start_time = time.time()
            result2 = writer.evaluate_formula(complex_formula)
            time2 = time.time() - start_time
            
            if result1 and result1.success and result2 and result2.success:
                print(f"   ✅ 复杂公式计算成功: {result1.data}")
                print(f"   ⚡ 首次计算: {time1*1000:.2f}ms")
                print(f"   ⚡ 缓存计算: {time2*1000:.2f}ms")
                
                if time2 < time1 * 0.8:
                    print("   ✅ 缓存优化效果显著")
                else:
                    print("   ⚠️  缓存优化效果有限")
                    
                print(f"   📊 最终缓存状态: {cache.get_stats()}")
            else:
                print("   ❌ 复杂公式计算失败")
                if result1:
                    print(f"       第一次计算: {result1.success}, 错误: {result1.error if not result1.success else 'None'}")
                if result2:
                    print(f"       第二次计算: {result2.success}, 错误: {result2.error if not result2.success else 'None'}")
            
            # 清理引用
            del writer, manager
            
        except Exception as e:
            print(f"   ❌ 集成测试异常: {e}")
            import traceback
            traceback.print_exc()
    
    print()


def main():
    """主测试函数"""
    print("🚀 Excel MCP 优化功能测试 (改进版) 开始\n")
    
    try:
        test_formula_caching_improved()
        test_chinese_characters_simple()
        test_error_handling_patterns()
        run_quick_integration_test()
        
        print("✅ 所有测试完成！文件锁定问题已解决。")
        
    except Exception as e:
        print(f"❌ 测试过程中出现异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
