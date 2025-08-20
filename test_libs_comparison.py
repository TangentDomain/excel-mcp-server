#!/usr/bin/env python3
"""
对比xlwings vs numpy+scipy vs xlcalculator的高级统计支持
"""

import sys
from pathlib import Path
import tempfile
import os
import time

# 添加src路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_xlwings_stats():
    """测试xlwings的统计函数支持"""
    
    print("🚀 测试xlwings统计支持...")
    
    try:
        import xlwings as xw
        print("✅ xlwings导入成功")
        
        # 检查Excel/LibreOffice是否可用
        try:
            # 尝试创建Excel应用实例
            app = xw.App(visible=False, add_book=False)
            print("✅ Excel应用可用")
            
            # 创建新工作簿
            wb = app.books.add()
            ws = wb.sheets[0]
            
            # 添加测试数据
            test_data = [10, 20, 30, 40, 50, 15, 25, 35, 45, 55]
            ws.range('A1').options(transpose=True).value = test_data
            
            print(f"📝 测试数据: {test_data}")
            
            # 测试统计函数
            stats_formulas = {
                "基础统计": {
                    "SUM": "=SUM(A1:A10)",
                    "AVERAGE": "=AVERAGE(A1:A10)", 
                    "COUNT": "=COUNT(A1:A10)",
                    "MIN": "=MIN(A1:A10)",
                    "MAX": "=MAX(A1:A10)",
                },
                "高级统计": {
                    "MEDIAN": "=MEDIAN(A1:A10)",
                    "STDEV": "=STDEV.S(A1:A10)",  # 样本标准差
                    "VAR": "=VAR.S(A1:A10)",      # 样本方差
                    "PERCENTILE": "=PERCENTILE(A1:A10,0.9)",
                    "QUARTILE": "=QUARTILE(A1:A10,1)",
                },
                "条件统计": {
                    "COUNTIF": "=COUNTIF(A1:A10,\">30\")",
                    "SUMIF": "=SUMIF(A1:A10,\">30\")",
                    "AVERAGEIF": "=AVERAGEIF(A1:A10,\">30\")",
                },
                "特殊函数": {
                    "MODE": "=MODE.SNGL(A1:A10)",     # 众数
                    "SKEW": "=SKEW(A1:A10)",          # 偏度
                    "KURT": "=KURT(A1:A10)",          # 峰度
                    "GEOMEAN": "=GEOMEAN(A1:A10)",    # 几何平均数
                    "HARMEAN": "=HARMEAN(A1:A10)",    # 调和平均数
                }
            }
            
            total_success = 0
            total_tests = 0
            results = {}
            
            start_time = time.time()
            
            for category, formulas in stats_formulas.items():
                print(f"\n📊 {category}:")
                category_success = 0
                results[category] = {}
                
                for func_name, formula in formulas.items():
                    total_tests += 1
                    try:
                        # 在Excel中计算公式
                        cell_result = ws.range('B1').formula = formula
                        result_value = ws.range('B1').value
                        
                        print(f"   ✅ {func_name}: {formula} = {result_value}")
                        results[category][func_name] = result_value
                        total_success += 1
                        category_success += 1
                        
                    except Exception as e:
                        print(f"   ❌ {func_name}: {formula} - 错误: {e}")
                        results[category][func_name] = None
                
                print(f"   📈 成功率: {category_success}/{len(formulas)}")
            
            execution_time = (time.time() - start_time) * 1000
            success_rate = total_success / total_tests
            
            print(f"\n🎯 xlwings总体结果:")
            print(f"   成功率: {total_success}/{total_tests} ({success_rate*100:.1f}%)")
            print(f"   执行时间: {execution_time:.1f}ms")
            
            # 关闭工作簿和应用
            wb.close()
            app.quit()
            
            return {
                "success_rate": success_rate,
                "execution_time": execution_time,
                "results": results,
                "available": True
            }
            
        except Exception as e:
            print(f"❌ Excel应用不可用: {e}")
            return {"available": False, "error": str(e)}
            
    except ImportError as e:
        print(f"❌ xlwings导入失败: {e}")
        return {"available": False, "error": str(e)}

def test_numpy_scipy_stats():
    """测试numpy+scipy的统计函数支持"""
    
    print("\n🧮 测试numpy+scipy统计支持...")
    
    try:
        import numpy as np
        from scipy import stats
        print("✅ numpy+scipy可用")
        
        # 测试数据
        data = np.array([10, 20, 30, 40, 50, 15, 25, 35, 45, 55])
        
        start_time = time.time()
        
        # 实现各种统计函数
        numpy_results = {
            "基础统计": {
                "SUM": float(np.sum(data)),
                "AVERAGE": float(np.mean(data)),
                "COUNT": len(data),
                "MIN": float(np.min(data)),
                "MAX": float(np.max(data)),
            },
            "高级统计": {
                "MEDIAN": float(np.median(data)),
                "STDEV": float(np.std(data, ddof=1)),      # 样本标准差
                "VAR": float(np.var(data, ddof=1)),        # 样本方差
                "PERCENTILE": float(np.percentile(data, 90)),
                "QUARTILE": float(np.percentile(data, 25)),
            },
            "条件统计": {
                "COUNTIF": int(np.sum(data > 30)),
                "SUMIF": float(np.sum(data[data > 30])),
                "AVERAGEIF": float(np.mean(data[data > 30])),
            },
            "特殊函数": {
                "MODE": float(stats.mode(data, keepdims=True)[0][0]),  # 众数
                "SKEW": float(stats.skew(data)),          # 偏度
                "KURT": float(stats.kurtosis(data)),      # 峰度
                "GEOMEAN": float(stats.gmean(data)),      # 几何平均数
                "HARMEAN": float(stats.hmean(data)),      # 调和平均数
            }
        }
        
        execution_time = (time.time() - start_time) * 1000
        
        # 统计成功率
        total_success = 0
        total_tests = 0
        
        for category, functions in numpy_results.items():
            print(f"\n📊 {category}:")
            category_success = 0
            
            for func_name, result in functions.items():
                total_tests += 1
                if result is not None and not (isinstance(result, float) and np.isnan(result)):
                    print(f"   ✅ {func_name}: {result}")
                    total_success += 1
                    category_success += 1
                else:
                    print(f"   ❌ {func_name}: 计算失败")
            
            print(f"   📈 成功率: {category_success}/{len(functions)}")
        
        success_rate = total_success / total_tests
        print(f"\n🎯 numpy+scipy总体结果:")
        print(f"   成功率: {total_success}/{total_tests} ({success_rate*100:.1f}%)")
        print(f"   执行时间: {execution_time:.1f}ms")
        
        return {
            "success_rate": success_rate,
            "execution_time": execution_time,
            "results": numpy_results,
            "available": True
        }
        
    except Exception as e:
        print(f"❌ numpy+scipy测试失败: {e}")
        return {"available": False, "error": str(e)}

def compare_libraries():
    """对比所有库的表现"""
    
    print("=" * 60)
    print("📊 高级统计函数库对比分析")
    print("=" * 60)
    
    # 测试xlwings
    xlwings_result = test_xlwings_stats()
    
    # 测试numpy+scipy
    numpy_result = test_numpy_scipy_stats()
    
    # 生成对比报告
    print("\n" + "=" * 60)
    print("📋 综合对比报告")
    print("=" * 60)
    
    libraries = [
        ("xlwings", xlwings_result),
        ("numpy+scipy", numpy_result)
    ]
    
    print(f"{'库名':<15} {'可用性':<8} {'成功率':<10} {'执行时间':<12} {'优势'}")
    print("-" * 60)
    
    for lib_name, result in libraries:
        if result.get("available"):
            success_rate = f"{result['success_rate']*100:.1f}%"
            exec_time = f"{result['execution_time']:.1f}ms"
            
            if lib_name == "xlwings":
                advantage = "100% Excel兼容"
            elif lib_name == "numpy+scipy":
                advantage = "高性能，无依赖"
            else:
                advantage = "-"
                
            print(f"{lib_name:<15} {'✅':<8} {success_rate:<10} {exec_time:<12} {advantage}")
        else:
            print(f"{lib_name:<15} {'❌':<8} {'-':<10} {'-':<12} {result.get('error', '')[:20]}")
    
    # 生成建议
    print(f"\n💡 建议:")
    
    if xlwings_result.get("available") and xlwings_result.get("success_rate", 0) > 0.9:
        print("   🚀 xlwings: 如果系统有Excel/LibreOffice，这是最完整的解决方案")
        print("   📊 支持所有Excel函数，包括最新的统计函数")
        print("   ⚠️  缺点：需要安装Office软件，跨平台兼容性问题")
    
    if numpy_result.get("available") and numpy_result.get("success_rate", 0) > 0.9:
        print("   🧮 numpy+scipy: 性能最佳，跨平台兼容")  
        print("   📈 科学计算标准，统计算法更精确")
        print("   ✅ 优点：无额外依赖，轻量级")
        
    print(f"\n🎯 最终推荐:")
    
    if (xlwings_result.get("available") and 
        numpy_result.get("available") and 
        xlwings_result.get("execution_time", 1000) > numpy_result.get("execution_time", 0)):
        print("   📊 推荐numpy+scipy方案：性能更好，依赖更少")
    elif xlwings_result.get("available"):
        print("   🚀 推荐xlwings方案：Excel兼容性最佳")
    else:
        print("   🧮 推荐numpy+scipy方案：唯一可用的完整解决方案")

if __name__ == "__main__":
    compare_libraries()
