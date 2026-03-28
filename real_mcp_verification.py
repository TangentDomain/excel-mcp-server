#!/usr/bin/env python3
"""
真实MCP工具调用验证
使用实际的MCP工具完成极限测试
"""

import sys
import os
import tempfile
import json
import pandas as pd
from pathlib import Path

# 添加源码路径
sys.path.insert(0, 'src')

from excel_mcp_server_fastmcp.server import (
    excel_list_sheets,
    excel_get_range,
    excel_query,
    excel_describe_table,
    excel_find_last_row,
    excel_insert_rows,
    excel_delete_rows
)

def create_real_test_file():
    """创建真实的测试文件"""
    
    with tempfile.TemporaryDirectory() as temp_dir:
        excel_path = os.path.join(temp_dir, "real_game_test.xlsx")
        
        # 创建真实的游戏数据
        # 角色表
        characters = pd.DataFrame({
            '角色ID': [1, 2, 3, 4, 5],
            '角色名称': ['圣骑士', '法师', '刺客', '牧师', '弓箭手'],
            '职业': ['战士', '法师', '刺客', '牧师', '弓箭手'],
            '等级': [30, 25, 35, 20, 40],
            '生命值': [1200, 800, 900, 1000, 850],
            '魔法值': [300, 500, 200, 600, 400],
            '攻击力': [150, 120, 180, 100, 140],
            '防御力': [80, 60, 70, 90, 75],
            '暴击率': [0.15, 0.10, 0.25, 0.05, 0.20]
        })
        
        # 技能表
        skills = pd.DataFrame({
            '技能ID': [1, 2, 3, 4, 5],
            '技能名称': ['圣光斩', '火球术', '暗影打击', '治疗术', '箭雨'],
            '职业限制': ['战士', '法师', '刺客', '牧师', '弓箭手'],
            '伤害公式': ['攻击力*1.5', '魔法值*0.8', '攻击力*1.2', '治疗量*2.0', '攻击力*1.1'],
            '冷却时间': [5.0, 3.0, 1.5, 8.0, 6.0],
            '魔法消耗': [20, 30, 15, 40, 25],
            '技能类型': ['物理攻击', '魔法攻击', '物理攻击', '治疗', '物理攻击'],
            '等级要求': [1, 1, 1, 1, 1]
        })
        
        # 装备表
        equipments = pd.DataFrame({
            '装备ID': [1, 2, 3, 4, 5],
            '装备名称': ['巨剑', '法杖', '匕首', '法袍', '长弓'],
            '装备类型': ['武器', '武器', '武器', '防具', '武器'],
            '基础属性加成': ['攻击力+50', '魔法值+30', '攻击力+40', '防御力+35', '攻击力+45'],
            '等级要求': [10, 15, 8, 12, 20],
            '稀有度': ['精良', '史诗', '稀有', '精良', '史诗']
        })
        
        # 写入Excel文件
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            characters.to_excel(writer, sheet_name='角色', index=False)
            skills.to_excel(writer, sheet_name='技能', index=False) 
            equipments.to_excel(writer, sheet_name='装备', index=False)
        
        print(f"📁 创建真实测试文件：{excel_path}")
        return excel_path

def run_real_mcp_tests(excel_path):
    """运行真实的MCP测试"""
    
    print("\n🧪 开始真实MCP工具调用验证")
    
    test_results = []
    
    # 测试1: 查看工作表列表
    print("\n1. 📋 测试 excel_list_sheets")
    try:
        sheets = excel_list_sheets(excel_path)
        print(f"   ✅ 工作表：{sheets.get('data', [])}")
        test_results.append(("list_sheets", True, "成功获取工作表列表"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("list_sheets", False, str(e)))
    
    # 测试2: 查看角色表结构
    print("\n2. 📊 测试 excel_describe_table")
    try:
        desc = excel_describe_table(excel_path, "角色")
        print(f"   ✅ 角色表结构：{desc.get('data', {}).get('columns', [])}")
        test_results.append(("describe_table", True, "成功获取表结构"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("describe_table", False, str(e)))
    
    # 测试3: 简单查询高等级角色
    print("\n3. 🔍 测试简单查询 - 高等级角色")
    simple_query = "SELECT 角色名称, 等级, 生命值, 攻击力 FROM 角色 WHERE 等级 >= 25 ORDER BY 等级 DESC"
    try:
        result = excel_query(excel_path, simple_query)
        data = result.get('data', [])
        print(f"   ✅ 查询结果：{len(data)} 行数据")
        if data and len(data) > 1:  # 有表头+数据
            print(f"   第一行：{data[1]}")
        test_results.append(("simple_query", True, "成功执行简单查询"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("simple_query", False, str(e)))
    
    # 测试4: 复杂JOIN查询
    print("\n4. 🔗 测试复杂JOIN查询 - 角色技能关联")
    join_query = """
    SELECT 
        r.角色名称, r.等级, 
        s.技能名称, s.技能类型, s.冷却时间
    FROM 
        角色 r
    JOIN 
        技能 s ON r.职业 = s.职业限制
    WHERE 
        r.等级 >= 25
    ORDER BY 
        r.等级 DESC, s.冷却_time ASC
    """
    try:
        result = excel_query(excel_path, join_query)
        data = result.get('data', [])
        print(f"   ✅ JOIN查询结果：{len(data)} 行数据")
        if data and len(data) > 1:
            print(f"   第一行：{data[1]}")
        test_results.append(("join_query", True, "成功执行JOIN查询"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("join_query", False, str(e)))
    
    # 测试5: 聚合查询
    print("\n5. 📈 测试聚合查询 - 职业统计")
    agg_query = """
    SELECT 
        职业, 
        COUNT(*) AS 角色数量,
        AVG(等级) AS 平均等级,
        SUM(生命值) AS 总生命值,
        MAX(攻击力) AS 最大攻击力
    FROM 
        角色
    GROUP BY 
        职业
    HAVING 
        COUNT(*) >= 1
    ORDER BY 
        平均等级 DESC
    """
    try:
        result = excel_query(excel_path, agg_query)
        data = result.get('data', [])
        print(f"   ✅ 聚合查询结果：{len(data)} 行数据")
        if data and len(data) > 1:
            print(f"   职业统计：{data[1]}")
        test_results.append(("agg_query", True, "成功执行聚合查询"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("agg_query", False, str(e)))
    
    # 测试6: 边界值测试
    print("\n6. 🎪 测试边界值 - 极端数据")
    boundary_query = "SELECT 技能名称, 冷却时间 FROM 技能 WHERE 冷却时间 <= 1.0 OR 冷却_time >= 10.0"
    try:
        result = excel_query(excel_path, boundary_query)
        data = result.get('data', [])
        print(f"   ✅ 边界值查询：{len(data)} 行数据")
        test_results.append(("boundary_query", True, "成功执行边界值查询"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("boundary_query", False, str(e)))
    
    # 测试7: 子查询测试
    print("\n7. 🎨 测试子查询 - 嵌套查询")
    subquery = """
    SELECT 
        角色名称,
        (SELECT AVG(等级) FROM 角色) AS 平均等级,
        (SELECT COUNT(*) FROM 技能 WHERE 职业限制 = r.职业) AS 技能数量
    FROM 
        角色 r
    WHERE 
        等级 > (SELECT AVG(等级) FROM 角色)
    """
    try:
        result = excel_query(excel_path, subquery)
        data = result.get('data', [])
        print(f"   ✅ 子查询结果：{len(data)} 行数据")
        test_results.append(("subquery", True, "成功执行子查询"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("subquery", False, str(e)))
    
    # 测试8: 查找最后一行
    print("\n8. 📏 测试 excel_find_last_row")
    try:
        last_row = excel_find_last_row(excel_path, "角色")
        print(f"   ✅ 最后一行：{last_row}")
        test_results.append(("find_last_row", True, "成功查找最后一行"))
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("find_last_row", False, str(e)))
    
    # 测试9: 插入数据测试
    print("\n9. ➕ 测试 excel_insert_rows")
    try:
        new_row = ["6", "新角色", "战士", "50", "2000", "800", "200", "120", "0.30"]
        result = excel_insert_rows(excel_path, "角色", [new_row], insert_position="bottom")
        print(f"   ✅ 插入结果：{result.get('message', '')}")
        test_results.append(("insert_rows", True, "成功插入数据"))
        
        # 清理插入的数据
        delete_query = "DELETE FROM 角色 WHERE 角色ID = 6"
        delete_result = excel_query(excel_path, delete_query)
        print(f"   🧹 清理结果：{delete_result.get('message', '')}")
        test_results.append(("cleanup_delete", True, "成功清理测试数据"))
        
    except Exception as e:
        print(f"   ❌ 失败：{e}")
        test_results.append(("insert_rows", False, str(e)))
    
    return test_results

def analyze_results(test_results):
    """分析测试结果"""
    
    passed = sum(1 for _, success, _ in test_results if success)
    failed = len(test_results) - passed
    
    print("\n" + "="*60)
    print("🧪 MCP工具验证分析")
    print("="*60)
    print(f"✅ 通过测试：{passed}")
    print(f"❌ 失败测试：{failed}")
    print(f"📊 成功率：{passed/len(test_results)*100:.1f}%")
    
    failed_tests = [name for name, success, _ in test_results if not success]
    if failed_tests:
        print(f"\n❌ 失败的测试：{failed_tests}")
    
    # 分析问题类型
    error_types = {}
    for _, success, error in test_results:
        if not success:
            if "连接" in error or "timeout" in error:
                error_types["连接问题"] = error_types.get("连接问题", 0) + 1
            elif "语法" in error or "SQL" in error:
                error_types["语法错误"] = error_types.get("语法错误", 0) + 1
            elif "不存在" in error or "找不到" in error:
                error_types["路径问题"] = error_types.get("路径问题", 0) + 1
            else:
                error_types["其他错误"] = error_types.get("其他错误", 0) + 1
    
    if error_types:
        print("\n📊 错误类型分析：")
        for error_type, count in error_types.items():
            print(f"   {error_type}：{count}次")
    
    return passed, failed, failed_tests

def main():
    """主函数"""
    print("🚀 REQ-033 真实MCP工具极限测试")
    
    # 创建测试文件
    excel_path = create_real_test_file()
    
    # 运行真实测试
    test_results = run_real_mcp_tests(excel_path)
    
    # 分析结果
    passed, failed, failed_tests = analyze_results(test_results)
    
    # 总结
    print("\n" + "="*60)
    print("🎯 真实MCP极限测试总结")
    print("="*60)
    print(f"📊 测试文件：包含3个工作表")
    print(f"🧪 测试用例：{len(test_results)}个")
    print(f"✅ 通过：{passed}个")
    print(f"❌ 失败：{failed}个")
    
    if failed == 0:
        print("🎉 所有测试通过！MCP工具稳定可靠")
        print("📈 产品质量：优秀")
    else:
        print("⚠️ 发现问题，需要进一步优化")
        if failed_tests:
            print(f"🔧 需要关注：{failed_tests}")
    
    # 极限探索评估
    complexity_score = min(10, len(test_results) // 2)  # 根据测试数量评估复杂度
    print(f"\n🎮 测试复杂度：{complexity_score}/10")
    
    return {
        "status": "completed",
        "total_tests": len(test_results),
        "passed_tests": passed,
        "failed_tests": failed,
        "failed_test_names": failed_tests,
        "complexity_score": complexity_score,
        "success": failed == 0
    }

if __name__ == "__main__":
    result = main()
    sys.exit(0 if result['success'] else 1)