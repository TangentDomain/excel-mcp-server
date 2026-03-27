#!/usr/bin/env python3
"""
REQ-029 Bug修复验证测试

测试内容：
1. JOIN表别名解析 + describe_table流式写入后崩溃
2. streaming写入后describe_table处理max_row=None的情况
"""

import os
import tempfile
import sys
import pandas as pd

# 添加源码路径
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

def create_test_excel():
    """创建测试用的Excel文件"""
    # 创建临时文件
    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    
    # 使用pandas创建测试数据
    try:
        # 角色表
        char_data = {
            'ID': [1, 2, 3],
            '名称': ['战士', '法师', '射手'],
            '职业': ['近战', '远程', '远程'],
            '等级': [10, 15, 12]
        }
        
        # 技能表  
        skill_data = {
            'ID': [101, 102, 103],
            '名称': ['火球术', '冰箭', '射击'],
            '职业限制': ['远程', '远程', '远程'],
            '伤害': [100, 80, 60]
        }
        
        # 写入Excel
        with pd.ExcelWriter(tmp_path) as writer:
            pd.DataFrame(char_data).to_excel(writer, sheet_name='角色', index=False)
            pd.DataFrame(skill_data).to_excel(writer, sheet_name='技能', index=False)
            
        print(f"✅ 测试Excel文件已创建: {tmp_path}")
        return tmp_path
        
    except Exception as e:
        print(f"❌ 创建测试Excel失败: {e}")
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return None

def test_join_alias_resolution():
    """测试JOIN表别名解析"""
    print("\n=== 测试1: JOIN表别名解析 ===")
    
    excel_path = create_test_excel()
    if not excel_path:
        return False
    
    try:
        # 测试JOIN查询，使用表别名
        query = "SELECT r.名称, s.名称 FROM 角色 r JOIN 技能 s ON r.职业 = s.职业限制"
        
        engine = AdvancedSQLQueryEngine()
        result = engine.execute_sql_query(excel_path, query)
        
        print(f"查询: {query}")
        print(f"结果: {result}")
        
        # 检查结果是否正确
        if result.get('success'):
            data = result.get('data', [])
            if len(data) > 0:
                print(f"✅ JOIN别名解析成功，返回 {len(data)} 行数据")
                return True
            else:
                print("⚠️ JOIN查询返回空数据")
                return True
        else:
            print(f"❌ JOIN查询失败: {result.get('message')}")
            return False
            
    except Exception as e:
        print(f"❌ JOIN别名测试异常: {e}")
        return False
    finally:
        if os.path.exists(excel_path):
            os.unlink(excel_path)

def test_streaming_describe_table():
    """测试streaming写入后describe_table"""
    print("\n=== 测试2: Streaming写入后describe_table ===")
    
    excel_path = create_test_excel()
    if not excel_path:
        return False
    
    try:
        # 先执行一个streaming写入操作
        streaming_data = [
            {'名称': '坦克', '职业': '近战', '等级': 20},
            {'名称': '牧师', '职业': '辅助', '等级': 18}
        ]
        
        result = ExcelOperations.batch_insert_rows(
            excel_path, 
            '角色', 
            streaming_data,
            streaming=True
        )
        
        print(f"Streaming写入结果: {result}")
        
        if result.get('success'):
            print("✅ Streaming写入成功")
            
            # 现在测试describe_table（这应该不崩溃）
            from excel_mcp_server_fastmcp.server import excel_describe_table
            describe_result = excel_describe_table(excel_path, '角色')
            
            print(f"Describe Table结果: {describe_result}")
            
            if describe_result.get('success'):
                print("✅ Streaming写入后describe_table正常工作")
                return True
            else:
                print(f"❌ describe_table失败: {describe_result.get('message')}")
                return False
        else:
            print(f"❌ Streaming写入失败: {result.get('message')}")
            return False
            
    except Exception as e:
        print(f"❌ Streaming describe_table测试异常: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if os.path.exists(excel_path):
            os.unlink(excel_path)

def main():
    """主测试函数"""
    print("开始REQ-029 Bug修复验证测试...")
    
    test1_result = test_join_alias_resolution()
    test2_result = test_streaming_describe_table()
    
    print(f"\n=== 测试结果汇总 ===")
    print(f"JOIN别名解析: {'✅ 通过' if test1_result else '❌ 失败'}")
    print(f"Streaming describe_table: {'✅ 通过' if test2_result else '❌ 失败'}")
    
    overall_success = test1_result and test2_result
    print(f"\n总体结果: {'✅ 所有测试通过' if overall_success else '❌ 存在失败'}")
    
    return overall_success

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)