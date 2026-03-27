#!/usr/bin/env python3
"""
测试批量操作功能
"""
import sys
import os
import json
from pathlib import Path

# 添加路径
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')

# 导入函数
try:
    from excel_mcp_server_fastmcp.server import (
        excel_batch_insert_rows, excel_delete_rows
    )
    print("✅ 成功导入批量操作函数")
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    sys.exit(1)

def test_batch_insert():
    """测试批量插入"""
    test_file = '/tmp/excel_mcp_test_fs5p18el/test_game_data.xlsx'
    
    print("\n🧪 测试批量插入行")
    print("-" * 30)
    
    try:
        # 先检查当前行数
        from excel_mcp_server_fastmcp.server import excel_describe_table
        before_result = excel_describe_table(file_path=test_file, sheet_name='角色')
        print(f"插入前: {before_result['data']['row_count']} 行")
        
        # 插入新行
        result = excel_batch_insert_rows(
            file_path=test_file,
            sheet_name='角色',
            data=[
                {'ID': 6, '名称': '游侠', '职业': '游侠', '等级': 46, '属性': '自然'},
                {'ID': 7, '名称': '圣骑士', '职业': '圣骑士', '等级': 52, '属性': '光'}
            ],
            header_row=1,
            streaming=True
        )
        
        print(f"✅ 插入结果: {result['message']}")
        
        # 检查插入后行数
        after_result = excel_describe_table(file_path=test_file, sheet_name='角色')
        print(f"插入后: {after_result['data']['row_count']} 行")
        
        return True
        
    except Exception as e:
        print(f"❌ 插入失败: {str(e)}")
        return False

def test_delete_rows():
    """测试删除行"""
    test_file = '/tmp/excel_mcp_test_fs5p18el/test_game_data.xlsx'
    
    print("\n🧪 测试删除行")
    print("-" * 30)
    
    try:
        # 先检查当前行数
        from excel_mcp_server_fastmcp.server import excel_describe_table
        before_result = excel_describe_table(file_path=test_file, sheet_name='角色')
        print(f"删除前: {before_result['data']['row_count']} 行")
        
        # 删除行（删除第9行，刚插入的）
        result = excel_delete_rows(
            file_path=test_file,
            sheet_name='角色',
            row_index=8,  # 第9行（从0开始是8）
            count=1,
            streaming=True
        )
        
        print(f"✅ 删除结果: {result['message']}")
        
        # 检查删除后行数
        after_result = excel_describe_table(file_path=test_file, sheet_name='角色')
        print(f"删除后: {after_result['data']['row_count']} 行")
        
        return True
        
    except Exception as e:
        print(f"❌ 删除失败: {str(e)}")
        return False

if __name__ == "__main__":
    print("🧪 批量操作功能测试（第145轮）")
    print("=" * 60)
    
    batch_success = test_batch_insert()
    delete_success = test_delete_rows()
    
    print("\n" + "=" * 60)
    print(f"📊 批量操作结果: {'全部通过' if batch_success and delete_success else '存在问题'}")
    
    if batch_success and delete_success:
        print("🎉 批量操作功能正常！")
    else:
        print("⚠️ 需要修复批量操作问题")