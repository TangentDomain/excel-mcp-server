#!/usr/bin/env python3
"""
MCP功能快速验证脚本
测试12项核心功能是否正常工作
"""
import os
import sys
import tempfile
import json
from pathlib import Path

# 添加源码路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

def test_mcp_core_functions():
    """测试MCP核心功能"""
    from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations
    from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
    
    # 创建临时测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
        temp_path = temp_file.name
    
    try:
        # 初始化Excel操作
        excel_ops = ExcelOperations()
        sql_engine = AdvancedSQLQueryEngine()
        
        test_results = []
        
        # 1. 测试工作表操作
        try:
            # 创建测试工作表
            test_data = [
                ["姓名", "年龄", "部门"],
                ["张三", "25", "研发"],
                ["李四", "30", "产品"],
                ["王五", "28", "测试"]
            ]
            
            excel_ops.create_workbook(temp_path)
            excel_ops.write_range(temp_path, "Sheet1", "A1", test_data)
            test_results.append(("create_workbook", "✅"))
        except Exception as e:
            test_results.append(("create_workbook", f"❌ {e}"))
        
        # 2. 测试读取数据
        try:
            data = excel_ops.get_range(temp_path, "Sheet1", "A1:C3")
            assert len(data) == 3 and len(data[0]) == 3
            test_results.append(("get_range", "✅"))
        except Exception as e:
            test_results.append(("get_range", f"❌ {e}"))
        
        # 3. 测试SQL查询
        try:
            # 创建一个包含测试数据的Excel文件
            excel_ops.create_workbook(temp_path)
            excel_ops.write_range(temp_path, "Sheet1", "A1", test_data)
            
            # 执行简单查询
            query = "SELECT * FROM Sheet1 WHERE 年龄 > 25"
            result = sql_engine.execute_query(temp_path, query)
            assert len(result["data"]) >= 1
            test_results.append(("sql_query", "✅"))
        except Exception as e:
            test_results.append(("sql_query", f"❌ {e}"))
        
        # 4. 测试写入操作
        try:
            new_data = [["赵六", "35", "运营"]]
            excel_ops.write_range(temp_path, "Sheet1", "A5", new_data)
            test_results.append(("write_range", "✅"))
        except Exception as e:
            test_results.append(("write_range", f"❌ {e}"))
        
        # 5. 测试获取表头
        try:
            headers = excel_ops.get_headers(temp_path, "Sheet1")
            assert len(headers) == 3 and "姓名" in headers
            test_results.append(("get_headers", "✅"))
        except Exception as e:
            test_results.append(("get_headers", f"❌ {e}"))
        
        # 6. 测试删除行
        try:
            excel_ops.delete_rows(temp_path, "Sheet1", condition="姓名 = '赵六'")
            test_results.append(("delete_rows", "✅"))
        except Exception as e:
            test_results.append(("delete_rows", f"❌ {e}"))
        
        # 7. 测试批插入
        try:
            batch_data = [["钱七", "40", "市场"], ["孙八", "32", "销售"]]
            excel_ops.batch_insert_rows(temp_path, "Sheet1", 6, batch_data)
            test_results.append(("batch_insert_rows", "✅"))
        except Exception as e:
            test_results.append(("batch_insert_rows", f"❌ {e}"))
        
        # 8. 测试查找最后一行
        try:
            last_row = excel_ops.find_last_row(temp_path, "Sheet1")
            assert last_row >= 6
            test_results.append(("find_last_row", "✅"))
        except Exception as e:
            test_results.append(("find_last_row", f"❌ {e}"))
        
        return test_results
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_path):
            os.unlink(temp_path)

def main():
    """主函数"""
    print("🧪 执行MCP核心功能验证...")
    
    try:
        results = test_mcp_core_functions()
        
        passed = sum(1 for _, status in results if status == "✅")
        total = len(results)
        
        print(f"\n📊 验证结果: {passed}/{total} 通过")
        
        for func, status in results:
            print(f"  {func}: {status}")
        
        if passed == total:
            print("✅ 所有核心功能验证通过")
            return 0
        else:
            print("❌ 部分功能验证失败")
            return 1
            
    except Exception as e:
        print(f"❌ 验证过程中发生错误: {e}")
        return 1

if __name__ == "__main__":
    exit(main())