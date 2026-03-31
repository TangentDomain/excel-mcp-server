#!/usr/bin/env python3
"""
使用正确的函数名和参数测试5个API问题
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

import tempfile
import pandas as pd
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

def create_test_excel():
    """创建测试用的Excel文件"""
    # 创建测试数据
    data = {
        'A': [1, 2, 3, 4, 5],
        'B': ['a', 'b', 'c', 'd', 'e'],
        'C': [10.5, 20.5, 30.5, 40.5, 50.5]
    }
    df = pd.DataFrame(data)
    
    # 创建临时文件
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_file.close()
    
    # 写入Excel
    df.to_excel(temp_file.name, index=False, sheet_name="Sheet1")
    return temp_file.name

def test_api_issues():
    """测试5个API问题"""
    excel_file = create_test_excel()
    sheet_name = "Sheet1"
    
    print(f"测试文件: {excel_file}")
    print("=" * 60)
    
    excel_ops = ExcelOperations()
    
    issues_found = []
    
    try:
        # 测试1: get_range 范围查询（参数顺序问题）
        print("\n1. 测试 get_range 范围查询...")
        try:
            # 正常情况：工作表名在range_expression中
            result1 = excel_ops.get_range(
                file_path=excel_file,
                range_expression="Sheet1!A1:C3"
            )
            print(f"正常情况结果: {result1}")
            
            # 测试颠倒的范围表达式（A1:C3 vs C3:A1）
            result2 = excel_ops.get_range(
                file_path=excel_file,
                range_expression="Sheet1!C3:A1"  # 颠倒的单元格范围
            )
            print(f"颠倒范围结果: {result2}")
            
            # 检查是否正确处理颠倒的范围
            if result1.get('success') and result2.get('success'):
                # 如果都能成功，需要检查数据是否正确
                data1 = result1.get('data', [])
                data2 = result2.get('data', [])
                if data1 != data2:
                    print("⚠️ 颠倒范围返回不同数据，可能存在参数顺序问题")
                    issues_found.append("get_range范围参数处理不当")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("get_range参数顺序处理异常")
        
        # 测试2: format_cells 缺少必要参数时的处理
        print("\n2. 测试 format_cells 缺少必要参数...")
        try:
            # 故意不传formatting参数，看看是否能优雅处理
            result = excel_ops.format_cells(
                file_path=excel_file,
                sheet_name=sheet_name,
                range="A1:A5",
                # 不传formatting参数，应该使用默认值或给出清晰错误
            )
            print(f"无formatting参数结果: {result}")
            
            if result.get('success') == False:
                if "缺少" in result.get('message', '') or "required" in result.get('message', '').lower():
                    issues_found.append("format_cells缺少参数校验不清晰")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("format_cells缺少参数处理异常")
        
        # 测试3: set_formula 缺少formula参数
        print("\n3. 测试 set_formula 缺少formula参数...")
        try:
            # 故意不传formula参数
            result = excel_ops.set_formula(
                file_path=excel_file,
                sheet_name=sheet_name,
                cell_range="A1",
                # 故意不传formula参数
            )
            print(f"无formula参数结果: {result}")
            
            if result.get('success') == False:
                if "formula" in result.get('message', '').lower() or "required" in result.get('message', '').lower():
                    print("✅ 正确校验了formula参数")
                else:
                    issues_found.append("set_formula缺少formula参数校验不清晰")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("set_formula缺少参数处理异常")
        
        # 测试4: get_range 搜索逻辑（sheet_name参数混淆）
        print("\n4. 测试 get_range sheet_name参数处理...")
        try:
            # 测试不包含工作表名的range_expression
            result = excel_ops.get_range(
                file_path=excel_file,
                range_expression="A1:B3"  # 不包含工作表名
            )
            print(f"无工作表名结果: {result}")
            
            # 应该返回明确的错误信息
            if result.get('success') == False:
                error_msg = result.get('message', '')
                if '工作表' in error_msg or 'sheet' in error_msg.lower():
                    print("✅ 正确提示了工作表名缺失")
                else:
                    issues_found.append("get_range工作表名错误提示不清晰")
            else:
                issues_found.append("get_range接受了无效的range_expression格式")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("get_range工作表名处理异常")
        
        # 测试5: update_range 数据格式不匹配时的处理
        print("\n5. 测试 update_range 数据格式校验...")
        try:
            # 传入错误的数据格式（应该传list of lists，但传string）
            result = excel_ops.update_range(
                file_path=excel_file,
                range_expression="Sheet1!A1:A5",
                data="invalid_format",  # 错误格式：应该是二维数组
                preserve_formulas=True,
                insert_mode=False
            )
            print(f"错误数据格式结果: {result}")
            
            if result.get('success') == False:
                error_msg = result.get('message', '')
                if '二维数组' in error_msg or 'list' in error_msg.lower():
                    print("✅ 正确校验了数据格式")
                else:
                    issues_found.append("update_range数据格式校验不清晰")
            else:
                issues_found.append("update_range接受了无效的数据格式")
            
        except Exception as e:
            print(f"错误: {e}")
            issues_found.append("update_range数据格式处理异常")
    
    finally:
        # 清理临时文件
        os.unlink(excel_file)
    
    print("\n" + "=" * 60)
    print("🔍 API问题总结:")
    if issues_found:
        for i, issue in enumerate(issues_found, 1):
            print(f"{i}. ❌ {issue}")
        print(f"\n📊 发现 {len(issues_found)} 个需要修复的API问题")
    else:
        print("✅ 所有API测试通过，未发现问题")
    
    return issues_found

if __name__ == "__main__":
    test_api_issues()