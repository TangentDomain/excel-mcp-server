import os
import sys
sys.path.append('src')

from excel_mcp_server_fastmcp.server import excel_list_sheets, excel_get_headers, _ok, _fail

# 创建一个简单的测试Excel文件来验证MCP功能
def create_test_excel():
    """创建一个测试用的Excel文件"""
    try:
        # 尝试导入并创建一个简单的测试文件
        import pandas as pd
        data = {
            '技能ID': [101, 102, 103],
            '技能名称': ['火球术', '冰箭术', '雷电术'],
            '伤害': [100, 80, 120],
            '冷却时间': [5, 6, 8]
        }
        df = pd.DataFrame(data)
        df.to_excel('test_skills.xlsx', index=False)
        print("✅ 创建测试Excel文件成功")
        return 'test_skills.xlsx'
    except Exception as e:
        print(f"❌ 创建测试Excel文件失败: {e}")
        return None

def test_mcp_functions():
    """测试MCP核心功能"""
    test_file = create_test_excel()
    if not test_file:
        return False
    
    tests_passed = 0
    total_tests = 0
    
    print("🧪 开始MCP功能验证...")
    
    # 测试1: 列出工作表
    total_tests += 1
    try:
        result = excel_list_sheets(test_file)
        if result.get('success') and 'sheets' in result:
            print(f"✅ 测试1通过: excel_list_sheets 返回 {len(result['sheets'])} 个工作表")
            tests_passed += 1
        else:
            print(f"❌ 测试1失败: excel_list_sheets 返回异常: {result}")
    except Exception as e:
        print(f"❌ 测试1异常: {e}")
    
    # 测试2: 获取表头
    total_tests += 1
    try:
        result = excel_get_headers(test_file)
        if result.get('success') and 'data' in result:
            headers_data = result['data']
            if 'sheets_with_headers' in headers_data:
                total_headers = sum(len(sheet['headers']) for sheet in headers_data['sheets_with_headers'])
                print(f"✅ 测试2通过: excel_get_headers 返回 {total_headers} 个表头")
                tests_passed += 1
            else:
                print(f"❌ 测试2失败: excel_get_headers 返回格式异常: {result}")
        else:
            print(f"❌ 测试2失败: excel_get_headers 返回异常: {result}")
    except Exception as e:
        print(f"❌ 测试2异常: {e}")
    
    # 测试3: _ok函数
    total_tests += 1
    try:
        result = _ok("测试成功", {"data": "test"}, {"meta": "test"})
        if result.get('success') and result.get('message') == "测试成功":
            print("✅ 测试3通过: _ok函数正常工作")
            tests_passed += 1
        else:
            print(f"❌ 测试3失败: _ok函数返回异常: {result}")
    except Exception as e:
        print(f"❌ 测试3异常: {e}")
    
    # 测试4: _fail函数
    total_tests += 1
    try:
        result = _fail("测试失败", {"error_code": "TEST_ERROR"})
        if not result.get('success') and result.get('message') == "测试失败":
            print("✅ 测试4通过: _fail函数正常工作")
            tests_passed += 1
        else:
            print(f"❌ 测试4失败: _fail函数返回异常: {result}")
    except Exception as e:
        print(f"❌ 测试4异常: {e}")
    
    # 清理测试文件
    try:
        if os.path.exists(test_file):
            os.remove(test_file)
            print("🧹 测试文件已清理")
    except:
        pass
    
    print(f"\n📊 MCP验证结果: {tests_passed}/{total_tests} 通过")
    return tests_passed == total_tests

if __name__ == "__main__":
    success = test_mcp_functions()
    sys.exit(0 if success else 1)