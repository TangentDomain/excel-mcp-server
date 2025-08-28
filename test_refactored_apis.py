#!/usr/bin/env python
"""
测试重构后的核心API功能
"""

from src.server import excel_get_range, excel_update_range, excel_list_sheets, excel_get_headers, excel_create_file
import tempfile
import os

def test_refactored_apis():
    # 测试创建文件
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_path = temp_file.name
    temp_file.close()

    try:
        # 测试创建文件
        print('测试创建文件...')
        result = excel_create_file(temp_path, ['测试表', '数据表'])
        print(f'创建文件结果: {result.get("success", False)}')

        if result.get('success'):
            # 测试列出工作表
            print('测试列出工作表...')
            sheets_result = excel_list_sheets(temp_path)
            print(f'工作表列表: {sheets_result.get("sheets", [])}')

            # 测试写入数据
            print('测试写入数据...')
            test_data = [['ID', '姓名', '年龄'], [1, '张三', 25], [2, '李四', 30]]
            update_result = excel_update_range(temp_path, '测试表!A1:C3', test_data)
            print(f'写入结果: {update_result.get("success", False)}')

            # 测试获取表头
            print('测试获取表头...')
            headers_result = excel_get_headers(temp_path, '测试表')
            print(f'表头信息: {headers_result.get("headers", [])}')

            # 测试读取数据
            print('测试读取数据...')
            read_result = excel_get_range(temp_path, '测试表!A1:C3')
            print(f'读取结果: {read_result.get("success", False)}')

        print('所有API测试完成!')

    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)

if __name__ == '__main__':
    test_refactored_apis()
