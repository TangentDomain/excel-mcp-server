#!/usr/bin/env python3
"""
错误处理改进脚本
为关键Excel操作函数添加更好的错误处理和AI修复建议
"""

import re
from pathlib import Path

def improve_error_handling(file_path):
    """改进指定文件中的错误处理"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    
    # 为关键Excel操作函数添加更好的错误处理
    key_functions = [
        'excel_get_range',
        'excel_update_range', 
        'excel_insert_rows',
        'excel_search',
        'excel_get_headers'
    ]
    
    for func in key_functions:
        # 找到函数并改进错误处理
        for i, line in enumerate(lines):
            if line.strip().startswith(f'def {func}('):
                # 找到函数的try-except块
                for j in range(i, min(i+100, len(lines))):
                    if 'try:' in lines[j]:
                        # 找到except块
                        for k in range(j, min(j+50, len(lines))):
                            if 'except Exception as e:' in lines[k]:
                                # 在except块前添加更详细的错误处理
                                indent = '            '
                                improved_error = f'''{indent}# 改进的错误处理，提供AI友好的错误信息
{indent}import sys
{indent}from ..utils.exceptions import ExcelException, SheetNotFoundError, DataValidationError
{indent}error_type = 'UNKNOWN_ERROR'
{indent}suggested_fix = '请检查Excel文件格式和路径是否正确'
{indent}context = {'file_path': file_path}
{indent}
{indent}try:
{indent}    # 原有的操作代码
{indent}'''
                                
                                # 在try块后添加错误分类和处理
                                improved_except = f'''{indent}except SheetNotFoundError as e:
{indent}    error_type = 'SHEET_NOT_FOUND'
{indent}    suggested_fix = f'工作表"{{getattr(e, "sheet_name", "unknown")}}"不存在，请检查工作表名称'
{indent}    context['sheet_name'] = getattr(e, 'sheet_name', 'unknown')
{indent}    raise ExcelException(
{indent}        message=str(e), 
{indent}        hint=f'无法找到工作表: {{getattr(e, "sheet_name", "unknown")}}',
{indent}        suggested_fix=suggested_fix,
{indent}        context=context
{indent}    )
{indent}except DataValidationError as e:
{indent}    error_type = 'DATA_VALIDATION_ERROR'
{indent}    suggested_fix = f'数据格式错误: {{str(e)}}，请检查数据是否符合预期格式'
{indent}    context['validation_error'] = str(e)
{indent}    raise ExcelException(
{indent}        message=str(e), 
{indent}        hint=f'数据验证失败: {{str(e)}}',
{indent}        suggested_fix=suggested_fix,
{indent}        context=context
{indent}except Exception as e:
{indent}    error_type = 'EXCEL_OPERATION_ERROR'
{indent}    suggested_fix = 'Excel操作失败，请检查文件格式、权限或网络连接'
{indent}    context['original_error'] = str(e)
{indent}    context['error_type'] = type(e).__name__
{indent}    raise ExcelException(
{indent}        message=f'Excel操作失败: {{str(e)}}', 
{indent}        hint=f'操作类型: {func.replace("_", " ")}',
{indent}        suggested_fix=suggested_fix,
{indent}        context=context
{indent})'''
                                
                                # 替换原有的try-except块
                                try_end = j + 1
                                while try_end < len(lines) and not lines[try_end].strip().startswith('except'):
                                    try_end += 1
                                
                                if try_end < len(lines):
                                    # 删除原有的try-except块
                                    del lines[j:try_end + 1]
                                    # 插入改进的代码
                                    lines.insert(j, improved_except)
                                    break
                        break
    
    # 写回文件
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    return True

def main():
    """主函数"""
    src_dir = Path('src/excel_mcp_server_fastmcp')
    server_file = src_dir / 'server.py'
    
    print("🔧 改进Excel操作函数的错误处理...")
    
    if improve_error_handling(server_file):
        print("✅ 成功改进错误处理")
    else:
        print("❌ 改进错误处理失败")
    
    # 验证文件语法
    try:
        import py_compile
        py_compile.compile(str(server_file), doraise=True)
        print("✅ 文件语法验证通过")
    except py_compile.PyCompileError as e:
        print(f"❌ 语法错误: {e}")
        return False
    
    return True

if __name__ == '__main__':
    success = main()
    exit(0 if success else 1)