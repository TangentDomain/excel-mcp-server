#!/usr/bin/env python3
"""
为核心函数添加docstring的精确脚本
正确识别函数定义并在函数体开始处添加docstring
"""

import re
from pathlib import Path

def add_docstring_to_file(file_path, target_functions):
    """为指定文件中的目标函数添加docstring"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    
    for target_func in target_functions:
        # Find the function definition and locate the function body start
        for i, line in enumerate(lines):
            if line.strip().startswith(f'def {target_func}('):
                # Found the function definition, find the end of parameters
                j = i
                while j < len(lines) and not lines[j].strip().endswith(':'):
                    j += 1
                
                if j < len(lines):
                    # Found the end of parameters (the line with ':')
                    # Look for the start of function body
                    k = j + 1
                    while k < len(lines) and lines[k].strip() == '':
                        k += 1
                    
                    if k < len(lines):
                        # Insert docstring before the first line of function body
                        indent = '        '  # 8 spaces for docstring
                        docstring = f'''{indent}\"\"\"
{indent}Excel {target_func.replace('_', ' ')} function

{indent}Args:
{indent}    file_path (str): Excel file path
{indent}    *args: Additional arguments specific to the function
{indent}    **kwargs: Keyword arguments specific to the function

{indent}Returns:
{indent}    Various: Result depends on the specific function
{indent}\"\"\"
{indent}'''
                        
                        lines.insert(k, docstring)
                        break
    
    # Write back to file
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    return True

def main():
    """主函数"""
    src_dir = Path('src/excel_mcp_server_fastmcp')
    server_file = src_dir / 'server.py'
    
    # 选择核心的Excel操作函数
    core_functions = [
        'excel_search',
        'excel_search_directory', 
        'excel_get_range',
        'excel_get_headers',
        'excel_update_range',
        'excel_insert_rows'
    ]
    
    print(f"📝 为 {len(core_functions)} 个核心函数添加docstring...")
    
    if add_docstring_to_file(server_file, core_functions):
        print("✅ 成功添加核心函数docstring")
    else:
        print("❌ 添加docstring失败")
    
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