#!/usr/bin/env python3
"""
Docstring批量添加脚本
为缺失docstring的函数添加标准格式的文档
"""

import re
import ast
from pathlib import Path

def generate_docstring(func_name, file_path):
    """根据函数名和文件路径生成docstring"""
    
    # 基础模板
    docstring_templates = {
        '__init__': f'''"""
初始化{func_name}组件

Args:
    file_path (str): Excel文件路径
    session_id (str, optional): 会话ID
"""''',
        'wrapper': f'''"""
包装函数，提供统一的错误处理和日志记录

Args:
    func: 被包装的函数
    *args: 位置参数
    **kwargs: 关键字参数

Returns:
    Any: 函数执行结果

Examples:
    >>> @wrapper
    ... def my_function():
    ...     pass
"""''',
        'format': f'''"""
格式化输出数据

Args:
    data (Any): 要格式化的数据
    format_type (str, optional): 格式类型

Returns:
    str: 格式化后的字符串
"""''',
        'excel_search': f'''"""
Excel文件搜索功能

Args:
    query (str): 搜索查询
    file_path (str): Excel文件路径
    sheet_name (str, optional): 工作表名称

Returns:
    List[Dict]: 搜索结果
"""''',
        'excel_search_directory': f'''"""
目录内Excel文件批量搜索

Args:
    query (str): 搜索查询
    directory (str): 目录路径
    recursive (bool, optional): 是否递归搜索

Returns:
    List[Dict]: 搜索结果
"""''',
        'protect_string': f'''"""
字符串保护，防止SQL注入

Args:
    value (str): 要保护的字符串值

Returns:
    str: 保护后的字符串
"""''',
        '_extract_selects': f'''"""
提取SQL查询中的SELECT字段

Args:
    query (str): SQL查询语句

Returns:
    List[str]: SELECT字段列表
"""''',
        'assign_row_number': f'''"""
为查询结果分配行号

Args:
    df (DataFrame): pandas DataFrame
    partition_by (List[str]): 分组字段
    order_by (str): 排序字段

Returns:
    DataFrame: 添加了行号的DataFrame
"""''',
        'assign_rank': f'''"""
为查询结果分配排名

Args:
    df (DataFrame): pandas DataFrame
    partition_by (List[str]): 分组字段
    order_by (str): 排序字段

Returns:
    DataFrame: 添加了排名的DataFrame
"""''',
        '_track_call': f'''"""
工具调用追踪装饰器

Args:
    func: 被装饰的函数

Returns:
    function: 装饰后的函数
"""''',
        '_save_log': f'''"""
保存日志到文件

Args:
    log_data (Dict): 日志数据

Returns:
    bool: 保存是否成功
"""''',
        'log_operation': f'''"""
记录操作到日志

Args:
    operation (str): 操作描述
    duration (float): 操作耗时（秒）
    result (Any): 操作结果
    error (str, optional): 错误信息

Returns:
    bool: 记录是否成功
"""''',
        'get_recent_operations': f'''"""
获取最近的操作记录

Args:
    limit (int, optional): 返回记录数量限制

Returns:
    List[Dict]: 操作记录列表
"""''',
        'record': f'''"""
记录一次工具调用

Args:
    tool_name (str): 工具名称
    args (Dict): 工具参数
    result (Any): 执行结果
    duration (float): 执行耗时
    error (str, optional): 错误信息

Returns:
    bool: 记录是否成功
"""''',
        'classify_error': f'''"""
根据错误消息自动分类错误类型

Args:
    error_message (str): 错误消息

Returns:
    str: 错误类型分类
"""''',
        'get_stats': f'''"""
获取工具调用统计

Args:
    None

Returns:
    Dict: 统计信息
"""''',
        'reset': f'''"""
重置所有统计信息

Args:
    None

Returns:
    bool: 重置是否成功
"""''',
        'start_session': f'''"""
开始新的操作会话

Args:
    file_path (str): Excel文件路径
    session_id (str, optional): 会话ID

Returns:
    str: 会话ID
"""''',
        'main': f'''"""
MCP服务器主入口

Args:
    config (Dict, optional): 配置信息

Returns:
    MCP: MCP服务器实例
"""'''
    }
    
    # 使用基础模板
    if func_name in docstring_templates:
        return docstring_templates[func_name]
    
    # 通用模板
    return f'''"""
{func_name}函数的功能描述

Args:
    *args: 位置参数
    **kwargs: 关键字参数

Returns:
    Any: 函数返回值

Examples:
    >>> {func_name}(*args, **kwargs)
    # 示例调用
"""'''

def add_docstrings_to_file(file_path):
    """为文件中缺失docstring的函数添加docstring"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 解析AST
    try:
        tree = ast.parse(content)
    except SyntaxError:
        print(f"⚠️ 语法错误 {file_path}")
        return False
    
    # 收集需要添加docstring的函数
    functions_to_update = []
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and not ast.get_docstring(node):
            functions_to_update.append({
                'name': node.name,
                'line': node.lineno,
                'end_line': node.end_lineno
            })
    
    if not functions_to_update:
        print(f"✅ {file_path}: 无需更新")
        return True
    
    print(f"📝 {file_path}: 发现 {len(functions_to_update)} 个函数需要添加docstring")
    
    # 按行号倒序排序，避免影响后续行号
    functions_to_update.sort(key=lambda x: x['line'], reverse=True)
    
    # 修改文件内容
    lines = content.split('\n')
    
    for func in functions_to_update:
        docstring = generate_docstring(func['name'], str(file_path))
        
        # 在函数定义后插入docstring
        insert_line = func['line']  # 转换为0-based
        indent = '    '  # 4个空格的缩进
        
        # 插入docstring
        docstring_lines = docstring.split('\n')
        for i, line in enumerate(docstring_lines):
            lines.insert(insert_line + i, indent + line)
        
        # 调整后续函数的行号
        line_offset = len(docstring_lines)
        for other_func in functions_to_update:
            if other_func['line'] > func['line']:
                other_func['line'] += line_offset
    
    # 写回文件
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    print(f"✅ {file_path}: 成功添加 {len(functions_to_update)} 个docstring")
    return True

def main():
    """主函数"""
    src_dir = Path('src/excel_mcp_server_fastmcp')
    processed_files = 0
    total_functions = 0
    
    print("🔍 开始批量添加docstring...")
    
    for py_file in src_dir.rglob('*.py'):
        if add_docstrings_to_file(py_file):
            processed_files += 1
            # 统计该文件的函数数量
            with open(py_file, 'r', encoding='utf-8') as f:
                content = f.read()
            functions = len(re.findall(r'^\s*def\s+\w+', content, re.MULTILINE))
            total_functions += functions
    
    print(f"\n📊 处理完成:")
    print(f"   - 处理文件数: {processed_files}")
    print(f"   - 总函数数: {total_functions}")
    print("✅ 批量docstring添加完成")

if __name__ == '__main__':
    main()