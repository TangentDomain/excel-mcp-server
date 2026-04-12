#!/usr/bin/env python3
"""
检查docstring完整性的脚本
扫描src/excel_mcp_server_fastmcp/目录下所有.py文件，统计缺失Args/Parameters和Returns段的公共函数
"""

import os
import re
import ast
from pathlib import Path

def has_complete_docstring(docstring):
    """检查docstring是否包含Args/Parameters和Returns段"""
    if not docstring:
        return False
    
    docstring = docstring.strip()
    
    # 检查Args/Parameters
    has_args = bool(re.search(r'^\s*Args:\s*$', docstring, re.MULTILINE))
    has_params = bool(re.search(r'^\s*Parameters:\s*$', docstring, re.MULTILINE))
    
    # 检查Returns
    has_returns = bool(re.search(r'^\s*Returns:\s*$', docstring, re.MULTILINE))
    
    # 必须有Args/Parameters其中之一，并且有Returns
    return (has_args or has_params) and has_returns

def get_public_functions(tree):
    """从AST中提取公共函数（非下划线开头）"""
    public_functions = []
    
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef):
            # 只处理公共函数（非下划线开头）
            if not node.name.startswith('_'):
                # 检查是否有docstring
                docstring = ast.get_docstring(node)
                has_doc = docstring is not None
                has_complete = has_complete_docstring(docstring) if has_doc else False
                
                public_functions.append({
                    'name': node.name,
                    'lineno': node.lineno,
                    'has_docstring': has_doc,
                    'has_complete_docstring': has_complete
                })
    
    return public_functions

def scan_directory(directory):
    """扫描目录中的所有Python文件"""
    results = {}
    
    for root, dirs, files in os.walk(directory):
        # 排除__pycache__目录
        dirs[:] = [d for d in dirs if d != '__pycache__']
        
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                relative_path = os.path.relpath(filepath, directory)
                
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 解析AST
                    tree = ast.parse(content)
                    public_functions = get_public_functions(tree)
                    
                    # 统计结果
                    total_functions = len(public_functions)
                    with_docstring = sum(1 for f in public_functions if f['has_docstring'])
                    complete_docstring = sum(1 for f in public_functions if f['has_complete_docstring'])
                    missing_docstring = total_functions - with_docstring
                    incomplete_docstring = with_docstring - complete_docstring
                    
                    results[relative_path] = {
                        'total_functions': total_functions,
                        'with_docstring': with_docstring,
                        'complete_docstring': complete_docstring,
                        'missing_docstring': missing_docstring,
                        'incomplete_docstring': incomplete_docstring,
                        'functions': public_functions
                    }
                    
                except SyntaxError as e:
                    print(f"语法错误 {filepath}: {e}")
                except Exception as e:
                    print(f"处理文件 {filepath} 时出错: {e}")
    
    return results

def main():
    directory = Path('src/excel_mcp_server_fastmcp')
    
    if not directory.exists():
        print(f"错误: 目录 {directory} 不存在")
        return
    
    print("开始扫描docstring完整性...")
    print("=" * 60)
    
    results = scan_directory(directory)
    
    total_all = 0
    total_with_doc = 0
    total_complete = 0
    total_missing = 0
    total_incomplete = 0
    
    print(f"{'文件路径':<40} {'总数':<5} {'有文档':<8} {'完整':<8} {'缺失':<8} {'不完整':<8}")
    print("-" * 60)
    
    for filepath, data in sorted(results.items()):
        print(f"{filepath:<40} {data['total_functions']:<5} {data['with_docstring']:<8} "
              f"{data['complete_docstring']:<8} {data['missing_docstring']:<8} {data['incomplete_docstring']:<8}")
        
        total_all += data['total_functions']
        total_with_doc += data['with_docstring']
        total_complete += data['complete_docstring']
        total_missing += data['missing_docstring']
        total_incomplete += data['incomplete_docstring']
    
    print("=" * 60)
    print(f"总计: {total_all} 个函数, {total_with_doc} 有文档, {total_complete} 完整, "
          f"{total_missing} 缺失, {total_incomplete} 不完整")
    
    # 计算覆盖率
    coverage_rate = (total_complete / total_all * 100) if total_all > 0 else 0
    print(f"docstring完整率: {coverage_rate:.1f}%")
    
    # 输出缺失docstring的函数详情
    if total_missing > 0:
        print("\n缺失docstring的函数:")
        print("-" * 40)
        for filepath, data in results.items():
            if data['missing_docstring'] > 0:
                for func in data['functions']:
                    if not func['has_docstring']:
                        print(f"{filepath}:{func['name']} (第{func['lineno']}行)")

if __name__ == "__main__":
    main()