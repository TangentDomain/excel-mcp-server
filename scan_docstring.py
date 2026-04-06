#!/usr/bin/env python3
import os
import ast
import re
from pathlib import Path

def scan_functions_for_docstrings():
    """扫描所有Python文件中的公共函数，统计docstring情况"""
    src_dir = Path("src/excel_mcp_server_fastmcp")
    results = []
    
    for py_file in src_dir.rglob("*.py"):
        if py_file.name.startswith("__"):
            continue
            
        try:
            with open(py_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 解析AST
            tree = ast.parse(content, filename=str(py_file))
            
            for node in ast.walk(tree):
                # 只检查函数定义（非私有函数）
                if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
                    # 检查是否有docstring
                    has_docstring = (
                        node.body and 
                        isinstance(node.body[0], ast.Expr) and
                        isinstance(node.body[0].value, ast.Constant) and
                        isinstance(node.body[0].value.value, str)
                    )
                    
                    # 检查docstring是否包含Args/Parameters和Returns
                    docstring_text = ""
                    if has_docstring:
                        docstring_text = node.body[0].value.value.strip()
                    
                    has_args = bool(re.search(r'\bArgs\b|\bParameters\b', docstring_text, re.IGNORECASE))
                    has_returns = bool(re.search(r'\bReturns\b', docstring_text, re.IGNORECASE))
                    
                    results.append({
                        'file': str(py_file),
                        'function': node.name,
                        'has_docstring': has_docstring,
                        'has_args': has_args,
                        'has_returns': has_returns,
                        'docstring_complete': has_args and has_returns
                    })
                    
        except Exception as e:
            print(f"Error parsing {py_file}: {e}")
    
    # 统计结果
    total_functions = len(results)
    complete_functions = sum(1 for r in results if r['docstring_complete'])
    incomplete_functions = total_functions - complete_functions
    
    print(f"扫描完成：")
    print(f"总函数数：{total_functions}")
    print(f"完整docstring函数数：{complete_functions}")
    print(f"不完整docstring函数数：{incomplete_functions}")
    print(f"合规率：{complete_functions/total_functions*100:.1f}%" if total_functions > 0 else "合规率：0%")
    
    # 输出不完整的函数详情
    print("\n需要修复的函数：")
    for result in results:
        if not result['docstring_complete']:
            print(f"  {result['file']}:{result['function']} (has_docstring={result['has_docstring']}, has_args={result['has_args']}, has_returns={result['has_returns']})")
    
    return results

if __name__ == "__main__":
    scan_functions_for_docstrings()
