#!/usr/bin/env python3
"""
Docstring质量检查脚本
检查项目中的docstring质量，统计评分和改进建议
"""

import re
import ast
import os
from pathlib import Path

def analyze_docstring_quality():
    """分析项目中的docstring质量"""
    
    excel_mcp_dir = Path("src/excel_mcp_server_fastmcp")
    docstring_stats = {
        'excellent': 0,  # 包含参数、返回值、异常
        'good': 0,       # 包含参数和返回值
        'basic': 0,      # 只有简单描述
        'none': 0,       # 没有docstring
        'issues': []     # 发现的问题
    }
    
    def get_docstring_score(docstring):
        """评估单个docstring的评分"""
        if not docstring:
            return 'none'
        
        docstring = docstring.strip()
        
        # 检查是否包含参数描述
        has_params = ':param' in docstring or ':type' in docstring or 'Args:' in docstring
        has_returns = ':return' in docstring or ':returns' in docstring or 'Returns:' in docstring
        has_raises = ':raises' in docstring or 'Raises:' in docstring
        
        if has_params and has_returns:
            return 'excellent'
        elif has_params or has_returns:
            return 'good'
        else:
            return 'basic'
    
    def analyze_file(file_path):
        """分析单个文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 使用AST解析，找出所有函数和类
            tree = ast.parse(content)
            
            for node in ast.walk(tree):
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                    # 获取docstring
                    if (node.body and 
                        isinstance(node.body[0], ast.Expr) and 
                        isinstance(node.body[0].value, ast.Constant)):
                        
                        docstring = node.body[0].value.value
                        score = get_docstring_score(docstring)
                        
                        docstring_stats[score] += 1
                        
                        # 收集问题
                        if score == 'basic':
                            docstring_stats['issues'].append({
                                'file': str(file_path),
                                'name': node.name,
                                'type': 'function' if isinstance(node, ast.FunctionDef) else 'class',
                                'score': score,
                                'first_line': docstring.split('\n')[0] if docstring else 'No docstring'
                            })
                    
                    elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        # 没有docstring的函数
                        docstring_stats['none'] += 1
                        
                        # 检查是否有明显的TODO注释
                        if any('TODO' in str(stmt) for stmt in node.body):
                            docstring_stats['issues'].append({
                                'file': str(file_path),
                                'name': node.name,
                                'type': 'function' if isinstance(node, ast.FunctionDef) else 'class',
                                'score': 'none',
                                'first_line': 'Missing docstring, has TODO'
                            })
                        
        except Exception as e:
            docstring_stats['issues'].append({
                'file': str(file_path),
                'name': 'Parse Error',
                'type': 'file',
                'score': 'error',
                'first_line': str(e)
            })
    
    # 遍历所有Python文件
    for py_file in excel_mcp_dir.rglob("*.py"):
        analyze_file(py_file)
    
    # 生成报告
    total = sum(docstring_stats[k] for k in ['excellent', 'good', 'basic', 'none'])
    
    print("=" * 60)
    print("DOCSTRING质量分析报告")
    print("=" * 60)
    print(f"总函数/类数量: {total}")
    print(f"优秀评分 (参数+返回值): {docstring_stats['excellent']} ({docstring_stats['excellent']/total*100:.1f}%)")
    print(f"良好评分 (参数或返回值): {docstring_stats['good']} ({docstring_stats['good']/total*100:.1f}%)")
    print(f"基础评分 (简单描述): {docstring_stats['basic']} ({docstring_stats['basic']/total*100:.1f}%)")
    print(f"无docstring: {docstring_stats['none']} ({docstring_stats['none']/total*100:.1f}%)")
    
    if docstring_stats['issues']:
        print("\n🔍 发现的问题:")
        for issue in docstring_stats['issues'][:10]:  # 只显示前10个
            print(f"  - {issue['file']}:{issue['name']} ({issue['type']}) - {issue['first_line']}")
    
    # 生成改进建议
    print("\n💡 改进建议:")
    if docstring_stats['excellent'] < total * 0.8:
        print("  • 建议优先为高频使用的函数添加完整的参数和返回值描述")
    if docstring_stats['none'] > 0:
        print("  • 建议为所有公共函数和类添加基础docstring")
    
    # 返回统计结果
    return docstring_stats

if __name__ == "__main__":
    analyze_docstring_quality()