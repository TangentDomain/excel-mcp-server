#!/usr/bin/env python3
"""
统计需要添加 docstring 的公共函数
扫描所有 .py 文件，统计缺失 Args/Parameters 和 Returns 段的公共函数数量
"""

import ast
import os
from pathlib import Path
from typing import Dict, List, Set, Tuple


def analyze_docstring(docstring: str) -> Tuple[bool, bool]:
    """
    分析 docstring 是否包含 Args/Parameters 和 Returns 段

    Args:
        docstring: 要分析的 docstring

    Returns:
        (has_args, has_returns) 元组
    """
    if not docstring:
        return False, False

    # 检查是否有 Args 或 Parameters 段
    has_args = False
    for keyword in ['Args:', 'Parameters:', 'Args\n', 'Parameters\n', 'Args ', 'Parameters ']:
        if keyword in docstring:
            has_args = True
            break

    # 检查是否有 Returns 段
    has_returns = False
    for keyword in ['Returns:', 'Returns\n', 'Returns ']:
        if keyword in docstring:
            has_returns = True
            break

    return has_args, has_returns


def analyze_file(file_path: Path) -> Dict:
    """
    分析单个 Python 文件

    Args:
        file_path: 文件路径

    Returns:
        包含分析结果的字典
    """
    result = {
        'file': str(file_path),
        'total_public_functions': 0,
        'missing_args': 0,
        'missing_returns': 0,
        'missing_both': 0,
        'functions_detail': []
    }

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            tree = ast.parse(content)

        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                # 跳过私有函数（以下划线开头）
                if node.name.startswith('_'):
                    continue

                result['total_public_functions'] += 1

                # 获取 docstring
                docstring = ast.get_docstring(node)
                has_args, has_returns = analyze_docstring(docstring or '')

                missing_args = not has_args
                missing_returns = not has_returns

                if missing_args:
                    result['missing_args'] += 1
                if missing_returns:
                    result['missing_returns'] += 1
                if missing_args and missing_returns:
                    result['missing_both'] += 1

                # 记录详细信息
                result['functions_detail'].append({
                    'name': node.name,
                    'line': node.lineno,
                    'has_docstring': bool(docstring),
                    'has_args': has_args,
                    'has_returns': has_returns,
                    'missing_args': missing_args,
                    'missing_returns': missing_returns
                })

    except Exception as e:
        print(f"Error analyzing {file_path}: {e}")

    return result


def scan_directory(directory: Path, exclude_dirs: Set[str] = None) -> List[Dict]:
    """
    扫描目录下的所有 Python 文件

    Args:
        directory: 要扫描的目录
        exclude_dirs: 要排除的目录名集合

    Returns:
        所有文件的分析结果列表
    """
    if exclude_dirs is None:
        exclude_dirs = {'.git', '__pycache__', '.venv', 'venv', 'env', 'build', 'dist', 'node_modules', '.pytest_cache', 'tests', 'test', 'scripts', 'temp', 'docs', 'examples'}

    results = []

    for root, dirs, files in os.walk(directory):
        # 过滤掉排除的目录
        dirs[:] = [d for d in dirs if d not in exclude_dirs and not d.startswith('.')]

        for file in files:
            if file.endswith('.py'):
                file_path = Path(root) / file
                result = analyze_file(file_path)
                if result['total_public_functions'] > 0:
                    results.append(result)

    return results


def main():
    """主函数"""
    # 获取项目根目录
    project_root = Path('/root/.openclaw/workspace/excel-mcp-server')
    src_dir = project_root / 'src'

    print("=" * 80)
    print("统计需要添加 docstring 的公共函数")
    print("=" * 80)
    print()

    # 扫描 src 目录
    print(f"扫描目录: {src_dir}")
    print()

    results = scan_directory(src_dir)

    # 汇总统计
    total_files = len(results)
    total_functions = sum(r['total_public_functions'] for r in results)
    total_missing_args = sum(r['missing_args'] for r in results)
    total_missing_returns = sum(r['missing_returns'] for r in results)
    total_missing_both = sum(r['missing_both'] for r in results)

    print("汇总统计:")
    print("-" * 80)
    print(f"总文件数: {total_files}")
    print(f"公共函数总数: {total_functions}")
    print(f"缺少 Args/Parameters 段的函数数: {total_missing_args}")
    print(f"缺少 Returns 段的函数数: {total_missing_returns}")
    print(f"同时缺少两者的函数数: {total_missing_both}")
    print()

    # 按文件分组显示
    print("按文件分组:")
    print("-" * 80)

    for result in sorted(results, key=lambda x: x['file']):
        print(f"\n文件: {result['file']}")
        print(f"  公共函数数: {result['total_public_functions']}")
        print(f"  缺少 Args: {result['missing_args']}")
        print(f"  缺少 Returns: {result['missing_returns']}")
        print(f"  缺少两者: {result['missing_both']}")

        # 显示有问题的函数
        problematic_funcs = [f for f in result['functions_detail']
                           if f['missing_args'] or f['missing_returns']]
        if problematic_funcs:
            print(f"  需要改进的函数:")
            for func in problematic_funcs:
                issues = []
                if func['missing_args']:
                    issues.append('Args')
                if func['missing_returns']:
                    issues.append('Returns')
                print(f"    - {func['name']} (行 {func['line']}) 缺少: {', '.join(issues)}")

    print()
    print("=" * 80)
    print("统计完成")
    print("=" * 80)


if __name__ == '__main__':
    main()
