#!/usr/bin/env python3
"""
Docstring 合规性检查脚本

检查所有公共函数的 docstring 是否包含 Args/Parameters 和 Returns 段
"""

import ast
import os
import sys
from pathlib import Path
from typing import List, Dict, Tuple


class DocstringChecker(ast.NodeVisitor):
    """Docstring 检查器"""

    def __init__(self):
        self.total_functions = 0
        self.compliant_functions = 0
        self.non_compliant_functions = []

    def visit_FunctionDef(self, node: ast.FunctionDef):
        """访问函数定义节点"""
        # 跳过私有函数（以下划线开头）
        if node.name.startswith('_'):
            return

        self.total_functions += 1

        # 获取 docstring
        docstring = ast.get_docstring(node)
        if not docstring:
            self.non_compliant_functions.append({
                'name': node.name,
                'line': node.lineno,
                'issue': '缺少 docstring'
            })
            return

        # 检查是否包含 Args 或 Parameters
        has_args = 'Args:' in docstring or 'Parameters:' in docstring

        # 检查是否包含 Returns
        has_returns = 'Returns:' in docstring

        if not (has_args and has_returns):
            issues = []
            if not has_args:
                issues.append('缺少 Args 或 Parameters 段')
            if not has_returns:
                issues.append('缺少 Returns 段')

            self.non_compliant_functions.append({
                'name': node.name,
                'line': node.lineno,
                'issue': ', '.join(issues)
            })
        else:
            self.compliant_functions += 1

        # 继续访问子节点
        self.generic_visit(node)


def check_file(file_path: Path) -> Dict:
    """检查单个 Python 文件

    Args:
        file_path: Python 文件路径

    Returns:
        检查结果字典
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 解析 AST
        tree = ast.parse(content, filename=str(file_path))

        # 创建检查器并访问
        checker = DocstringChecker()
        checker.visit(tree)

        return {
            'file_path': str(file_path),
            'total_functions': checker.total_functions,
            'compliant_functions': checker.compliant_functions,
            'non_compliant_functions': checker.non_compliant_functions
        }
    except Exception as e:
        return {
            'file_path': str(file_path),
            'error': str(e),
            'total_functions': 0,
            'compliant_functions': 0,
            'non_compliant_functions': []
        }


def scan_directory(directory: Path) -> List[Dict]:
    """扫描目录下的所有 Python 文件

    Args:
        directory: 要扫描的目录路径

    Returns:
        所有文件的检查结果列表
    """
    results = []

    # 遍历目录
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                file_path = Path(root) / file
                result = check_file(file_path)
                results.append(result)

    return results


def print_results(results: List[Dict]):
    """打印检查结果

    Args:
        results: 检查结果列表
    """
    total_functions = 0
    total_compliant = 0
    total_non_compliant = 0
    all_issues = []

    print("=" * 80)
    print("Docstring 合规性检查报告")
    print("=" * 80)
    print()

    # 统计总体数据
    for result in results:
        if 'error' in result:
            print(f"⚠️  文件: {result['file_path']}")
            print(f"   错误: {result['error']}")
            print()
            continue

        total_functions += result['total_functions']
        total_compliant += result['compliant_functions']
        total_non_compliant += len(result['non_compliant_functions'])

        # 收集所有问题
        for issue in result['non_compliant_functions']:
            all_issues.append({
                'file': result['file_path'],
                'name': issue['name'],
                'line': issue['line'],
                'issue': issue['issue']
            })

    # 打印汇总统计
    print("📊 总体统计:")
    print(f"   扫描文件数: {len(results)}")
    print(f"   总函数数: {total_functions}")
    print(f"   合规函数数: {total_compliant}")
    print(f"   不合规函数数: {total_non_compliant}")

    if total_functions > 0:
        compliance_rate = (total_compliant / total_functions) * 100
        print(f"   合规率: {compliance_rate:.2f}%")
    print()

    # 打印不合规函数清单
    if all_issues:
        print("❌ 不合规函数清单:")
        print("-" * 80)

        # 按文件分组
        issues_by_file = {}
        for issue in all_issues:
            file_path = issue['file']
            if file_path not in issues_by_file:
                issues_by_file[file_path] = []
            issues_by_file[file_path].append(issue)

        for file_path, issues in sorted(issues_by_file.items()):
            # 显示相对路径
            rel_path = str(Path(file_path).relative_to(Path.cwd()))
            print(f"\n📄 {rel_path}")
            for issue in sorted(issues, key=lambda x: x['line']):
                print(f"   行 {issue['line']:<4} | {issue['name']:<30} | {issue['issue']}")
    else:
        print("✅ 所有函数都符合 docstring 规范！")

    print()
    print("=" * 80)


def main():
    """主函数"""
    # 获取要扫描的目录
    script_dir = Path(__file__).parent
    project_dir = script_dir.parent
    target_dir = project_dir / "src" / "excel_mcp_server_fastmcp"

    if not target_dir.exists():
        print(f"❌ 错误: 目标目录不存在: {target_dir}")
        sys.exit(1)

    # 扫描目录
    print(f"🔍 正在扫描目录: {target_dir}")
    print()

    results = scan_directory(target_dir)

    # 打印结果
    print_results(results)


if __name__ == "__main__":
    main()
