#!/usr/bin/env python3
"""
统计缺失 docstring 的公共函数
"""
import ast
import os
from pathlib import Path
from typing import List, Dict, Set

def has_public_decorator(node: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    """检查是否有 @mcp.tool 或其他公共 API 装饰器"""
    for decorator in node.decorator_list:
        if isinstance(decorator, ast.Attribute):
            if isinstance(decorator.value, ast.Name) and decorator.value.id == 'mcp' and decorator.attr == 'tool':
                return True
        elif isinstance(decorator, ast.Call):
            if isinstance(decorator.func, ast.Attribute):
                if isinstance(decorator.func.value, ast.Name) and decorator.func.value.id == 'mcp' and decorator.func.attr == 'tool':
                    return True
    return False

def get_function_docstring(node: ast.FunctionDef | ast.AsyncFunctionDef) -> str:
    """获取函数的 docstring"""
    if (node.body and isinstance(node.body[0], ast.Expr) and
        isinstance(node.body[0].value, ast.Constant) and
        isinstance(node.body[0].value.value, str)):
        return node.body[0].value.value
    return ""

def has_parameter_section(docstring: str) -> bool:
    """检查 docstring 是否包含 Args 或 Parameters 段"""
    lines = docstring.strip().split('\n')
    for line in lines:
        line = line.strip()
        if line.startswith('Args:') or line.startswith('Parameters:'):
            return True
    return False

def has_returns_section(docstring: str) -> bool:
    """检查 docstring 是否包含 Returns 段"""
    lines = docstring.strip().split('\n')
    for line in lines:
        line = line.strip()
        if line.startswith('Returns:') or line.startswith('Return:'):
            return True
    return False

def count_parameters(node: ast.FunctionDef | ast.AsyncFunctionDef) -> int:
    """统计函数的参数数量（不包括 self 和 cls）"""
    params = []
    for arg in node.args.args:
        if arg.arg not in ('self', 'cls'):
            params.append(arg)
    return len(params)

def analyze_file(filepath: Path) -> Dict:
    """分析单个 Python 文件"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    try:
        tree = ast.parse(content)
    except SyntaxError:
        return {
            'filepath': filepath,
            'error': 'SyntaxError',
            'functions': [],
            'methods': []
        }

    result = {
        'filepath': filepath,
        'functions': [],
        'methods': []
    }

    class Visitor(ast.NodeVisitor):
        def __init__(self):
            self.current_class = None

        def visit_ClassDef(self, node):
            self.current_class = node.name
            self.generic_visit(node)
            self.current_class = None

        def visit_FunctionDef(self, node):
            self._process_function(node)
            self.generic_visit(node)

        def visit_AsyncFunctionDef(self, node):
            self._process_function(node)
            self.generic_visit(node)

        def _process_function(self, node: ast.FunctionDef | ast.AsyncFunctionDef):
            # 跳过私有函数（下划线开头）
            if node.name.startswith('_'):
                return

            # 跳过 __init__ 方法
            if node.name == '__init__':
                return

            docstring = get_function_docstring(node)
            param_count = count_parameters(node)
            has_params = has_parameter_section(docstring)
            has_returns = has_returns_section(docstring)

            # 检查是否缺少必要的 docstring 段
            missing_params = param_count > 0 and not has_params
            missing_returns = (node.returns is not None or "return" in docstring.lower()) and not has_returns

            func_info = {
                'name': node.name,
                'line': node.lineno,
                'param_count': param_count,
                'has_docstring': bool(docstring),
                'has_params_section': has_params,
                'has_returns_section': has_returns,
                'missing_params': missing_params,
                'missing_returns': missing_returns,
                'has_public_decorator': has_public_decorator(node)
            }

            if self.current_class:
                result['methods'].append(func_info)
            else:
                result['functions'].append(func_info)

    visitor = Visitor()
    visitor.visit(tree)

    return result

def main():
    base_dir = Path('src/excel_mcp_server_fastmcp')
    python_files = list(base_dir.rglob('*.py'))

    all_results = []
    missing_docstring_functions = []
    missing_docstring_methods = []

    for filepath in python_files:
        result = analyze_file(filepath)
        all_results.append(result)

        # 收集缺失 docstring 的函数
        for func in result['functions']:
            if func['missing_params'] or func['missing_returns']:
                missing_docstring_functions.append({
                    'file': str(filepath.relative_to(base_dir)),
                    **func
                })

        for method in result['methods']:
            if method['missing_params'] or method['missing_returns']:
                missing_docstring_methods.append({
                    'file': str(filepath.relative_to(base_dir)),
                    **method
                })

    # 输出统计结果
    print("=" * 80)
    print("缺失 docstring 段的公共函数统计")
    print("=" * 80)
    print(f"\n总文件数: {len(python_files)}")

    total_functions = sum(len(r['functions']) for r in all_results)
    total_methods = sum(len(r['methods']) for r in all_results)
    print(f"总函数数: {total_functions}")
    print(f"总方法数: {total_methods}")

    missing_func_count = len(missing_docstring_functions)
    missing_method_count = len(missing_docstring_methods)
    print(f"\n缺失 Args/Parameters 的函数: {sum(1 for f in missing_docstring_functions if f['missing_params'])}")
    print(f"缺失 Returns 的函数: {sum(1 for f in missing_docstring_functions if f['missing_returns'])}")
    print(f"缺失 Args/Parameters 的方法: {sum(1 for m in missing_docstring_methods if m['missing_params'])}")
    print(f"缺失 Returns 的方法: {sum(1 for m in missing_docstring_methods if m['missing_returns'])}")

    print("\n" + "=" * 80)
    print("缺失 docstring 段的函数清单")
    print("=" * 80)

    if missing_docstring_functions:
        print("\n### 模块函数\n")
        for func in missing_docstring_functions:
            print(f"\n文件: {func['file']}")
            print(f"  函数: {func['name']} (第 {func['line']} 行)")
            print(f"  参数数: {func['param_count']}")
            if func['missing_params']:
                print(f"  ❌ 缺失 Args/Parameters 段")
            if func['missing_returns']:
                print(f"  ❌ 缺失 Returns 段")
            if func['has_public_decorator']:
                print(f"  ℹ️  有 @mcp.tool 装饰器")

    if missing_docstring_methods:
        print("\n### 类方法\n")
        for method in missing_docstring_methods:
            print(f"\n文件: {method['file']}")
            print(f"  方法: {method['name']} (第 {method['line']} 行)")
            print(f"  参数数: {method['param_count']}")
            if method['missing_params']:
                print(f"  ❌ 缺失 Args/Parameters 段")
            if method['missing_returns']:
                print(f"  ❌ 缺失 Returns 段")

    # 保存结果到文件
    output_file = Path('missing_docstring_report.txt')
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("缺失 docstring 段的公共函数统计\n")
        f.write("=" * 80 + "\n")
        f.write(f"\n总文件数: {len(python_files)}\n")
        f.write(f"总函数数: {total_functions}\n")
        f.write(f"总方法数: {total_methods}\n")
        f.write(f"\n缺失 Args/Parameters 的函数: {sum(1 for f in missing_docstring_functions if f['missing_params'])}\n")
        f.write(f"缺失 Returns 的函数: {sum(1 for f in missing_docstring_functions if f['missing_returns'])}\n")
        f.write(f"缺失 Args/Parameters 的方法: {sum(1 for m in missing_docstring_methods if m['missing_params'])}\n")
        f.write(f"缺失 Returns 的方法: {sum(1 for m in missing_docstring_methods if m['missing_returns'])}\n")

        f.write("\n" + "=" * 80 + "\n")
        f.write("缺失 docstring 段的函数清单\n")
        f.write("=" * 80 + "\n")

        if missing_docstring_functions:
            f.write("\n### 模块函数\n\n")
            for func in missing_docstring_functions:
                f.write(f"文件: {func['file']}\n")
                f.write(f"  函数: {func['name']} (第 {func['line']} 行)\n")
                f.write(f"  参数数: {func['param_count']}\n")
                if func['missing_params']:
                    f.write(f"  缺失 Args/Parameters 段\n")
                if func['missing_returns']:
                    f.write(f"  缺失 Returns 段\n")
                f.write("\n")

        if missing_docstring_methods:
            f.write("\n### 类方法\n\n")
            for method in missing_docstring_methods:
                f.write(f"文件: {method['file']}\n")
                f.write(f"  方法: {method['name']} (第 {method['line']} 行)\n")
                f.write(f"  参数数: {method['param_count']}\n")
                if method['missing_params']:
                    f.write(f"  缺失 Args/Parameters 段\n")
                if method['missing_returns']:
                    f.write(f"  缺失 Returns 段\n")
                f.write("\n")

    print(f"\n详细报告已保存至: {output_file}")

if __name__ == '__main__':
    main()
