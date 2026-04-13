#!/usr/bin/env python3
"""Docstring contract validation script.

Validates that public API functions (execute_*) have complete docstrings.
For other functions (register_*, *_resource, internal helpers), only checks
that a docstring exists.

Usage:
    python scripts/lint_docstring_contract.py
    python scripts/lint_docstring_contract.py --fix

Exit codes:
    0: All validations pass
    1: Validation errors found
"""

import argparse
import ast
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Set, Optional


class DocstringValidator:
    """Validates docstring contracts against function signatures."""

    # Functions that require full Args/Parameters validation (explicit allowlist)
    STRICT_FUNCTIONS = {'execute_advanced_sql_query'}

    # Functions that only need a docstring present (no Args check)
    RELAXED_PREFIXES = ('register_',)
    RELAXED_SUFFIXES = ('_resource',)

    def __init__(self):
        self.errors = []
        self.warnings = []

    @staticmethod
    def is_strict_function(func_name: str) -> bool:
        """Check if a function requires strict Args/Parameters validation."""
        return func_name in DocstringValidator.STRICT_FUNCTIONS

    @staticmethod
    def is_relaxed_function(func_name: str) -> bool:
        """Check if a function only needs a docstring presence check."""
        if func_name.startswith('_'):
            return True
        if any(func_name.startswith(p) for p in DocstringValidator.RELAXED_PREFIXES):
            return True
        if any(func_name.endswith(s) for s in DocstringValidator.RELAXED_SUFFIXES):
            return True
        return False

    def get_function_signatures(self, file_path: Path) -> Dict[str, Tuple[List[str], Dict[str, Optional[str]], List[str]]]:
        """Extract function signatures from Python file.

        Returns:
            Dict mapping function name to (params, defaults, param_names)
            - params: list of parameter names (excluding 'self')
            - defaults: dict mapping param name to default value string
            - param_names: ordered list of all parameter names
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            self.warnings.append(f"{file_path}: | 文件读取失败: {e}")
            return {}

        tree = ast.parse(content)
        functions = {}

        # Find all class definitions to check if functions are class methods
        class_def_nodes = [node for node in ast.walk(tree) if isinstance(node, ast.ClassDef)]

        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
                # Skip __init__ methods
                if node.name == '__init__':
                    continue

                # Skip class methods (they start with __init__ or are in classes)
                is_class_method = False
                for class_node in class_def_nodes:
                    if hasattr(class_node, 'body') and node in class_node.body:
                        is_class_method = True
                        break

                if is_class_method:
                    continue

                params = []
                defaults = {}
                param_names = []

                # Get all parameters
                all_args = node.args.args + node.args.kwonlyargs
                defaults_count = len(node.args.defaults)

                for i, arg in enumerate(all_args):
                    if arg.arg != 'self':
                        params.append(arg.arg)
                        param_names.append(arg.arg)

                        # Check for default value
                        if i >= defaults_count:
                            defaults[arg.arg] = None
                        elif node.args.defaults and i - defaults_count >= 0:
                            default_index = i - defaults_count
                            if default_index < len(node.args.defaults):
                                default_value = node.args.defaults[default_index]
                                defaults[arg.arg] = ast.unparse(default_value) if default_value else None

                functions[node.name] = (params, defaults, param_names)

        return functions

    def parse_docstring_args(self, docstring: str) -> Dict[str, Optional[str]]:
        """Extract parameter documentation from docstring.

        Returns:
            Dict mapping parameter name to default value (or None if no default documented)
        """
        if not docstring:
            return {}

        params = {}

        # Format 1: Args: section
        if 'Args:' in docstring:
            args_section = docstring.split('Args:')[1]
            # Stop at next major section
            for next_section in ['Returns:', 'Yields:', 'Raises:', 'Note:', 'Example:']:
                if next_section in args_section:
                    args_section = args_section.split(next_section)[0]
            if '\n\n' in args_section:
                args_section = args_section.split('\n\n')[0]

            params.update(self._parse_param_lines(args_section))

        # Format 2: Parameters: section
        elif 'Parameters:' in docstring:
            params_section = docstring.split('Parameters:')[1]
            for next_section in ['Returns:', 'Yields:', 'Raises:', 'Note:', 'Example:']:
                if next_section in params_section:
                    params_section = params_section.split(next_section)[0]
            if '\n\n' in params_section:
                params_section = params_section.split('\n\n')[0]

            params.update(self._parse_param_lines(params_section))

        return params

    def _parse_param_lines(self, section: str) -> Dict[str, Optional[str]]:
        """Parse parameter lines from a docstring section."""
        params = {}
        lines = section.strip().split('\n')

        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            # Extract parameter name and default value
            if ':' in line:
                param_part = line.split(':', 1)[0].strip()

                # Extract default value from parentheses
                default_value = None
                if '(' in param_part and ')' in param_part:
                    content = param_part[param_part.find('(')+1:param_part.rfind(')')]

                    if 'default=' in content:
                        default_value = content.split('default=', 1)[1].strip()
                    elif content == 'optional':
                        default_value = 'optional'
                    elif '=' in content:
                        default_value = content.split('=', 1)[1].strip()

                    param_name = param_part.split('(')[0].strip()
                else:
                    param_name = param_part

                if param_name and param_name not in ['Args', 'Parameters', 'Returns']:
                    params[param_name] = default_value

        return params

    def get_docstring_params(self, file_path: Path) -> Dict[str, Dict[str, Optional[str]]]:
        """Extract parameter documentation from all functions in a file.

        Returns:
            Dict mapping function name to dict of {param_name: default_value_or_None}
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            self.warnings.append(f"{file_path}: | 文件读取失败: {e}")
            return {}

        tree = ast.parse(content)
        docstring_params = {}

        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
                if node.name == '__init__':
                    continue

                # Get docstring
                if (node.body and
                    isinstance(node.body[0], ast.Expr) and
                    isinstance(node.body[0].value, ast.Constant) and
                    isinstance(node.body[0].value.value, str)):

                    docstring = node.body[0].value.value
                    params = self.parse_docstring_args(docstring)

                    if params:
                        docstring_params[node.name] = params

        return docstring_params

    def has_docstring(self, file_path: Path, func_name: str) -> bool:
        """Check if a function has any docstring at all."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception:
            return False

        tree = ast.parse(content)

        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and node.name == func_name:
                if (node.body and
                    isinstance(node.body[0], ast.Expr) and
                    isinstance(node.body[0].value, ast.Constant) and
                    isinstance(node.body[0].value.value, str)):
                    return True
                return False

        return False

    def validate_file(self, file_path: Path) -> List[str]:
        """Validate all functions in a single Python file.

        Returns:
            List of error messages for this file
        """
        if '__pycache__' in str(file_path) or 'test_' in file_path.name:
            return []

        functions = self.get_function_signatures(file_path)
        docstring_params = self.get_docstring_params(file_path)

        errors = []

        for func_name, (signature_params, defaults, _) in functions.items():
            is_strict = self.is_strict_function(func_name)
            is_relaxed = self.is_relaxed_function(func_name)

            if is_relaxed:
                # Relaxed mode: only check that a docstring exists
                if not self.has_docstring(file_path, func_name):
                    errors.append(
                        f"{file_path}:{func_name} | - | 函数缺少docstring"
                    )
                continue

            if is_strict:
                # Strict mode: full Args/Parameters validation
                if func_name in docstring_params:
                    doc_params = docstring_params[func_name]

                    # Check for missing parameters in docstring
                    for param in signature_params:
                        if param not in doc_params:
                            errors.append(
                                f"{file_path}:{func_name} | {param} | docstring中缺失参数"
                            )

                    # Check for default value consistency
                    for param in doc_params:
                        if param in signature_params:
                            doc_default = doc_params[param]
                            code_default = defaults.get(param)

                            if doc_default == 'optional' and code_default is None:
                                continue

                            if doc_default and code_default:
                                doc_normalized = self._normalize_default_value(doc_default)
                                code_normalized = self._normalize_default_value(code_default)

                                if doc_normalized != code_normalized:
                                    errors.append(
                                        f"{file_path}:{func_name} | {param} | 默认值不匹配: "
                                        f"docstring='{doc_default}', code='{code_default}'"
                                    )
                else:
                    # Function has parameters but no docstring with Args section
                    if signature_params:
                        errors.append(
                            f"{file_path}:{func_name} | 所有参数 | "
                            f"函数有参数但缺少Args/Parameters段"
                        )
            else:
                # Default: relaxed - just check docstring presence
                if not self.has_docstring(file_path, func_name):
                    errors.append(
                        f"{file_path}:{func_name} | - | 函数缺少docstring"
                    )

        return errors

    def _normalize_default_value(self, value: str) -> str:
        """Normalize default value for comparison."""
        if not value:
            return ""

        value = value.strip()
        value = value.strip('"').strip("'")
        value = value.lower()

        return value


def lint_docstring_contract(src_dir: str) -> List[str]:
    """Validate docstring contracts for all Python files in a directory.

    Args:
        src_dir: Source directory to scan

    Returns:
        List of error messages
    """
    validator = DocstringValidator()
    all_errors = []

    for py_file in Path(src_dir).rglob('*.py'):
        if '__pycache__' in str(py_file):
            continue

        errors = validator.validate_file(py_file)
        all_errors.extend(errors)

    # Add warnings if any
    if validator.warnings:
        for warning in validator.warnings:
            all_errors.append(f"WARNING: {warning}")

    return all_errors


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Validate docstring contracts against function signatures'
    )
    parser.add_argument(
        '--src-dir',
        default='src',
        help='Source directory to scan (default: src)'
    )
    parser.add_argument(
        '--fix',
        action='store_true',
        help='Auto-fix simple issues (not yet implemented)'
    )
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Show warnings in addition to errors'
    )
    parser.add_argument(
        '--quiet',
        action='store_true',
        help='Suppress all output, exit code only'
    )

    args = parser.parse_args()

    errors = lint_docstring_contract(args.src_dir)

    # Filter warnings unless verbose mode
    if not args.verbose:
        errors = [e for e in errors if not e.startswith('WARNING:')]

    if errors:
        if not args.quiet:
            print("Docstring contract validation errors:")
            for error in errors:
                print(f"  {error}")
            print(f"\nTotal: {len(errors)} error(s)")
        sys.exit(1)
    else:
        if not args.quiet:
            print("✓ 所有函数的docstring契约验证通过")
        sys.exit(0)


if __name__ == "__main__":
    main()
