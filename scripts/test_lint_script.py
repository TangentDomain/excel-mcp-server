#!/usr/bin/env python3
"""Test script for lint_docstring_contract.py"""

import os
import subprocess
import tempfile
import sys
from pathlib import Path

def create_test_function_file():
    """创建一个有参数但没有docstring的测试函数"""
    test_content = '''
def test_function(param1, param2="default"):
    """Simple test function.
    
    Args:
        param1: First parameter
        param2: Second parameter with default
    
    Returns:
        str: Result string
    """
    return f"{param1}-{param2}"

def function_missing_args_docs(required_param, optional_param="hello"):
    """This function is missing proper Args section.
    
    Returns:
        str: Some result
    """
    return f"{required_param}-{optional_param}"
'''
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as f:
        f.write(test_content.strip())
        return f.name

def test_lint_script():
    """测试lint脚本"""
    test_file = create_test_function_file()
    
    try:
        # 运行lint脚本
        result = subprocess.run([
            sys.executable, 'scripts/lint_docstring_contract.py'
        ], capture_output=True, text=True, timeout=30)
        
        print(f"Exit code: {result.returncode}")
        print(f"STDOUT:\n{result.stdout}")
        if result.stderr:
            print(f"STDERR:\n{result.stderr}")
        
        # 清理
        os.unlink(test_file)
        
        return result.returncode != 0  # 应该有错误
        
    except subprocess.TimeoutExpired:
        print("Test timed out")
        os.unlink(test_file)
        return False

if __name__ == "__main__":
    success = test_lint_script()
    print(f"Test result: {'PASS' if success else 'FAIL'}")
    sys.exit(0 if success else 1)