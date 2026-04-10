#!/usr/bin/env python3
"""
运行修复验证测试
"""
import subprocess
import sys

def run_test(script_name):
    """运行测试脚本并返回结果"""
    print(f"\n{'='*70}")
    print(f"运行测试: {script_name}")
    print('='*70)
    result = subprocess.run(
        [sys.executable, script_name],
        capture_output=True,
        text=True
    )
    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
    return result.returncode == 0

if __name__ == '__main__':
    import os
    os.chdir('/root/.openclaw/workspace/excel-mcp-server')

    print("\n" + "="*70)
    print("ExcelMCP 修复验证测试")
    print("="*70)

    test1_passed = run_test('test_same_file_join.py')
    test2_passed = run_test('test_group_concat_complex.py')

    print("\n" + "="*70)
    print("测试总结")
    print("="*70)
    print(f"test_same_file_join.py (P0): {'✅ 通过' if test1_passed else '❌ 失败'}")
    print(f"test_group_concat_complex.py (P1): {'✅ 通过' if test2_passed else '❌ 失败'}")

    if test1_passed and test2_passed:
        print("\n✅ 所有测试通过！修复成功！")
        sys.exit(0)
    else:
        print("\n❌ 部分测试失败，需要检查")
        sys.exit(1)
