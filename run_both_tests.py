#!/usr/bin/env python3
import subprocess
import sys

print("=" * 80)
print("执行测试脚本验证修复效果")
print("=" * 80)

# Test 1: test_same_file_join.py
print("\n" + "=" * 80)
print("测试 1: test_same_file_join.py")
print("=" * 80)
result1 = subprocess.run(
    [sys.executable, "/root/.openclaw/workspace/excel-mcp-server/test_same_file_join.py"],
    capture_output=True,
    text=True,
    cwd="/root/.openclaw/workspace/excel-mcp-server"
)
print(result1.stdout)
if result1.stderr:
    print("STDERR:", result1.stderr)
print(f"退出码: {result1.returncode}")

# Test 2: test_group_concat_complex.py
print("\n" + "=" * 80)
print("测试 2: test_group_concat_complex.py")
print("=" * 80)
result2 = subprocess.run(
    [sys.executable, "/root/.openclaw/workspace/excel-mcp-server/test_group_concat_complex.py"],
    capture_output=True,
    text=True,
    cwd="/root/.openclaw/workspace/excel-mcp-server"
)
print(result2.stdout)
if result2.stderr:
    print("STDERR:", result2.stderr)
print(f"退出码: {result2.returncode}")

# Summary
print("\n" + "=" * 80)
print("测试总结")
print("=" * 80)
print(f"test_same_file_join.py: {'✅ 通过' if result1.returncode == 0 else '❌ 失败'}")
print(f"test_group_concat_complex.py: {'✅ 通过' if result2.returncode == 0 else '❌ 失败'}")
