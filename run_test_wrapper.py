#!/usr/bin/env python3
"""
运行 test_fixes.py 并捕获所有输出
"""
import sys
import subprocess
from pathlib import Path

# 设置工作目录
work_dir = Path('/root/.openclaw/workspace/excel-mcp-server')
sys.path.insert(0, str(work_dir / 'src'))

# 运行测试脚本
test_script = work_dir / 'test_fixes.py'

print("="*80)
print("正在运行测试脚本: test_fixes.py")
print("="*80)

result = subprocess.run(
    [sys.executable, str(test_script)],
    capture_output=True,
    text=True,
    cwd=str(work_dir)
)

print("\n" + "="*80)
print("STDOUT 输出:")
print("="*80)
print(result.stdout)

if result.stderr:
    print("\n" + "="*80)
    print("STDERR 输出:")
    print("="*80)
    print(result.stderr)

print("\n" + "="*80)
print(f"返回码: {result.returncode}")
print("="*80)

# 保存输出到文件
output_file = work_dir / 'test_fixes_output.log'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write("="*80 + "\n")
    f.write("STDOUT:\n")
    f.write("="*80 + "\n")
    f.write(result.stdout)
    if result.stderr:
        f.write("\n" + "="*80 + "\n")
        f.write("STDERR:\n")
        f.write("="*80 + "\n")
        f.write(result.stderr)
    f.write("\n" + "="*80 + "\n")
    f.write(f"Return Code: {result.returncode}\n")

print(f"\n完整输出已保存到: {output_file}")
