#!/usr/bin/env python3
"""Execute quick test and report results"""
import subprocess
import sys

print("Running quick test to verify fixes...")
print("="*60)

result = subprocess.run(
    [sys.executable, "quick_test.py"],
    cwd="/root/.openclaw/workspace/excel-mcp-server",
    capture_output=True,
    text=True
)

print(result.stdout)
if result.stderr:
    print("STDERR:", result.stderr)

print("\n" + "="*60)
print(f"Exit code: {result.returncode}")
print("="*60)

if result.returncode == 0:
    print("\n✅ All quick tests PASSED!")
else:
    print("\n❌ Some tests FAILED!")

sys.exit(result.returncode)
