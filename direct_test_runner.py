#!/usr/bin/env python3
"""Direct test runner to check if tests work"""
import sys
import os
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server/src')
sys.path.insert(0, '/root/.openclaw/workspace/excel-mcp-server')

os.chdir('/root/.openclaw/workspace/excel-mcp-server')

# Try to import the test module
try:
    from tests.test_group_concat import TestGroupConcat
    print("✓ Successfully imported test_group_concat.TestGroupConcat")
except Exception as e:
    print(f"✗ Failed to import test_group_concat: {e}")
    import traceback
    traceback.print_exc()

try:
    from tests.test_join_types import TestRightJoin
    print("✓ Successfully imported test_join_types.TestRightJoin")
except Exception as e:
    print(f"✗ Failed to import test_join_types: {e}")
    import traceback
    traceback.print_exc()

# Now try to run pytest
print("\n" + "="*70)
print("Running pytest on GROUP_CONCAT tests...")
print("="*70)
import subprocess
result = subprocess.run(
    ["python3", "-m", "pytest", "tests/test_group_concat.py", "-v"],
    cwd='/root/.openclaw/workspace/excel-mcp-server',
    capture_output=True,
    text=True
)
print(result.stdout)
if result.stderr:
    print("STDERR:", result.stderr)
print(f"\nExit code: {result.returncode}")

print("\n" + "="*70)
print("Running pytest on Right JOIN test...")
print("="*70)
result2 = subprocess.run(
    ["python3", "-m", "pytest", "tests/test_join_types.py::TestRightJoin::test_basic_right_join", "-v"],
    cwd='/root/.openclaw/workspace/excel-mcp-server',
    capture_output=True,
    text=True
)
print(result2.stdout)
if result2.stderr:
    print("STDERR:", result2.stderr)
print(f"\nExit code: {result2.returncode}")

print("\n" + "="*70)
print("SUMMARY")
print("="*70)
print(f"GROUP_CONCAT tests: {'PASSED ✓' if result.returncode == 0 else 'FAILED ✗'}")
print(f"Right JOIN test: {'PASSED ✓' if result2.returncode == 0 else 'FAILED ✗'}")
