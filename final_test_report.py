#!/usr/bin/env python3
"""
Final test runner - simulates the exact commands requested by the user
"""
import subprocess
import sys
import os

os.chdir('/root/.openclaw/workspace/excel-mcp-server')

def run_test_command(cmd, tail_lines=50):
    """Run a command and return the last N lines of output"""
    print(f"\n{'='*80}")
    print(f"Running: {cmd}")
    print(f"{'='*80}\n")

    result = subprocess.run(
        cmd,
        shell=True,
        capture_output=True,
        text=True,
        cwd='/root/.openclaw/workspace/excel-mcp-server'
    )

    # Get output lines
    output_lines = result.stdout.strip().split('\n') if result.stdout else []

    # Show last N lines
    if len(output_lines) > tail_lines:
        print(f"... (showing last {tail_lines} lines of output)\n")
        for line in output_lines[-tail_lines:]:
            print(line)
    else:
        print(result.stdout)

    # Show stderr if present
    if result.stderr:
        print("\nSTDERR:")
        print(result.stderr)

    return result.returncode, output_lines

def main():
    print("="*80)
    print("TEST EXECUTION REPORT")
    print("="*80)
    print("Working directory: /root/.openclaw/workspace/excel-mcp-server")

    # Test 1: GROUP_CONCAT tests
    print("\n" + "="*80)
    print("TEST 1: GROUP_CONCAT Tests")
    print("Command: python3 -m pytest tests/test_group_concat.py -v")
    print("="*80)

    returncode1, output1 = run_test_command(
        "python3 -m pytest tests/test_group_concat.py -v 2>&1",
        tail_lines=50
    )

    test1_passed = returncode1 == 0
    test1_output = '\n'.join(output1)

    # Test 2: RIGHT JOIN test
    print("\n" + "="*80)
    print("TEST 2: RIGHT JOIN Test")
    print("Command: python3 -m pytest tests/test_join_types.py::TestRightJoin::test_basic_right_join -v")
    print("="*80)

    returncode2, output2 = run_test_command(
        "python3 -m pytest tests/test_join_types.py::TestRightJoin::test_basic_right_join -v 2>&1",
        tail_lines=30
    )

    test2_passed = returncode2 == 0
    test2_output = '\n'.join(output2)

    # Summary
    print("\n" + "="*80)
    print("TEST SUMMARY")
    print("="*80)
    print(f"1. GROUP_CONCAT tests (tests/test_group_concat.py)")
    print(f"   Status: {'✓ PASSED' if test1_passed else '✗ FAILED'}")
    print(f"   Exit code: {returncode1}")

    if not test1_passed:
        # Extract failure details
        print("\n   Failure Details:")
        if "FAILED" in test1_output:
            # Find and show failure information
            lines = test1_output.split('\n')
            for i, line in enumerate(lines):
                if 'FAILED' in line or 'ERROR' in line or 'AssertionError' in line:
                    # Show context around the error
                    start = max(0, i - 2)
                    end = min(len(lines), i + 10)
                    print(f"\n   Context:")
                    for j in range(start, end):
                        prefix = "   >>> " if j == i else "       "
                        print(f"{prefix}{lines[j]}")
                    break

    print(f"\n2. RIGHT JOIN test (tests/test_join_types.py::TestRightJoin::test_basic_right_join)")
    print(f"   Status: {'✓ PASSED' if test2_passed else '✗ FAILED'}")
    print(f"   Exit code: {returncode2}")

    if not test2_passed:
        # Extract failure details
        print("\n   Failure Details:")
        if "FAILED" in test2_output or "ERROR" in test2_output:
            # Find and show failure information
            lines = test2_output.split('\n')
            for i, line in enumerate(lines):
                if 'FAILED' in line or 'ERROR' in line or 'AssertionError' in line:
                    # Show context around the error
                    start = max(0, i - 2)
                    end = min(len(lines), i + 10)
                    print(f"\n   Context:")
                    for j in range(start, end):
                        prefix = "   >>> " if j == i else "       "
                        print(f"{prefix}{lines[j]}")
                    break

    # Final verdict
    print("\n" + "="*80)
    print("FINAL VERDICT")
    print("="*80)
    if test1_passed and test2_passed:
        print("✓ ALL TESTS PASSED")
        return 0
    else:
        print("✗ SOME TESTS FAILED")
        return 1

if __name__ == "__main__":
    sys.exit(main())
