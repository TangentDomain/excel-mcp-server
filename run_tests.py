#!/usr/bin/env python3
"""Run specific tests and report results"""
import subprocess
import sys

def run_command(cmd, description):
    """Run a command and return output"""
    print(f"\n{'='*70}")
    print(f"Running: {description}")
    print(f"Command: {cmd}")
    print(f"{'='*70}\n")

    result = subprocess.run(
        cmd,
        shell=True,
        capture_output=True,
        text=True,
        cwd='/root/.openclaw/workspace/excel-mcp-server'
    )

    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)

    return result.returncode, result.stdout, result.stderr

def main():
    # Test 1: GROUP_CONCAT tests
    print("\n" + "="*70)
    print("TEST 1: GROUP_CONCAT Tests")
    print("="*70)
    returncode1, stdout1, stderr1 = run_command(
        "python3 -m pytest tests/test_group_concat.py -v",
        "GROUP_CONCAT tests"
    )

    # Test 2: Right JOIN test
    print("\n" + "="*70)
    print("TEST 2: Right JOIN Test")
    print("="*70)
    returncode2, stdout2, stderr2 = run_command(
        "python3 -m pytest tests/test_join_types.py::TestRightJoin::test_basic_right_join -v",
        "Basic Right JOIN test"
    )

    # Summary
    print("\n" + "="*70)
    print("TEST SUMMARY")
    print("="*70)
    print(f"GROUP_CONCAT tests: {'PASSED ✓' if returncode1 == 0 else 'FAILED ✗'}")
    print(f"Right JOIN test: {'PASSED ✓' if returncode2 == 0 else 'FAILED ✗'}")

    if returncode1 != 0 or returncode2 != 0:
        print("\nFAILURE DETAILS:")
        if returncode1 != 0:
            print("\n--- GROUP_CONCAT Test Failures ---")
        if returncode2 != 0:
            print("\n--- Right JOIN Test Failures ---")

    return returncode1 + returncode2

if __name__ == "__main__":
    sys.exit(main())
