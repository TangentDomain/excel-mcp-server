#!/usr/bin/env python3
"""
Enhanced Test Runner

提供多种测试运行模式，包括覆盖率报告、性能测试、并发测试等
"""

import argparse
import subprocess
import sys
import os
from pathlib import Path

def get_project_root():
    """获取项目根目录"""
    return Path(__file__).parent.parent

def run_command(cmd, cwd=None, capture_output=False):
    """运行命令并返回结果"""
    print(f"Running: {' '.join(cmd)}")

    try:
        result = subprocess.run(
            cmd,
            cwd=cwd or get_project_root(),
            capture_output=capture_output,
            text=True,
            check=True
        )

        if capture_output:
            return result.stdout, result.stderr
        return True, ""

    except subprocess.CalledProcessError as e:
        print(f"Command failed: {e}")
        if capture_output:
            return e.stdout, e.stderr
        return False, str(e)

def run_unit_tests(verbose=True):
    """运行单元测试"""
    print("[TEST] Running Unit Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v"])

    # 只运行单元测试（非集成测试）
    cmd.extend(["-m", "not integration and not performance"])
    cmd.append("tests/")

    return run_command(cmd)

def run_integration_tests(verbose=True):
    """运行集成测试"""
    print("[TEST] Running Integration Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v"])

    cmd.extend(["-m", "integration"])
    cmd.append("tests/")

    return run_command(cmd)

def run_performance_tests(verbose=True):
    """运行性能测试"""
    print("[TEST] Running Performance Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v", "-s"])

    cmd.extend(["-m", "performance"])
    cmd.append("tests/")

    return run_command(cmd)

def run_all_tests(verbose=True):
    """运行所有测试"""
    print("[TEST] Running All Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v"])

    cmd.append("tests/")

    return run_command(cmd)

def run_coverage_report(verbose=True):
    """运行覆盖率报告"""
    print("[TEST] Running Coverage Report...")

    cmd = [
        "python", "-m", "pytest",
        "--cov=src",
        "--cov-report=html",
        "--cov-report=term",
        "--cov-report=xml",
        "tests/"
    ]

    if verbose:
        cmd.append("-v")

    success, output = run_command(cmd)

    if success:
        print("[SUCCESS] Coverage report generated successfully!")
        print("[INFO] HTML report: htmlcov/index.html")
        print("[INFO] XML report: coverage.xml")
    else:
        print("[ERROR] Coverage report failed!")

    return success, output

def run_parallel_tests(verbose=True):
    """运行并行测试"""
    print("[TEST] Running Parallel Tests...")

    # 使用pytest-xdist进行并行测试
    cmd = [
        "python", "-m", "pytest",
        "-n", "auto",  # 自动检测CPU核心数
        "--dist", "loadscope",  # 按负载分配测试
    ]

    if verbose:
        cmd.append("-v")

    cmd.append("tests/")

    return run_command(cmd)

def run_benchmark_tests(verbose=True):
    """运行基准测试"""
    print("[TEST] Running Benchmark Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v", "-s"])

    cmd.extend(["--benchmark-only"])
    cmd.append("tests/")

    return run_command(cmd)

def run_slow_tests(verbose=True):
    """运行慢速测试"""
    print("[TEST] Running Slow Tests...")

    cmd = ["python", "-m", "pytest"]

    if verbose:
        cmd.extend(["-v", "-s"])

    cmd.extend(["-m", "slow"])
    cmd.append("tests/")

    return run_command(cmd)

def run_failed_tests(verbose=True):
    """运行失败的测试"""
    print("[TEST] Running Failed Tests...")

    # 使用last-failed只运行上一次失败的测试
    cmd = [
        "python", "-m", "pytest",
        "--lf",  # last failed
    ]

    if verbose:
        cmd.append("-v")

    return run_command(cmd)

def cleanup_test_artifacts():
    """清理测试产生的文件"""
    print("[CLEAN] Cleaning up test artifacts...")

    artifacts = [
        ".pytest_cache",
        "htmlcov",
        ".coverage",
        "coverage.xml",
        ".benchmarks",
    ]

    for artifact in artifacts:
        artifact_path = get_project_root() / artifact
        if artifact_path.exists():
            if artifact_path.is_dir():
                import shutil
                shutil.rmtree(artifact_path)
                print(f"  [DIR] Removed directory: {artifact}")
            else:
                artifact_path.unlink()
                print(f"  [FILE] Removed file: {artifact}")

def list_test_categories():
    """列出可用的测试类别"""
    print("Available Test Categories:")
    print("  unit         - Unit tests only")
    print("  integration  - Integration tests")
    print("  performance - Performance tests")
    print("  all          - All tests")
    print("  coverage     - Tests with coverage report")
    print("  parallel     - Parallel execution")
    print("  benchmark    - Benchmark tests")
    print("  slow         - Slow tests")
    print("  failed       - Failed tests only")
    print("  cleanup      - Clean up test artifacts")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Enhanced Test Runner for Excel MCP Server",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scripts/run_tests_enhanced.py unit
  python scripts/run_tests_enhanced.py coverage
  python scripts/run_tests_enhanced.py all --verbose
  python scripts/run_tests_enhanced.py cleanup
        """
    )

    parser.add_argument(
        "category",
        choices=["unit", "integration", "performance", "all", "coverage",
                 "parallel", "benchmark", "slow", "failed", "cleanup", "list"],
        help="Test category to run"
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        default=True,
        help="Verbose output"
    )

    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Quiet output"
    )

    args = parser.parse_args()

    # 处理verbose/quiet选项
    verbose = args.verbose and not args.quiet

    # 根据类别运行相应的测试
    if args.category == "list":
        list_test_categories()
    elif args.category == "unit":
        run_unit_tests(verbose=verbose)
    elif args.category == "integration":
        run_integration_tests(verbose=verbose)
    elif args.category == "performance":
        run_performance_tests(verbose=verbose)
    elif args.category == "all":
        run_all_tests(verbose=verbose)
    elif args.category == "coverage":
        run_coverage_report(verbose=verbose)
    elif args.category == "parallel":
        run_parallel_tests(verbose=verbose)
    elif args.category == "benchmark":
        run_benchmark_tests(verbose=verbose)
    elif args.category == "slow":
        run_slow_tests(verbose=verbose)
    elif args.category == "failed":
        run_failed_tests(verbose=verbose)
    elif args.category == "cleanup":
        cleanup_test_artifacts()
    else:
        parser.print_help()
        sys.exit(1)

if __name__ == "__main__":
    main()