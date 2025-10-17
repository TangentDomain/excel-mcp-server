#!/usr/bin/env python3
"""
Excel MCP Server - 综合测试运行脚本

这是一个全面的测试运行器，包含：
1. 测试环境检查
2. 依赖项验证
3. 单元测试运行
4. 集成测试运行
5. 性能测试运行
6. 覆盖率报告生成
7. 结果汇总和报告
8. 错误分析和建议

使用方法:
    python run-all-tests.py                    # 运行所有检查和测试
    python run-all-tests.py --quick           # 快速模式（跳过慢速测试）
    python run-all-tests.py --coverage-only   # 仅运行覆盖率测试
    python run-all-tests.py --performance     # 仅运行性能测试
    python run-all-tests.py --validate-only   # 仅验证环境和依赖
"""

import argparse
import asyncio
import json
import os
import platform
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

# 可选依赖处理
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# 颜色输出支持
class Colors:
    """终端颜色常量"""
    RESET = '\033[0m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'

# Windows 兼容性处理
if platform.system() == 'Windows':
    # 禁用 ANSI 颜色，或者启用 Windows 10+ 的 ANSI 支持
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except:
        # 如果失败，创建无颜色的版本
        class Colors:
            RESET = ''
            RED = ''
            GREEN = ''
            YELLOW = ''
            BLUE = ''
            PURPLE = ''
            CYAN = ''
            WHITE = ''
            BOLD = ''

class TestResult:
    """测试结果数据类"""
    def __init__(self, name: str, category: str):
        self.name = name
        self.category = category
        self.success = False
        self.duration = 0.0
        self.stdout = ""
        self.stderr = ""
        self.error_message = ""
        self.test_count = 0
        self.failure_count = 0
        self.error_count = 0
        self.skip_count = 0
        self.coverage_percentage = 0.0

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            'name': self.name,
            'category': self.category,
            'success': self.success,
            'duration': self.duration,
            'test_count': self.test_count,
            'failure_count': self.failure_count,
            'error_count': self.error_count,
            'skip_count': self.skip_count,
            'coverage_percentage': self.coverage_percentage,
            'error_message': self.error_message
        }

class EnvironmentChecker:
    """环境检查器"""

    def __init__(self):
        self.project_root = Path(__file__).parent
        self.src_path = self.project_root / "src"
        self.tests_path = self.project_root / "tests"
        self.requirements = {
            'python_version': '3.10',
            'required_modules': [
                'fastmcp', 'openpyxl', 'mcp', 'xlcalculator',
                'formulas', 'xlwings'
            ],
            'dev_modules': [
                'pytest', 'pytest-asyncio', 'pytest-cov',
                'pytest-mock', 'pytest-xdist'
            ]
        }

    def check_python_version(self) -> Tuple[bool, str]:
        """检查Python版本"""
        version = sys.version_info
        required_version = tuple(map(int, self.requirements['python_version'].split('.')))

        if version >= required_version:
            return True, f"Python {version.major}.{version.minor}.{version.micro} OK"
        else:
            return False, f"Python {version.major}.{version.minor}.{version.micro} FAIL (需要 >= {self.requirements['python_version']})"

    def check_project_structure(self) -> Tuple[bool, List[str]]:
        """检查项目结构"""
        required_dirs = ['src', 'src/api', 'src/core', 'src/utils', 'tests']
        required_files = [
            'src/server.py',
            'src/api/excel_operations.py',
            'pyproject.toml',
            'CLAUDE.md'
        ]

        missing_items = []

        for dir_path in required_dirs:
            if not (self.project_root / dir_path).exists():
                missing_items.append(f"目录: {dir_path}")

        for file_path in required_files:
            if not (self.project_root / file_path).exists():
                missing_items.append(f"文件: {file_path}")

        return len(missing_items) == 0, missing_items

    def check_modules(self) -> Tuple[bool, List[str]]:
        """检查Python模块"""
        missing_modules = []

        for module in self.requirements['required_modules']:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)

        return len(missing_modules) == 0, missing_modules

    def check_dev_modules(self) -> Tuple[bool, List[str]]:
        """检查开发依赖模块"""
        missing_modules = []

        for module in self.requirements['dev_modules']:
            try:
                __import__(module.replace('-', '_'))
            except ImportError:
                missing_modules.append(module)

        return len(missing_modules) == 0, missing_modules

    def check_system_resources(self) -> Dict[str, Any]:
        """检查系统资源"""
        if HAS_PSUTIL:
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage(str(self.project_root))

            return {
                'cpu_count': psutil.cpu_count(),
                'memory_total_gb': round(memory.total / (1024**3), 2),
                'memory_available_gb': round(memory.available / (1024**3), 2),
                'memory_percent': memory.percent,
                'disk_free_gb': round(disk.free / (1024**3), 2),
                'disk_total_gb': round(disk.total / (1024**3), 2),
            }
        else:
            # 基本系统信息，不依赖psutil
            import os
            try:
                disk_usage = os.statvfs(str(self.project_root)) if hasattr(os, 'statvfs') else None
                disk_free = round(disk_usage.f_frsize * disk_usage.f_bavail / (1024**3), 2) if disk_usage else "Unknown"
            except:
                disk_free = "Unknown"

            return {
                'cpu_count': os.cpu_count() if hasattr(os, 'cpu_count') else "Unknown",
                'memory_total_gb': "Unknown",
                'memory_available_gb': "Unknown",
                'memory_percent': "Unknown",
                'disk_free_gb': disk_free,
                'disk_total_gb': "Unknown",
            }

class TestRunner:
    """测试运行器"""

    def __init__(self, project_root: Path):
        self.project_root = project_root
        self.results: List[TestResult] = []
        self.start_time = time.time()

    def run_command(self, cmd: List[str], timeout: int = 300, capture_output: bool = True) -> Tuple[bool, str, str]:
        """运行命令并返回结果"""
        try:
            print(f"  执行命令: {' '.join(cmd)}")

            result = subprocess.run(
                cmd,
                cwd=str(self.project_root),
                timeout=timeout,
                capture_output=capture_output,
                text=True,
                encoding='utf-8'
            )

            return result.returncode == 0, result.stdout, result.stderr

        except subprocess.TimeoutExpired:
            return False, "", f"命令执行超时 ({timeout}秒)"
        except Exception as e:
            return False, "", f"命令执行异常: {str(e)}"

    def parse_pytest_output(self, stdout: str) -> Dict[str, int]:
        """解析pytest输出，提取测试统计信息"""
        stats = {
            'test_count': 0,
            'failure_count': 0,
            'error_count': 0,
            'skip_count': 0
        }

        # 查找pytest的总结行
        lines = stdout.split('\n')
        for line in lines:
            if ' passed' in line or ' failed' in line or ' errors' in line or ' skipped' in line:
                # 示例: "10 passed, 2 failed, 1 errors, 3 skipped in 15.3s"
                parts = line.split()[0]  # "10"
                if parts.replace(',', '').isdigit():
                    stats['test_count'] = int(parts.replace(',', ''))

                if ' failed' in line:
                    for i, part in enumerate(line.split()):
                        if part.isdigit() and 'failed' in line.split()[i+1]:
                            stats['failure_count'] = int(part)
                            break

                if ' errors' in line:
                    for i, part in enumerate(line.split()):
                        if part.isdigit() and 'errors' in line.split()[i+1]:
                            stats['error_count'] = int(part)
                            break

                if ' skipped' in line:
                    for i, part in enumerate(line.split()):
                        if part.isdigit() and 'skipped' in line.split()[i+1]:
                            stats['skip_count'] = int(part)
                            break

        return stats

    def parse_coverage_output(self, stdout: str) -> float:
        """解析覆盖率输出"""
        lines = stdout.split('\n')
        for line in lines:
            if 'TOTAL' in line and '%' in line:
                # 示例: "TOTAL                            15      5    67%"
                parts = line.strip().split()
                for part in parts:
                    if part.endswith('%'):
                        try:
                            return float(part.rstrip('%'))
                        except ValueError:
                            pass
        return 0.0

    def run_unit_tests(self, quick_mode: bool = False) -> TestResult:
        """运行单元测试"""
        result = TestResult("单元测试", "unit")
        start_time = time.time()

        cmd = [
            "python", "-m", "pytest", "tests/",
            "-m", "not integration and not performance and not slow",
            "-v", "--tb=short", "--color=yes"
        ]

        if quick_mode:
            cmd.extend(["--maxfail=5", "-x"])  # 快速失败

        success, stdout, stderr = self.run_command(cmd, timeout=600)

        result.duration = time.time() - start_time
        result.success = success
        result.stdout = stdout
        result.stderr = stderr

        # 解析测试统计
        stats = self.parse_pytest_output(stdout)
        result.test_count = stats['test_count']
        result.failure_count = stats['failure_count']
        result.error_count = stats['error_count']
        result.skip_count = stats['skip_count']

        if not success:
            result.error_message = stderr or "单元测试执行失败"

        return result

    def run_integration_tests(self) -> TestResult:
        """运行集成测试"""
        result = TestResult("集成测试", "integration")
        start_time = time.time()

        cmd = [
            "python", "-m", "pytest", "tests/",
            "-m", "integration",
            "-v", "--tb=short", "--color=yes"
        ]

        success, stdout, stderr = self.run_command(cmd, timeout=900)

        result.duration = time.time() - start_time
        result.success = success
        result.stdout = stdout
        result.stderr = stderr

        stats = self.parse_pytest_output(stdout)
        result.test_count = stats['test_count']
        result.failure_count = stats['failure_count']
        result.error_count = stats['error_count']
        result.skip_count = stats['skip_count']

        if not success:
            result.error_message = stderr or "集成测试执行失败"

        return result

    def run_performance_tests(self) -> TestResult:
        """运行性能测试"""
        result = TestResult("性能测试", "performance")
        start_time = time.time()

        cmd = [
            "python", "-m", "pytest", "tests/",
            "-m", "performance",
            "-v", "-s", "--tb=short", "--color=yes"
        ]

        success, stdout, stderr = self.run_command(cmd, timeout=1200)

        result.duration = time.time() - start_time
        result.success = success
        result.stdout = stdout
        result.stderr = stderr

        stats = self.parse_pytest_output(stdout)
        result.test_count = stats['test_count']
        result.failure_count = stats['failure_count']
        result.error_count = stats['error_count']
        result.skip_count = stats['skip_count']

        if not success:
            result.error_message = stderr or "性能测试执行失败"

        return result

    def run_coverage_tests(self) -> TestResult:
        """运行覆盖率测试"""
        result = TestResult("覆盖率测试", "coverage")
        start_time = time.time()

        cmd = [
            "python", "-m", "pytest", "tests/",
            "--cov=src",
            "--cov-report=term-missing",
            "--cov-report=html",
            "--cov-report=xml",
            "--cov-fail-under=70",
            "-v", "--tb=short", "--color=yes"
        ]

        success, stdout, stderr = self.run_command(cmd, timeout=900)

        result.duration = time.time() - start_time
        result.success = success
        result.stdout = stdout
        result.stderr = stderr

        # 解析覆盖率百分比
        result.coverage_percentage = self.parse_coverage_output(stdout)

        stats = self.parse_pytest_output(stdout)
        result.test_count = stats['test_count']
        result.failure_count = stats['failure_count']
        result.error_count = stats['error_count']
        result.skip_count = stats['skip_count']

        if not success:
            result.error_message = stderr or "覆盖率测试执行失败"

        return result

    def run_security_tests(self) -> TestResult:
        """运行安全测试"""
        result = TestResult("安全测试", "security")
        start_time = time.time()

        cmd = [
            "python", "-m", "pytest", "tests/",
            "-m", "security",
            "-v", "--tb=short", "--color=yes"
        ]

        success, stdout, stderr = self.run_command(cmd, timeout=600)

        result.duration = time.time() - start_time
        result.success = success
        result.stdout = stdout
        result.stderr = stderr

        stats = self.parse_pytest_output(stdout)
        result.test_count = stats['test_count']
        result.failure_count = stats['failure_count']
        result.error_count = stats['error_count']
        result.skip_count = stats['skip_count']

        if not success:
            result.error_message = stderr or "安全测试执行失败"

        return result

class ReportGenerator:
    """报告生成器"""

    def __init__(self, results: List[TestResult], env_info: Dict[str, Any]):
        self.results = results
        self.env_info = env_info
        self.total_duration = sum(r.duration for r in results)

    def print_summary(self):
        """打印测试总结"""
        print(f"\n{Colors.BOLD}{Colors.CYAN}{'='*80}")
        print("测试执行总结")
        print(f"{'='*80}{Colors.RESET}")

        total_tests = sum(r.test_count for r in self.results)
        total_failures = sum(r.failure_count for r in self.results)
        total_errors = sum(r.error_count for r in self.results)
        total_skips = sum(r.skip_count for r in self.results)

        # 总体状态
        all_success = all(r.success for r in self.results)
        status_color = Colors.GREEN if all_success else Colors.RED
        status_text = "✓ 全部通过" if all_success else "✗ 存在失败"

        print(f"\n{Colors.BOLD}总体状态: {status_color}{status_text}{Colors.RESET}")
        print(f"总执行时间: {self.total_duration:.2f} 秒")
        print(f"测试总数: {total_tests}")
        print(f"失败数量: {total_failures}")
        print(f"错误数量: {total_errors}")
        print(f"跳过数量: {total_skips}")

        # 分类结果
        print(f"\n{Colors.BOLD}分类结果:{Colors.RESET}")
        for result in self.results:
            status_color = Colors.GREEN if result.success else Colors.RED
            status_icon = "✓" if result.success else "✗"

            coverage_info = ""
            if result.coverage_percentage > 0:
                coverage_color = Colors.GREEN if result.coverage_percentage >= 70 else Colors.YELLOW
                coverage_info = f" (覆盖率: {coverage_color}{result.coverage_percentage:.1f}%{Colors.RESET})"

            print(f"  {status_icon} {result.name}: {status_color}{result.category}{Colors.RESET} "
                  f"({result.duration:.1f}s, {result.test_count} 测试){coverage_info}")

            if not result.success:
                print(f"    错误: {Colors.RED}{result.error_message[:100]}{Colors.RESET}")

    def print_environment_info(self):
        """打印环境信息"""
        print(f"\n{Colors.BOLD}{Colors.BLUE}环境信息{Colors.RESET}")
        print(f"Python版本: {self.env_info.get('python_version', 'Unknown')}")
        print(f"操作系统: {platform.system()} {platform.release()}")
        print(f"CPU核心数: {self.env_info.get('system_resources', {}).get('cpu_count', 'Unknown')}")
        print(f"可用内存: {self.env_info.get('system_resources', {}).get('memory_available_gb', 'Unknown')} GB")
        print(f"磁盘空间: {self.env_info.get('system_resources', {}).get('disk_free_gb', 'Unknown')} GB")

    def print_failed_analysis(self):
        """打印失败分析"""
        failed_results = [r for r in self.results if not r.success]

        if not failed_results:
            print(f"\n{Colors.GREEN}{Colors.BOLD}🎉 所有测试都通过了！{Colors.RESET}")
            return

        print(f"\n{Colors.RED}{Colors.BOLD}失败测试分析{Colors.RESET}")

        for result in failed_results:
            print(f"\n{Colors.RED}● {result.name} ({result.category}){Colors.RESET}")
            print(f"  错误信息: {result.error_message}")

            if result.failure_count > 0:
                print(f"  失败测试数: {result.failure_count}")
            if result.error_count > 0:
                print(f"  错误测试数: {result.error_count}")

            # 提供建议
            print(f"  {Colors.YELLOW}建议:{Colors.RESET}")

            if result.category == "unit":
                print("    - 检查代码逻辑是否正确")
                print("    - 确认测试用例是否覆盖了所有边界情况")
                print("    - 运行: python -m pytest tests/test_api_excel_operations.py -v -s")

            elif result.category == "integration":
                print("    - 检查模块间的集成是否正常")
                print("    - 确认外部依赖是否可用")
                print("    - 运行: python -m pytest tests/test_integration_comprehensive.py -v -s")

            elif result.category == "performance":
                print("    - 检查系统资源是否充足")
                print("    - 优化算法或增加缓存")
                print("    - 运行: python -m pytest tests/test_performance.py -v -s")

            elif result.category == "coverage":
                print("    - 增加测试用例以提高覆盖率")
                print("    - 目标覆盖率应达到 70% 以上")
                print("    - 查看详细报告: htmlcov/index.html")

    def save_report(self, filename: str = None):
        """保存测试报告到JSON文件"""
        if filename is None:
            filename = f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"

        report_data = {
            'timestamp': datetime.now().isoformat(),
            'environment': self.env_info,
            'summary': {
                'total_duration': self.total_duration,
                'total_tests': sum(r.test_count for r in self.results),
                'total_failures': sum(r.failure_count for r in self.results),
                'total_errors': sum(r.error_count for r in self.results),
                'total_skips': sum(r.skip_count for r in self.results),
                'all_success': all(r.success for r in self.results)
            },
            'results': [r.to_dict() for r in self.results]
        }

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)

        print(f"\n{Colors.CYAN}详细报告已保存到: {filename}{Colors.RESET}")

def cleanup_test_artifacts(project_root: Path):
    """清理测试产生的临时文件"""
    print(f"\n{Colors.YELLOW}清理测试临时文件...{Colors.RESET}")

    artifacts = [
        ".pytest_cache",
        ".coverage",
        "coverage.xml",
        "htmlcov",
        ".benchmarks",
        "*.pyc",
        "__pycache__"
    ]

    removed_count = 0

    for pattern in artifacts:
        if pattern.startswith("*."):
            # 处理文件模式
            for file_path in project_root.rglob(pattern[2:]):
                try:
                    file_path.unlink()
                    print(f"  删除文件: {file_path.relative_to(project_root)}")
                    removed_count += 1
                except:
                    pass
        else:
            # 处理目录
            artifact_path = project_root / pattern
            if artifact_path.exists():
                try:
                    if artifact_path.is_dir():
                        shutil.rmtree(artifact_path)
                        print(f"  删除目录: {pattern}")
                    else:
                        artifact_path.unlink()
                        print(f"  删除文件: {pattern}")
                    removed_count += 1
                except Exception as e:
                    print(f"  无法删除 {pattern}: {e}")

    print(f"{Colors.GREEN}清理完成，共删除 {removed_count} 个项目{Colors.RESET}")

def print_banner():
    """打印程序横幅"""
    print(f"{Colors.BOLD}{Colors.CYAN}")
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║         Excel MCP Server - 综合测试运行脚本                    ║")
    print("║                                                              ║")
    print("║  完整的测试解决方案：环境检查、单元测试、集成测试、            ║")
    print("║  性能测试、覆盖率分析和详细的错误报告                        ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print(f"{Colors.RESET}")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Excel MCP Server 综合测试运行脚本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python run-all-tests.py                          # 运行所有测试
  python run-all-tests.py --quick                  # 快速模式
  python run-all-tests.py --coverage-only          # 仅覆盖率测试
  python run-all-tests.py --performance            # 仅性能测试
  python run-all-tests.py --validate-only          # 仅验证环境
  python run-all-tests.py --cleanup                # 清理临时文件
  python run-all-tests.py --report-file my.json    # 指定报告文件
        """
    )

    parser.add_argument(
        '--quick', '-q',
        action='store_true',
        help='快速模式（跳过慢速测试和性能测试）'
    )

    parser.add_argument(
        '--coverage-only', '-c',
        action='store_true',
        help='仅运行覆盖率测试'
    )

    parser.add_argument(
        '--performance', '-p',
        action='store_true',
        help='仅运行性能测试'
    )

    parser.add_argument(
        '--validate-only', '-v',
        action='store_true',
        help='仅验证环境和依赖'
    )

    parser.add_argument(
        '--cleanup',
        action='store_true',
        help='清理测试临时文件'
    )

    parser.add_argument(
        '--no-color',
        action='store_true',
        help='禁用彩色输出'
    )

    parser.add_argument(
        '--report-file', '-r',
        type=str,
        help='指定测试报告文件名'
    )

    args = parser.parse_args()

    # 禁用颜色输出
    if args.no_color:
        global Colors
        Colors = type('Colors', (), {
            'RESET': '', 'RED': '', 'GREEN': '', 'YELLOW': '',
            'BLUE': '', 'PURPLE': '', 'CYAN': '', 'WHITE': '', 'BOLD': ''
        })()

    print_banner()

    project_root = Path(__file__).parent

    # 清理模式
    if args.cleanup:
        cleanup_test_artifacts(project_root)
        return

    # 环境检查
    print(f"{Colors.BOLD}{Colors.BLUE}第一阶段：环境检查{Colors.RESET}")
    print("-" * 50)

    checker = EnvironmentChecker()
    env_info = {}
    env_valid = True

    # Python版本检查
    python_ok, python_msg = checker.check_python_version()
    print(f"Python版本: {python_msg}")
    if not python_ok:
        env_valid = False

    env_info['python_version'] = python_msg

    # 项目结构检查
    structure_ok, missing_items = checker.check_project_structure()
    if structure_ok:
        print("项目结构: OK 所有必需目录和文件都存在")
    else:
        print("项目结构: FAIL 缺少以下项目:")
        for item in missing_items:
            print(f"  - {item}")
        env_valid = False

    # 模块依赖检查
    modules_ok, missing_modules = checker.check_modules()
    if modules_ok:
        print("核心模块: ✓ 所有必需模块都已安装")
    else:
        print(f"核心模块: ✗ 缺少以下模块:")
        for module in missing_modules:
            print(f"  - {module}")
        env_valid = False

    # 开发依赖检查
    dev_modules_ok, missing_dev_modules = checker.check_dev_modules()
    if dev_modules_ok:
        print("开发模块: ✓ 所有开发模块都已安装")
    else:
        print(f"开发模块: ✗ 缺少以下开发模块:")
        for module in missing_dev_modules:
            print(f"  - {module}")
        print("  建议: pip install -e .[dev]")
        # 开发模块不是必须的，只警告

    # 系统资源检查
    system_resources = checker.check_system_resources()
    env_info['system_resources'] = system_resources
    print(f"系统资源: CPU {system_resources['cpu_count']} 核, "
          f"内存 {system_resources['memory_available_gb']:.1f}GB 可用, "
          f"磁盘 {system_resources['disk_free_gb']:.1f}GB 可用")

    # 如果环境验证失败，退出
    if not env_valid:
        print(f"\n{Colors.RED}环境验证失败，请解决上述问题后重新运行{Colors.RESET}")
        sys.exit(1)

    # 仅验证环境模式
    if args.validate_only:
        print(f"\n{Colors.GREEN}环境验证完成！{Colors.RESET}")
        return

    # 测试执行
    print(f"\n{Colors.BOLD}{Colors.BLUE}第二阶段：测试执行{Colors.RESET}")
    print("-" * 50)

    runner = TestRunner(project_root)

    # 根据参数选择运行的测试
    if args.coverage_only:
        print("运行覆盖率测试...")
        result = runner.run_coverage_tests()
        runner.results.append(result)

    elif args.performance:
        print("运行性能测试...")
        result = runner.run_performance_tests()
        runner.results.append(result)

    else:
        # 完整测试流程
        if args.quick:
            print("快速模式：运行基础测试...")
            # 单元测试
            print("运行单元测试...")
            unit_result = runner.run_unit_tests(quick_mode=True)
            runner.results.append(unit_result)

            # 快速覆盖率测试
            print("运行覆盖率测试...")
            coverage_result = runner.run_coverage_tests()
            runner.results.append(coverage_result)

        else:
            print("完整模式：运行所有测试...")

            # 单元测试
            print("运行单元测试...")
            unit_result = runner.run_unit_tests()
            runner.results.append(unit_result)

            # 集成测试
            print("运行集成测试...")
            integration_result = runner.run_integration_tests()
            runner.results.append(integration_result)

            # 性能测试
            print("运行性能测试...")
            performance_result = runner.run_performance_tests()
            runner.results.append(performance_result)

            # 安全测试
            print("运行安全测试...")
            security_result = runner.run_security_tests()
            runner.results.append(security_result)

            # 覆盖率测试
            print("运行覆盖率测试...")
            coverage_result = runner.run_coverage_tests()
            runner.results.append(coverage_result)

    # 生成报告
    print(f"\n{Colors.BOLD}{Colors.BLUE}第三阶段：报告生成{Colors.RESET}")
    print("-" * 50)

    report_generator = ReportGenerator(runner.results, env_info)
    report_generator.print_environment_info()
    report_generator.print_summary()
    report_generator.print_failed_analysis()

    # 保存报告
    report_generator.save_report(args.report_file)

    # 最终状态
    all_success = all(r.success for r in runner.results)
    if all_success:
        print(f"\n{Colors.GREEN}{Colors.BOLD}🎉 所有测试都成功完成！{Colors.RESET}")
        exit_code = 0
    else:
        print(f"\n{Colors.RED}{Colors.BOLD}❌ 存在测试失败，请查看详细报告{Colors.RESET}")
        exit_code = 1

    # 返回适当的退出码
    sys.exit(exit_code)

if __name__ == "__main__":
    main()