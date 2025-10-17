#!/usr/bin/env python3
"""
Excel MCP Server - 监控和维护脚本

该脚本提供全面的监控和维护功能，包括：
1. 覆盖率监控 - 检查测试覆盖率是否达到要求
2. 测试运行时间监控 - 监控测试执行时间和性能
3. 内存使用监控 - 检查测试过程中的内存使用情况
4. 测试质量评估 - 评估测试质量和稳定性
5. 自动报告生成 - 生成详细的监控报告
6. 维护建议 - 提供改进建议

使用方法:
    python scripts/monitor-and-maintain.py [选项]

选项:
    --coverage-only     仅运行覆盖率监控
    --performance-only  仅运行性能监控
    --memory-only       仅运行内存监控
    --quality-only      仅运行质量评估
    --report-file       指定报告文件路径 (默认: reports/monitoring-report.html)
    --threshold         覆盖率阈值 (默认: 85)
    --no-html           不生成HTML报告，仅输出文本
    --continuous        连续监控模式，每5分钟运行一次
"""

import argparse
import json
import os
import psutil
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import logging
from dataclasses import dataclass, asdict
import threading
import signal
import tempfile

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/monitor.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class CoverageMetrics:
    """覆盖率指标"""
    total_coverage: float
    file_coverage: Dict[str, float]
    missing_lines: Dict[str, List[int]]
    uncovered_files: List[str]
    timestamp: datetime


@dataclass
class PerformanceMetrics:
    """性能指标"""
    total_time: float
    test_count: int
    success_rate: float
    slowest_tests: List[Tuple[str, float]]
    fastest_tests: List[Tuple[str, float]]
    memory_usage_mb: float
    timestamp: datetime


@dataclass
class MemoryMetrics:
    """内存使用指标"""
    peak_memory_mb: float
    average_memory_mb: float
    memory_growth_mb: float
    process_count: int
    timestamp: datetime


@dataclass
class QualityMetrics:
    """质量指标"""
    test_stability: float
    flaky_tests: List[str]
    code_quality_score: float
    duplicate_coverage: float
    test_complexity_score: float
    timestamp: datetime


@dataclass
class MonitoringReport:
    """监控报告"""
    timestamp: datetime
    coverage: CoverageMetrics
    performance: PerformanceMetrics
    memory: MemoryMetrics
    quality: QualityMetrics
    recommendations: List[str]
    summary: Dict[str, Any]


class MemoryMonitor:
    """内存监控器"""

    def __init__(self):
        self.measurements = []
        self.monitoring = False
        self.monitor_thread = None
        self.start_time = None

    def start_monitoring(self):
        """开始监控内存使用"""
        self.monitoring = True
        self.start_time = time.time()
        self.measurements = []

        def monitor():
            process = psutil.Process()
            while self.monitoring:
                memory_info = process.memory_info()
                memory_mb = memory_info.rss / 1024 / 1024
                self.measurements.append({
                    'timestamp': time.time() - self.start_time,
                    'memory_mb': memory_mb
                })
                time.sleep(0.5)  # 每0.5秒采集一次

        self.monitor_thread = threading.Thread(target=monitor)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        logger.info("内存监控已启动")

    def stop_monitoring(self) -> MemoryMetrics:
        """停止监控并返回内存指标"""
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join()

        if not self.measurements:
            return MemoryMetrics(0, 0, 0, 0, datetime.now())

        memory_values = [m['memory_mb'] for m in self.measurements]
        peak_memory = max(memory_values)
        average_memory = sum(memory_values) / len(memory_values)

        # 计算内存增长
        if len(self.measurements) >= 2:
            initial_memory = self.measurements[0]['memory_mb']
            final_memory = self.measurements[-1]['memory_mb']
            memory_growth = final_memory - initial_memory
        else:
            memory_growth = 0

        logger.info(f"内存监控结束 - 峰值: {peak_memory:.1f}MB, 平均: {average_memory:.1f}MB, 增长: {memory_growth:.1f}MB")

        return MemoryMetrics(
            peak_memory_mb=peak_memory,
            average_memory_mb=average_memory,
            memory_growth_mb=memory_growth,
            process_count=len(self.measurements),
            timestamp=datetime.now()
        )


class TestMonitor:
    """测试监控器"""

    def __init__(self, coverage_threshold: float = 85.0):
        self.coverage_threshold = coverage_threshold
        self.base_dir = Path(__file__).parent.parent
        self.reports_dir = self.base_dir / "reports"
        self.logs_dir = self.base_dir / "logs"

        # 确保目录存在
        self.reports_dir.mkdir(exist_ok=True)
        self.logs_dir.mkdir(exist_ok=True)

    def run_coverage_monitoring(self) -> CoverageMetrics:
        """运行覆盖率监控"""
        logger.info("开始覆盖率监控...")

        try:
            # 运行覆盖率测试
            cmd = [
                sys.executable, "-m", "pytest",
                "tests/",
                "--cov=src",
                "--cov-report=json",
                "--cov-report=term-missing",
                "--cov-report=html:htmlcov",
                "-v"
            ]

            result = subprocess.run(
                cmd,
                cwd=self.base_dir,
                capture_output=True,
                text=True,
                timeout=300  # 5分钟超时
            )

            # 读取覆盖率报告
            coverage_file = self.base_dir / "coverage.json"
            if coverage_file.exists():
                with open(coverage_file, 'r', encoding='utf-8') as f:
                    coverage_data = json.load(f)

                total_coverage = coverage_data['totals']['percent_covered']
                file_coverage = {}
                missing_lines = {}
                uncovered_files = []

                for filename, file_data in coverage_data['files'].items():
                    file_coverage[filename] = file_data['summary']['percent_covered']

                    if file_data['summary']['missing_lines']:
                        missing_lines[filename] = file_data['summary']['missing_lines']

                    if file_data['summary']['percent_covered'] == 0:
                        uncovered_files.append(filename)

                logger.info(f"总覆盖率: {total_coverage:.1f}%")

                return CoverageMetrics(
                    total_coverage=total_coverage,
                    file_coverage=file_coverage,
                    missing_lines=missing_lines,
                    uncovered_files=uncovered_files,
                    timestamp=datetime.now()
                )
            else:
                logger.warning("未找到覆盖率报告文件")
                return CoverageMetrics(0, {}, {}, [], datetime.now())

        except subprocess.TimeoutExpired:
            logger.error("覆盖率测试超时")
            return CoverageMetrics(0, {}, {}, [], datetime.now())
        except Exception as e:
            logger.error(f"覆盖率监控失败: {e}")
            return CoverageMetrics(0, {}, {}, [], datetime.now())

    def run_performance_monitoring(self, memory_monitor: MemoryMonitor) -> PerformanceMetrics:
        """运行性能监控"""
        logger.info("开始性能监控...")

        start_time = time.time()

        try:
            # 启动内存监控
            memory_monitor.start_monitoring()

            # 运行测试并收集性能数据
            cmd = [
                sys.executable, "-m", "pytest",
                "tests/",
                "--tb=short",
                "-v",
                "--durations=10"  # 显示最慢的10个测试
            ]

            result = subprocess.run(
                cmd,
                cwd=self.base_dir,
                capture_output=True,
                text=True,
                timeout=600  # 10分钟超时
            )

            total_time = time.time() - start_time

            # 解析测试结果
            output_lines = result.stdout.split('\n')
            test_count = 0
            failed_count = 0
            slow_tests = []

            for line in output_lines:
                if "passed" in line and "failed" in line:
                    # 提取测试数量
                    import re
                    match = re.search(r'(\d+)\s+passed.*?(\d+)\s+failed', line)
                    if match:
                        test_count = int(match.group(1))
                        failed_count = int(match.group(2))
                elif "seconds" in line and "::" in line:
                    # 提取慢速测试
                    parts = line.split()
                    if len(parts) >= 2:
                        test_name = parts[-1]
                        duration = float(parts[0])
                        slow_tests.append((test_name, duration))

            success_rate = (test_count - failed_count) / max(test_count, 1) * 100

            # 停止内存监控
            memory_metrics = memory_monitor.stop_monitoring()

            logger.info(f"性能监控完成 - 总时间: {total_time:.1f}s, 测试数: {test_count}, 成功率: {success_rate:.1f}%")

            return PerformanceMetrics(
                total_time=total_time,
                test_count=test_count,
                success_rate=success_rate,
                slowest_tests=sorted(slow_tests, key=lambda x: x[1], reverse=True)[:5],
                fastest_tests=sorted(slow_tests, key=lambda x: x[1])[:5],
                memory_usage_mb=memory_metrics.peak_memory_mb,
                timestamp=datetime.now()
            )

        except subprocess.TimeoutExpired:
            memory_monitor.stop_monitoring()
            logger.error("性能测试超时")
            return PerformanceMetrics(600, 0, 0, [], [], 0, datetime.now())
        except Exception as e:
            memory_monitor.stop_monitoring()
            logger.error(f"性能监控失败: {e}")
            return PerformanceMetrics(0, 0, 0, [], [], 0, datetime.now())

    def run_quality_assessment(self, coverage_metrics: CoverageMetrics,
                             performance_metrics: PerformanceMetrics) -> QualityMetrics:
        """运行质量评估"""
        logger.info("开始质量评估...")

        try:
            # 计算测试稳定性
            test_stability = self._calculate_test_stability()

            # 检测flaky测试
            flaky_tests = self._detect_flaky_tests()

            # 计算代码质量分数
            code_quality_score = self._calculate_code_quality_score(coverage_metrics)

            # 计算重复覆盖率
            duplicate_coverage = self._calculate_duplicate_coverage()

            # 计算测试复杂度分数
            test_complexity_score = self._calculate_test_complexity()

            logger.info(f"质量评估完成 - 稳定性: {test_stability:.1f}, 质量分数: {code_quality_score:.1f}")

            return QualityMetrics(
                test_stability=test_stability,
                flaky_tests=flaky_tests,
                code_quality_score=code_quality_score,
                duplicate_coverage=duplicate_coverage,
                test_complexity_score=test_complexity_score,
                timestamp=datetime.now()
            )

        except Exception as e:
            logger.error(f"质量评估失败: {e}")
            return QualityMetrics(0, [], 0, 0, 0, datetime.now())

    def _calculate_test_stability(self) -> float:
        """计算测试稳定性"""
        try:
            # 运行测试多次来检测不稳定的测试
            stability_scores = []

            for i in range(3):  # 运行3次
                cmd = [sys.executable, "-m", "pytest", "tests/", "-q"]
                result = subprocess.run(cmd, cwd=self.base_dir, capture_output=True)

                # 统计通过的测试数量
                if result.returncode == 0:
                    output = result.stdout.decode('utf-8', errors='ignore')
                    import re
                    match = re.search(r'(\d+)\s+passed', output)
                    if match:
                        stability_scores.append(int(match.group(1)))

            if len(stability_scores) > 1:
                # 计算标准差
                mean = sum(stability_scores) / len(stability_scores)
                variance = sum((x - mean) ** 2 for x in stability_scores) / len(stability_scores)
                std_dev = variance ** 0.5

                # 稳定性分数 (标准差越小，稳定性越高)
                stability = max(0, 100 - std_dev)
                return stability
            else:
                return 100.0  # 无法计算稳定性，假设稳定

        except Exception as e:
            logger.warning(f"测试稳定性计算失败: {e}")
            return 0.0

    def _detect_flaky_tests(self) -> List[str]:
        """检测不稳定的测试"""
        flaky_tests = []

        try:
            # 这里可以实现更复杂的flaky测试检测逻辑
            # 比如运行多次并比较结果
            test_results = {}

            for i in range(3):
                cmd = [sys.executable, "-m", "pytest", "tests/", "--tb=no", "-v"]
                result = subprocess.run(cmd, cwd=self.base_dir, capture_output=True)

                if result.returncode != 0:
                    # 解析失败的测试
                    output = result.stderr.decode('utf-8', errors='ignore')
                    lines = output.split('\n')
                    for line in lines:
                        if '::' in line and 'FAILED' in line:
                            test_name = line.split('FAILED')[0].strip()
                            test_results[test_name] = test_results.get(test_name, 0) + 1

            # 找出失败的测试
            flaky_tests = [test for test, count in test_results.items() if count > 1]

        except Exception as e:
            logger.warning(f"Flaky测试检测失败: {e}")

        return flaky_tests

    def _calculate_code_quality_score(self, coverage_metrics: CoverageMetrics) -> float:
        """计算代码质量分数"""
        score = 0.0

        # 覆盖率分数 (40%)
        coverage_score = min(100, coverage_metrics.total_coverage * 100 / self.coverage_threshold)
        score += coverage_score * 0.4

        # 文件覆盖率一致性分数 (20%)
        if coverage_metrics.file_coverage:
            coverages = list(coverage_metrics.file_coverage.values())
            avg_coverage = sum(coverages) / len(coverages)
            consistency = 100 - (max(coverages) - min(coverages))
            score += min(100, consistency) * 0.2

        # 无未覆盖文件分数 (20%)
        uncovered_penalty = len(coverage_metrics.uncovered_files) * 10
        uncovered_score = max(0, 100 - uncovered_penalty)
        score += uncovered_score * 0.2

        # 基础分数 (20%)
        score += 20

        return min(100, score)

    def _calculate_duplicate_coverage(self) -> float:
        """计算重复覆盖率"""
        try:
            # 分析测试用例的重复性
            cmd = [sys.executable, "-m", "pytest", "tests/", "--collect-only", "-q"]
            result = subprocess.run(cmd, cwd=self.base_dir, capture_output=True, text=True)

            if result.returncode == 0:
                output = result.stdout
                test_names = [line.strip() for line in output.split('\n') if '::test_' in line]

                # 简单的重复检测：查找相似的测试名称
                duplicate_count = 0
                for i, test1 in enumerate(test_names):
                    for test2 in test_names[i+1:]:
                        # 简单的相似度检测
                        if test1.split('::')[-1].split('_')[0] == test2.split('::')[-1].split('_')[0]:
                            duplicate_count += 1

                if len(test_names) > 0:
                    duplicate_percentage = (duplicate_count / len(test_names)) * 100
                    return duplicate_percentage

            return 0.0

        except Exception as e:
            logger.warning(f"重复覆盖率计算失败: {e}")
            return 0.0

    def _calculate_test_complexity(self) -> float:
        """计算测试复杂度分数"""
        try:
            complexity_scores = []

            # 分析测试文件的复杂度
            test_files = list(Path(self.base_dir / "tests").glob("test_*.py"))

            for test_file in test_files:
                try:
                    with open(test_file, 'r', encoding='utf-8') as f:
                        content = f.read()

                    # 简单的复杂度指标
                    lines = len(content.split('\n'))
                    functions = content.count('def test_')
                    asserts = content.count('assert')

                    # 计算复杂度分数
                    if functions > 0:
                        avg_lines_per_test = lines / functions
                        asserts_per_test = asserts / functions

                        # 理想情况下每个测试10-20行，包含1-3个断言
                        length_score = max(0, 100 - abs(avg_lines_per_test - 15) * 2)
                        assert_score = max(0, 100 - abs(asserts_per_test - 2) * 20)

                        test_score = (length_score + assert_score) / 2
                        complexity_scores.append(test_score)

                except Exception:
                    continue

            if complexity_scores:
                return sum(complexity_scores) / len(complexity_scores)
            else:
                return 50.0  # 默认分数

        except Exception as e:
            logger.warning(f"测试复杂度计算失败: {e}")
            return 0.0

    def generate_recommendations(self, report: MonitoringReport) -> List[str]:
        """生成维护建议"""
        recommendations = []

        # 覆盖率建议
        if report.coverage.total_coverage < self.coverage_threshold:
            recommendations.append(
                f"⚠️ 测试覆盖率 ({report.coverage.total_coverage:.1f}%) 低于要求 ({self.coverage_threshold}%)，"
                f"建议增加测试用例，特别是以下文件：{', '.join(report.coverage.uncovered_files[:3])}"
            )

        # 性能建议
        if report.performance.total_time > 300:  # 5分钟
            recommendations.append(
                f"⏱️ 测试执行时间过长 ({report.performance.total_time:.1f}s)，"
                f"建议优化以下慢速测试：{', '.join([name for name, _ in report.performance.slowest_tests[:3]])}"
            )

        if report.performance.memory_usage_mb > 1000:  # 1GB
            recommendations.append(
                f"💾 内存使用过高 ({report.performance.memory_usage_mb:.1f}MB)，"
                "建议检查内存泄漏或优化测试数据管理"
            )

        # 质量建议
        if report.quality.test_stability < 90:
            recommendations.append(
                f"🔄 测试稳定性较低 ({report.quality.test_stability:.1f}%)，"
                f"需要修复以下不稳定的测试：{', '.join(report.quality.flaky_tests[:3])}"
            )

        if report.quality.code_quality_score < 80:
            recommendations.append(
                f"📊 代码质量分数较低 ({report.quality.code_quality_score:.1f}%)，"
                "建议改进测试结构和覆盖率分布"
            )

        if report.quality.duplicate_coverage > 20:
            recommendations.append(
                f"🔄 发现重复测试用例 ({report.quality.duplicate_coverage:.1f}%)，"
                "建议重构测试以减少重复"
            )

        if not recommendations:
            recommendations.append("✅ 所有指标都在良好范围内，继续保持！")

        return recommendations

    def generate_report(self, metrics: Dict[str, Any], output_file: str = None) -> str:
        """生成监控报告"""
        logger.info("生成监控报告...")

        # 创建报告数据
        coverage = metrics['coverage']
        performance = metrics['performance']
        memory = metrics['memory']
        quality = metrics['quality']

        report = MonitoringReport(
            timestamp=datetime.now(),
            coverage=coverage,
            performance=performance,
            memory=memory,
            quality=quality,
            recommendations=metrics.get('recommendations', []),
            summary={
                'overall_score': self._calculate_overall_score(coverage, performance, memory, quality),
                'status': self._get_overall_status(coverage, performance, memory, quality),
                'key_metrics': {
                    'coverage': coverage.total_coverage,
                    'test_count': performance.test_count,
                    'success_rate': performance.success_rate,
                    'memory_peak': memory.peak_memory_mb,
                    'quality_score': quality.code_quality_score
                }
            }
        )

        # 生成HTML报告
        if output_file and not output_file.endswith('.txt'):
            html_content = self._generate_html_report(report)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            logger.info(f"HTML报告已生成: {output_file}")

        # 生成文本报告
        text_report = self._generate_text_report(report)

        # 如果是连续监控模式，追加到日志文件
        if hasattr(self, '_continuous_mode') and self._continuous_mode:
            log_file = self.logs_dir / "continuous_monitor.log"
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n{datetime.now().isoformat()}\n{'='*60}\n")
                f.write(text_report)

        return text_report

    def _calculate_overall_score(self, coverage, performance, memory, quality) -> float:
        """计算总体分数"""
        score = 0.0

        # 覆盖率分数 (30%)
        coverage_score = min(100, coverage.total_coverage)
        score += coverage_score * 0.3

        # 性能分数 (25%)
        performance_score = 0
        if performance.success_rate > 95:
            performance_score += 50
        if performance.total_time < 180:  # 3分钟
            performance_score += 30
        if performance.memory_usage_mb < 500:  # 500MB
            performance_score += 20
        score += min(100, performance_score) * 0.25

        # 内存分数 (20%)
        memory_score = 0
        if memory.peak_memory_mb < 500:
            memory_score = 100
        elif memory.peak_memory_mb < 1000:
            memory_score = 80
        elif memory.peak_memory_mb < 1500:
            memory_score = 60
        else:
            memory_score = 40
        score += memory_score * 0.2

        # 质量分数 (25%)
        score += quality.code_quality_score * 0.25

        return min(100, score)

    def _get_overall_status(self, coverage, performance, memory, quality) -> str:
        """获取总体状态"""
        overall_score = self._calculate_overall_score(coverage, performance, memory, quality)

        if overall_score >= 90:
            return "优秀"
        elif overall_score >= 80:
            return "良好"
        elif overall_score >= 70:
            return "一般"
        elif overall_score >= 60:
            return "需要改进"
        else:
            return "严重问题"

    def _generate_html_report(self, report: MonitoringReport) -> str:
        """生成HTML报告"""
        html_template = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel MCP Server - 监控报告</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 2.5em; }
        .header .timestamp { opacity: 0.9; margin-top: 10px; }
        .status { display: inline-block; padding: 8px 16px; border-radius: 20px; margin-top: 15px; font-weight: bold; }
        .status.优秀 { background: #4caf50; }
        .status.良好 { background: #8bc34a; }
        .status.一般 { background: #ff9800; }
        .status.需要改进 { background: #f44336; }
        .status.严重问题 { background: #d32f2f; }
        .content { padding: 30px; }
        .section { margin-bottom: 40px; }
        .section h2 { color: #333; border-bottom: 2px solid #667eea; padding-bottom: 10px; margin-bottom: 20px; }
        .metrics-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .metric-card { background: #f8f9fa; border-radius: 8px; padding: 20px; border-left: 4px solid #667eea; }
        .metric-card h3 { margin: 0 0 15px 0; color: #495057; }
        .metric-value { font-size: 2em; font-weight: bold; color: #667eea; margin-bottom: 5px; }
        .metric-label { color: #6c757d; font-size: 0.9em; }
        .progress-bar { background: #e9ecef; border-radius: 4px; height: 8px; overflow: hidden; margin-top: 10px; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); transition: width 0.3s ease; }
        .recommendations { background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; padding: 20px; }
        .recommendations h3 { color: #856404; margin-top: 0; }
        .recommendations ul { margin: 10px 0; padding-left: 20px; }
        .recommendations li { margin-bottom: 8px; color: #856404; }
        .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        .table th, .table td { padding: 12px; text-align: left; border-bottom: 1px solid #dee2e6; }
        .table th { background: #f8f9fa; font-weight: 600; color: #495057; }
        .chart-placeholder { background: #f8f9fa; border: 2px dashed #dee2e6; border-radius: 8px; padding: 40px; text-align: center; color: #6c757d; }
        .footer { background: #f8f9fa; padding: 20px; text-align: center; color: #6c757d; font-size: 0.9em; }
        .score-circle { width: 120px; height: 120px; border-radius: 50%; background: conic-gradient(#667eea 0deg, #667eea {score}deg, #e9ecef {score}deg); display: flex; align-items: center; justify-content: center; margin: 20px auto; position: relative; }
        .score-circle::before { content: ''; position: absolute; width: 100px; height: 100px; background: white; border-radius: 50%; }
        .score-text { position: relative; font-size: 1.8em; font-weight: bold; color: #667eea; }
        .good { color: #28a745; }
        .warning { color: #ffc107; }
        .bad { color: #dc3545; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 Excel MCP Server 监控报告</h1>
            <div class="timestamp">{timestamp}</div>
            <div class="status {status}">总体状态: {status}</div>
        </div>

        <div class="content">
            <!-- 总体分数 -->
            <div class="section">
                <h2>📊 总体评估</h2>
                <div class="metrics-grid">
                    <div class="metric-card" style="text-align: center;">
                        <h3>综合评分</h3>
                        <div class="score-circle">
                            <div class="score-text">{overall_score:.0f}</div>
                        </div>
                        <div class="metric-label">满分100分</div>
                    </div>
                    <div class="metric-card">
                        <h3>关键指标</h3>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 10px;">
                            <div>
                                <div class="metric-value" style="font-size: 1.2em;">{key_coverage:.1f}%</div>
                                <div class="metric-label">测试覆盖率</div>
                            </div>
                            <div>
                                <div class="metric-value" style="font-size: 1.2em;">{key_tests}</div>
                                <div class="metric-label">测试用例数</div>
                            </div>
                            <div>
                                <div class="metric-value" style="font-size: 1.2em;">{key_success:.1f}%</div>
                                <div class="metric-label">成功率</div>
                            </div>
                            <div>
                                <div class="metric-value" style="font-size: 1.2em;">{key_quality:.0f}</div>
                                <div class="metric-label">代码质量</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 覆盖率分析 -->
            <div class="section">
                <h2>📈 测试覆盖率分析</h2>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <h3>总覆盖率</h3>
                        <div class="metric-value {coverage_class}">{total_coverage:.1f}%</div>
                        <div class="progress-bar">
                            <div class="progress-fill" style="width: {total_coverage}%"></div>
                        </div>
                        <div class="metric-label">目标: {threshold}%</div>
                    </div>
                    <div class="metric-card">
                        <h3>覆盖率分布</h3>
                        <div style="margin-top: 15px;">
                            {file_coverage_summary}
                        </div>
                    </div>
                </div>
                {uncovered_files_section}
            </div>

            <!-- 性能分析 -->
            <div class="section">
                <h2>⚡ 性能分析</h2>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <h3>执行时间</h3>
                        <div class="metric-value {performance_class}">{total_time:.1f}s</div>
                        <div class="metric-label">共 {test_count} 个测试用例</div>
                    </div>
                    <div class="metric-card">
                        <h3>成功率</h3>
                        <div class="metric-value {success_class}">{success_rate:.1f}%</div>
                        <div class="progress-bar">
                            <div class="progress-fill" style="width: {success_rate}%"></div>
                        </div>
                    </div>
                    <div class="metric-card">
                        <h3>内存使用</h3>
                        <div class="metric-value {memory_class}">{memory_usage:.1f}MB</div>
                        <div class="metric-label">峰值内存占用</div>
                    </div>
                    <div class="metric-card">
                        <h3>最慢测试</h3>
                        {slowest_tests_list}
                    </div>
                </div>
            </div>

            <!-- 质量分析 -->
            <div class="section">
                <h2>🎯 代码质量分析</h2>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <h3>质量分数</h3>
                        <div class="metric-value {quality_class}">{quality_score:.0f}</div>
                        <div class="metric-label">基于覆盖率、稳定性等综合评估</div>
                    </div>
                    <div class="metric-card">
                        <h3>测试稳定性</h3>
                        <div class="metric-value {stability_class}">{test_stability:.1f}%</div>
                        <div class="metric-label">多次运行结果一致性</div>
                    </div>
                    <div class="metric-card">
                        <h3>测试复杂度</h3>
                        <div class="metric-value {complexity_class}">{test_complexity:.0f}</div>
                        <div class="metric-label">测试用例复杂度评分</div>
                    </div>
                    <div class="metric-card">
                        <h3>重复检测</h3>
                        <div class="metric-value {duplicate_class}">{duplicate_coverage:.1f}%</div>
                        <div class="metric-label">重复测试用例比例</div>
                    </div>
                </div>
            </div>

            <!-- 建议和改进 -->
            <div class="section">
                <div class="recommendations">
                    <h3>💡 改进建议</h3>
                    {recommendations_html}
                </div>
            </div>
        </div>

        <div class="footer">
            <p>报告生成时间: {timestamp} | Excel MCP Server 监控系统</p>
        </div>
    </div>
</body>
</html>
        """

        # 准备数据
        overall_score = report.summary['overall_score']
        status = report.summary['status']
        key_metrics = report.summary['key_metrics']

        # 覆盖率数据
        coverage = report.coverage
        coverage_class = "good" if coverage.total_coverage >= self.coverage_threshold else "bad"

        # 文件覆盖率摘要
        file_coverage_summary = ""
        if coverage.file_coverage:
            sorted_files = sorted(coverage.file_coverage.items(), key=lambda x: x[1], reverse=True)[:5]
            file_coverage_summary = "<ul>"
            for filename, cov in sorted_files:
                file_coverage_summary += f"<li>{Path(filename).name}: {cov:.1f}%</li>"
            file_coverage_summary += "</ul>"
        else:
            file_coverage_summary = "<p>无覆盖率数据</p>"

        # 未覆盖文件
        uncovered_files_section = ""
        if coverage.uncovered_files:
            uncovered_files_section = f"""
            <div style="margin-top: 20px;">
                <h4>⚠️ 未覆盖的文件:</h4>
                <ul>
                    {"".join([f"<li>{file}</li>" for file in coverage.uncovered_files[:5]])}
                </ul>
            </div>
            """

        # 性能数据
        performance = report.performance
        performance_class = "good" if performance.total_time < 180 else ("warning" if performance.total_time < 300 else "bad")
        success_class = "good" if performance.success_rate >= 95 else ("warning" if performance.success_rate >= 90 else "bad")
        memory_class = "good" if performance.memory_usage_mb < 500 else ("warning" if performance.memory_usage_mb < 1000 else "bad")

        # 最慢测试列表
        slowest_tests_list = ""
        if performance.slowest_tests:
            slowest_tests_list = "<ul>"
            for test_name, duration in performance.slowest_tests[:3]:
                slowest_tests_list += f"<li>{Path(test_name).name}: {duration:.1f}s</li>"
            slowest_tests_list += "</ul>"
        else:
            slowest_tests_list = "<p>无数据</p>"

        # 质量数据
        quality = report.quality
        quality_class = "good" if quality.code_quality_score >= 80 else ("warning" if quality.code_quality_score >= 60 else "bad")
        stability_class = "good" if quality.test_stability >= 90 else ("warning" if quality.test_stability >= 80 else "bad")
        complexity_class = "good" if quality.test_complexity_score >= 70 else ("warning" if quality.test_complexity_score >= 50 else "bad")
        duplicate_class = "good" if quality.duplicate_coverage < 10 else ("warning" if quality.duplicate_coverage < 20 else "bad")

        # 建议HTML
        recommendations_html = "<ul>"
        for rec in report.recommendations:
            recommendations_html += f"<li>{rec}</li>"
        recommendations_html += "</ul>"

        return html_template.format(
            timestamp=report.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
            status=status,
            overall_score=overall_score,
            score=int(overall_score * 3.6),  # 转换为角度
            key_coverage=key_metrics['coverage'],
            key_tests=key_metrics['test_count'],
            key_success=key_metrics['success_rate'],
            key_quality=key_metrics['quality_score'],
            threshold=self.coverage_threshold,
            total_coverage=coverage.total_coverage,
            coverage_class=coverage_class,
            file_coverage_summary=file_coverage_summary,
            uncovered_files_section=uncovered_files_section,
            total_time=performance.total_time,
            test_count=performance.test_count,
            success_rate=performance.success_rate,
            performance_class=performance_class,
            success_class=success_class,
            memory_usage=performance.memory_usage_mb,
            memory_class=memory_class,
            slowest_tests_list=slowest_tests_list,
            quality_score=quality.code_quality_score,
            quality_class=quality_class,
            test_stability=quality.test_stability,
            stability_class=stability_class,
            test_complexity=quality.test_complexity,
            complexity_class=complexity_class,
            duplicate_coverage=quality.duplicate_coverage,
            duplicate_class=duplicate_class,
            recommendations_html=recommendations_html
        )

    def _generate_text_report(self, report: MonitoringReport) -> str:
        """生成文本报告"""
        text = f"""
{'='*80}
Excel MCP Server 监控报告
{'='*80}
生成时间: {report.timestamp.strftime("%Y-%m-%d %H:%M:%S")}
总体状态: {report.summary['status']} (评分: {report.summary['overall_score']:.1f}/100)

📊 关键指标:
- 测试覆盖率: {report.coverage.total_coverage:.1f}%
- 测试用例数: {report.performance.test_count}
- 成功率: {report.performance.success_rate:.1f}%
- 内存峰值: {report.performance.memory_usage_mb:.1f}MB
- 代码质量: {report.quality.code_quality_score:.0f}

📈 覆盖率分析:
- 总覆盖率: {report.coverage.total_coverage:.1f}% (目标: {self.coverage_threshold}%)
- 已覆盖文件: {len(report.coverage.file_coverage)}
- 未覆盖文件: {len(report.coverage.uncovered_files)}
"""

        if report.coverage.uncovered_files:
            text += f"  未覆盖文件: {', '.join(report.coverage.uncovered_files[:5])}\n"

        text += f"""
⚡ 性能分析:
- 执行时间: {report.performance.total_time:.1f}s
- 成功率: {report.performance.success_rate:.1f}%
- 内存使用: {report.performance.memory_usage_mb:.1f}MB (峰值)
"""

        if report.performance.slowest_tests:
            text += "  最慢的测试:\n"
            for test_name, duration in report.performance.slowest_tests[:3]:
                text += f"    - {test_name}: {duration:.1f}s\n"

        text += f"""
🎯 质量分析:
- 代码质量分数: {report.quality.code_quality_score:.0f}/100
- 测试稳定性: {report.quality.test_stability:.1f}%
- 测试复杂度: {report.quality.test_complexity:.0f}/100
- 重复测试: {report.quality.duplicate_coverage:.1f}%
"""

        if report.quality.flaky_tests:
            text += f"  不稳定测试: {', '.join(report.quality.flaky_tests[:3])}\n"

        text += "\n💡 改进建议:\n"
        for i, rec in enumerate(report.recommendations, 1):
            text += f"{i}. {rec}\n"

        text += f"\n{'='*80}\n"

        return text

    def run_full_monitoring(self, report_file: str = None) -> Dict[str, Any]:
        """运行完整监控"""
        logger.info("开始完整监控流程...")

        metrics = {}

        try:
            # 1. 覆盖率监控
            coverage_metrics = self.run_coverage_monitoring()
            metrics['coverage'] = coverage_metrics

            # 2. 性能监控 (包含内存监控)
            memory_monitor = MemoryMonitor()
            performance_metrics = self.run_performance_monitoring(memory_monitor)
            metrics['performance'] = performance_metrics
            metrics['memory'] = memory_monitor.stop_monitoring()

            # 3. 质量评估
            quality_metrics = self.run_quality_assessment(coverage_metrics, performance_metrics)
            metrics['quality'] = quality_metrics

            # 4. 生成建议
            temp_report = MonitoringReport(
                timestamp=datetime.now(),
                coverage=coverage_metrics,
                performance=performance_metrics,
                memory=metrics['memory'],
                quality=quality_metrics,
                recommendations=[],
                summary={}
            )
            recommendations = self.generate_recommendations(temp_report)
            metrics['recommendations'] = recommendations

            # 5. 生成报告
            if not report_file:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_file = self.reports_dir / f"monitoring-report-{timestamp}.html"

            report_text = self.generate_report(metrics, str(report_file))

            logger.info("监控流程完成")
            return {
                'success': True,
                'metrics': metrics,
                'report_file': str(report_file),
                'text_report': report_text
            }

        except Exception as e:
            logger.error(f"监控流程失败: {e}")
            return {
                'success': False,
                'error': str(e),
                'metrics': metrics
            }

    def run_continuous_monitoring(self, interval_minutes: int = 5):
        """运行连续监控"""
        logger.info(f"启动连续监控模式，间隔: {interval_minutes} 分钟")

        self._continuous_mode = True

        def signal_handler(signum, frame):
            logger.info("接收到停止信号，正在退出...")
            self._continuous_mode = False
            sys.exit(0)

        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)

        while self._continuous_mode:
            try:
                logger.info("开始新一轮监控...")
                result = self.run_full_monitoring()

                if result['success']:
                    metrics = result['metrics']
                    summary = metrics.get('summary', {})

                    # 输出到控制台
                    print(f"\n{datetime.now().strftime('%H:%M:%S')} - "
                          f"覆盖率: {metrics['coverage'].total_coverage:.1f}%, "
                          f"成功率: {metrics['performance'].success_rate:.1f}%, "
                          f"内存: {metrics['performance'].memory_usage_mb:.0f}MB, "
                          f"质量: {metrics['quality'].code_quality_score:.0f}")
                else:
                    logger.error(f"监控失败: {result.get('error', 'Unknown error')}")

                # 等待下一次监控
                if self._continuous_mode:
                    time.sleep(interval_minutes * 60)

            except KeyboardInterrupt:
                logger.info("用户中断，退出连续监控")
                break
            except Exception as e:
                logger.error(f"连续监控异常: {e}")
                time.sleep(interval_minutes * 60)

        self._continuous_mode = False
        logger.info("连续监控已停止")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="Excel MCP Server 监控和维护脚本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
    python scripts/monitor-and-maintain.py
    python scripts/monitor-and-maintain.py --coverage-only
    python scripts/monitor-and-maintain.py --continuous --threshold 90
    python scripts/monitor-and-maintain.py --no-html --report-file report.txt
        """
    )

    parser.add_argument(
        "--coverage-only",
        action="store_true",
        help="仅运行覆盖率监控"
    )

    parser.add_argument(
        "--performance-only",
        action="store_true",
        help="仅运行性能监控"
    )

    parser.add_argument(
        "--memory-only",
        action="store_true",
        help="仅运行内存监控"
    )

    parser.add_argument(
        "--quality-only",
        action="store_true",
        help="仅运行质量评估"
    )

    parser.add_argument(
        "--report-file",
        default=None,
        help="指定报告文件路径 (默认: reports/monitoring-report.html)"
    )

    parser.add_argument(
        "--threshold",
        type=float,
        default=85.0,
        help="覆盖率阈值 (默认: 85)"
    )

    parser.add_argument(
        "--no-html",
        action="store_true",
        help="不生成HTML报告，仅输出文本"
    )

    parser.add_argument(
        "--continuous",
        action="store_true",
        help="连续监控模式，每5分钟运行一次"
    )

    parser.add_argument(
        "--interval",
        type=int,
        default=5,
        help="连续监控间隔(分钟) (默认: 5)"
    )

    args = parser.parse_args()

    # 创建监控器
    monitor = TestMonitor(coverage_threshold=args.threshold)

    try:
        if args.continuous:
            # 连续监控模式
            monitor.run_continuous_monitoring(args.interval)
        elif args.coverage_only:
            # 仅覆盖率监控
            metrics = monitor.run_coverage_monitoring()
            print(f"测试覆盖率: {metrics.total_coverage:.1f}%")
            if metrics.uncovered_files:
                print(f"未覆盖文件: {', '.join(metrics.uncovered_files)}")
        elif args.performance_only:
            # 仅性能监控
            memory_monitor = MemoryMonitor()
            metrics = monitor.run_performance_monitoring(memory_monitor)
            print(f"执行时间: {metrics.total_time:.1f}s")
            print(f"测试数量: {metrics.test_count}")
            print(f"成功率: {metrics.success_rate:.1f}%")
            print(f"内存使用: {metrics.memory_usage_mb:.1f}MB")
        elif args.memory_only:
            # 仅内存监控 (需要配合其他监控)
            print("内存监控需要配合其他监控选项使用")
        elif args.quality_only:
            # 仅质量评估 (需要先运行其他监控获取数据)
            print("质量评估需要先运行覆盖率或性能监控")
        else:
            # 完整监控
            result = monitor.run_full_monitoring(args.report_file)

            if result['success']:
                print("\n" + "="*60)
                print("监控完成！")
                print("="*60)

                if not args.no_html and result.get('report_file'):
                    print(f"📄 HTML报告: {result['report_file']}")

                # 输出简要信息
                metrics = result['metrics']
                print(f"📊 覆盖率: {metrics['coverage'].total_coverage:.1f}%")
                print(f"⚡ 执行时间: {metrics['performance'].total_time:.1f}s")
                print(f"💾 内存使用: {metrics['performance'].memory_usage_mb:.1f}MB")
                print(f"🎯 质量分数: {metrics['quality'].code_quality_score:.0f}")
                print(f"✅ 建议: {len(metrics['recommendations'])} 条")
            else:
                print(f"❌ 监控失败: {result.get('error', 'Unknown error')}")
                return 1

    except KeyboardInterrupt:
        print("\n用户中断操作")
        return 0
    except Exception as e:
        logger.error(f"程序异常: {e}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())