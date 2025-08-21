"""
专门测试Excel MCP Server优化功能的pytest测试套件
验证缓存机制、中文字符处理和统一错误处理
"""

import pytest
import time
import tempfile
from pathlib import Path

from src.core.excel_writer import ExcelWriter
from src.core.excel_manager import ExcelManager
from src.utils.formula_cache import get_formula_cache
from src.server import excel_create_sheet, excel_evaluate_formula
from src.models.types import OperationResult


class TestOptimizationFeatures:
    """测试优化功能"""

    def test_formula_cache_performance_boost(self, sample_excel_file):
        """测试公式计算缓存的性能提升"""
        writer = ExcelWriter(sample_excel_file)
        cache = get_formula_cache()
        cache.clear()

        # 测试公式
        formula = "SUM(D2:D5)"

        # 第一次计算（无缓存）
        start_time = time.time()
        result1 = writer.evaluate_formula(formula)
        time1 = time.time() - start_time

        assert result1.success is True
        assert result1.data == 35000  # 从conftest.py中的样本数据计算：8000+9000+8500+9500

        # 第二次计算（有缓存）
        start_time = time.time()
        result2 = writer.evaluate_formula(formula)
        time2 = time.time() - start_time

        assert result2.success is True
        assert result2.data == result1.data

        # 验证缓存效果（第二次应该明显更快）
        assert time2 < time1 * 0.5, f"缓存未生效：首次{time1*1000:.2f}ms，缓存{time2*1000:.2f}ms"

        # 验证缓存统计
        cache_stats = cache.get_stats()
        assert cache_stats['hit_count'] > 0, "缓存命中数应该大于0"
        assert cache_stats['cache_size'] > 0, "缓存大小应该大于0"

    def test_cache_metadata_in_result(self, sample_excel_file):
        """测试缓存相关的元数据是否正确添加到结果中"""
        writer = ExcelWriter(sample_excel_file)
        cache = get_formula_cache()
        cache.clear()

        # 第一次计算
        result1 = writer.evaluate_formula("AVERAGE(D2:D5)")
        assert result1.success is True
        assert 'cached' in result1.metadata
        assert result1.metadata['cached'] is False  # 首次计算不是缓存

        # 第二次计算
        result2 = writer.evaluate_formula("AVERAGE(D2:D5)")
        assert result2.success is True
        assert 'cached' in result2.metadata
        assert result2.metadata['cached'] is True  # 第二次应该是缓存

        # 验证缓存统计信息包含在元数据中
        assert 'cache_stats' in result2.metadata

    def test_chinese_sheet_name_handling(self, sample_excel_file):
        """测试中文工作表名称处理"""
        manager = ExcelManager(sample_excel_file)

        # 测试各种中文名称场景
        test_cases = [
            ("数据分析", "数据分析"),  # 普通中文
            ("销售报表2024", "销售报表2024"),  # 中英文数字混合
            ("测试/数据", "测试_数据"),  # 特殊字符替换
            ("  空格测试  ", "空格测试"),  # 空格处理
            ("Sheet*Test", "Sheet_Test"),  # 非法字符替换
        ]

        for input_name, expected_output in test_cases:
            result = manager.create_sheet(input_name)
            assert result.success is True, f"创建工作表失败：{input_name}"
            assert result.data.name == expected_output, f"名称处理不正确：期望'{expected_output}'，实际'{result.data.name}'"

    def test_chinese_sheet_long_name_handling(self, sample_excel_file):
        """测试超长中文工作表名称处理"""
        manager = ExcelManager(sample_excel_file)

        # Excel工作表名称限制是31个字符
        long_name = "很长的中文工作表名称测试超过三十一个字符的情况处理"
        result = manager.create_sheet(long_name)

        if result.success:
            # 如果成功，名称应该被适当处理
            assert len(result.data.name) <= 31, "工作表名称长度应该不超过31个字符"
        else:
            # 如果失败，应该有合理的错误信息
            assert "长度" in result.error or "字符" in result.error

    def test_chinese_sheet_empty_name_error(self, sample_excel_file):
        """测试空工作表名称的错误处理"""
        manager = ExcelManager(sample_excel_file)

        result = manager.create_sheet("")
        assert result.success is False
        assert "空" in result.error or "不能为空" in result.error

    def test_unified_error_handling_structure(self, sample_excel_file):
        """测试统一错误处理的返回结构"""
        # 测试MCP接口的错误处理
        result = excel_create_sheet("不存在的文件.xlsx", "测试工作表")

        assert isinstance(result, dict), "MCP接口应该返回字典"
        assert 'success' in result, "结果应该包含success字段"
        assert result['success'] is False, "不存在文件的操作应该失败"
        assert 'error' in result, "失败结果应该包含error字段"

        # 检查错误格式
        error = result['error']
        if isinstance(error, dict):
            assert 'code' in error or 'message' in error, "错误信息应该包含code或message"

    def test_unified_error_handling_with_formula(self):
        """测试公式计算的统一错误处理"""
        # 测试不存在文件的公式计算
        result = excel_evaluate_formula("不存在的文件.xlsx", "SUM(A1:A10)")

        assert isinstance(result, dict), "MCP接口应该返回字典"
        assert result['success'] is False, "不存在文件的操作应该失败"
        assert 'error' in result, "失败结果应该包含error字段"

    def test_error_handling_consistency_between_layers(self, sample_excel_file):
        """测试不同层级错误处理的一致性"""
        # 核心层API（返回OperationResult）
        writer = ExcelWriter(sample_excel_file)
        core_result = writer.evaluate_formula("")  # 空公式

        assert isinstance(core_result, OperationResult)
        assert core_result.success is False
        assert "空" in core_result.error

        # MCP接口层（返回字典）
        mcp_result = excel_evaluate_formula(sample_excel_file, "")  # 空公式

        assert isinstance(mcp_result, dict)
        assert mcp_result['success'] is False
        # 两层的错误信息应该类似
        assert "空" in str(mcp_result['error'])

    def test_cache_isolation_between_files(self, temp_dir):
        """测试不同文件之间的缓存隔离"""
        # 创建两个不同的Excel文件
        file1 = temp_dir / "test1.xlsx"
        file2 = temp_dir / "test2.xlsx"

        ExcelManager.create_file(str(file1), ["Sheet1"])
        ExcelManager.create_file(str(file2), ["Sheet1"])

        writer1 = ExcelWriter(str(file1))
        writer2 = ExcelWriter(str(file2))

        # 在两个文件中写入不同数据
        writer1.update_range("A1:A3", [[10], [20], [30]])
        writer2.update_range("A1:A3", [[100], [200], [300]])

        cache = get_formula_cache()
        cache.clear()

        # 计算相同公式但在不同文件
        result1 = writer1.evaluate_formula("SUM(A1:A3)")
        result2 = writer2.evaluate_formula("SUM(A1:A3)")

        assert result1.success is True
        assert result2.success is True
        assert result1.data == 60  # 10+20+30
        assert result2.data == 600  # 100+200+300

        # 验证缓存正确隔离了不同文件的结果
        assert result1.data != result2.data

    def test_cache_ttl_behavior(self, sample_excel_file):
        """测试缓存TTL（生存时间）行为"""
        writer = ExcelWriter(sample_excel_file)
        cache = get_formula_cache()
        cache.clear()

        # 设置短TTL用于测试（注意：这里只是验证接口，不修改全局TTL）
        formula = "MAX(D2:D5)"

        # 第一次计算
        result1 = writer.evaluate_formula(formula)
        assert result1.success is True

        # 立即第二次计算，应该命中缓存
        result2 = writer.evaluate_formula(formula)
        assert result2.success is True
        assert result2.metadata['cached'] is True

        # 验证缓存统计
        stats = cache.get_stats()
        assert stats['hit_count'] > 0
        assert stats['ttl'] > 0  # 确认TTL设置存在

    def test_comprehensive_workflow(self, temp_dir):
        """综合工作流测试：组合使用所有优化功能"""
        # 创建测试文件
        file_path = temp_dir / "综合测试.xlsx"

        # 1. 使用中文文件名和工作表
        result = ExcelManager.create_file(str(file_path), ["数据汇总"])
        assert result.success is True

        manager = ExcelManager(str(file_path))

        # 2. 创建多个中文工作表
        chinese_sheets = ["销售数据", "成本分析", "利润/统计"]
        for sheet_name in chinese_sheets:
            result = manager.create_sheet(sheet_name)
            assert result.success is True

        # 3. 写入数据并测试公式缓存
        writer = ExcelWriter(str(file_path))
        test_data = [
            ["项目", "收入", "成本"],
            ["项目A", 100000, 80000],
            ["项目B", 150000, 100000],
            ["项目C", 120000, 90000]
        ]
        writer.update_range("A1:C4", test_data)

        cache = get_formula_cache()
        cache.clear()

        # 4. 测试多个公式的缓存效果
        formulas = [
            "SUM(B2:B4)",      # 总收入
            "SUM(C2:C4)",      # 总成本
            "SUM(B2:B4)-SUM(C2:C4)"  # 总利润
        ]

        # 首次计算
        first_results = []
        first_times = []
        for formula in formulas:
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time

            assert result.success is True
            first_results.append(result.data)
            first_times.append(elapsed)

        # 二次计算（缓存）
        second_times = []
        for i, formula in enumerate(formulas):
            start_time = time.time()
            result = writer.evaluate_formula(formula)
            elapsed = time.time() - start_time

            assert result.success is True
            assert result.data == first_results[i]  # 结果一致
            assert result.metadata['cached'] is True  # 确认使用了缓存
            second_times.append(elapsed)

        # 5. 验证整体性能提升
        total_first = sum(first_times)
        total_second = sum(second_times)

        if total_first > 0:
            improvement = (total_first - total_second) / total_first * 100
            assert improvement > 0, "缓存应该带来性能提升"

        # 6. 验证错误处理
        error_result = writer.evaluate_formula("")
        assert error_result.success is False
        assert "空" in error_result.error

        print(f"✅ 综合测试完成！缓存提升：{improvement:.1f}%")


class TestRegressionPrevention:
    """回归测试：确保优化不会破坏现有功能"""

    def test_basic_functionality_unchanged(self, sample_excel_file):
        """确保基本功能没有被破坏"""
        # 基本读写功能
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("A1", [["测试数据"]])
        assert result.success is True

        # 基本公式功能
        result = writer.evaluate_formula("SUM(D2:D5)")
        assert result.success is True
        assert result.data > 0

    def test_existing_error_scenarios_still_work(self, sample_excel_file):
        """确保现有的错误场景仍然正常工作"""
        writer = ExcelWriter(sample_excel_file)

        # 测试明确指定不存在的工作表（使用工作表!范围语法）
        result = writer.update_range("这个工作表确实不存在!A1:A1", [["测试"]])
        assert result.success is False, "应该因为工作表不存在而失败"

        # 无效范围格式
        result = writer.update_range("INVALID_RANGE_FORMAT", [["测试"]])
        assert result.success is False, "应该因为范围格式无效而失败"

        # 空公式
        result = writer.evaluate_formula("")
        assert result.success is False, "空公式应该失败"
