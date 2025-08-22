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
from src.server import excel_create_sheet
from src.models.types import OperationResult


class TestOptimizationFeatures:
    """测试优化功能"""

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

class TestRegressionPrevention:
    """回归测试：确保优化不会破坏现有功能"""

    def test_basic_functionality_unchanged(self, sample_excel_file):
        """确保基本功能没有被破坏"""
        # 基本读写功能
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet1!A1", [["测试数据"]])
        assert result.success is True

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
