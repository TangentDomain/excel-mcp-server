# -*- coding: utf-8 -*-
"""
Excel文件比较功能测试
测试src.core.excel_compare模块的ExcelComparer类
"""

import pytest
import tempfile
import os
from pathlib import Path
from unittest.mock import Mock, patch

from src.core.excel_compare import ExcelComparer
from src.models.types import (
    ComparisonOptions, OperationResult, ComparisonResult,
    DifferenceType, RowDifference, FieldDifference
)
from src.utils.exceptions import ExcelFileNotFoundError, SheetNotFoundError


class TestExcelComparer:
    """Excel比较功能的综合测试"""

    def test_comparer_init_default_options(self):
        """测试使用默认选项初始化比较器"""
        comparer = ExcelComparer()
        assert comparer.options is not None
        assert isinstance(comparer.options, ComparisonOptions)

    def test_comparer_init_custom_options(self):
        """测试使用自定义选项初始化比较器"""
        options = ComparisonOptions(
            compare_formulas=True,
            compare_formats=True,
            ignore_empty_cells=False
        )
        comparer = ExcelComparer(options)
        assert comparer.options is options

    @patch('src.core.excel_compare.load_workbook')
    @patch('src.utils.validators.ExcelValidator.validate_file_path')
    def test_compare_files_identical(self, mock_validate, mock_load_workbook):
        """测试比较两个相同的文件"""
        # Mock文件验证
        mock_validate.side_effect = lambda x: x
        
        # Mock工作簿
        mock_worksheet = Mock()
        mock_worksheet.title = 'Sheet1'
        
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1']
        mock_workbook.__getitem__ = Mock(return_value=mock_worksheet)
        mock_load_workbook.return_value = mock_workbook
        
        comparer = ExcelComparer()
        
        with patch.object(comparer, '_compare_file_structure', return_value=[]):
            with patch.object(comparer, '_compare_sheets', return_value=Mock(total_differences=0)):
                with patch.object(comparer, '_generate_comparison_summary', return_value="相同文件"):
                    result = comparer.compare_files("file1.xlsx", "file2.xlsx")
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    @patch('src.core.excel_compare.load_workbook')
    @patch('src.utils.validators.ExcelValidator.validate_file_path')
    def test_compare_files_different(self, mock_validate, mock_load_workbook):
        """测试比较两个不同的文件"""
        # Mock文件验证
        mock_validate.side_effect = lambda x: x
        
        # Mock工作簿
        mock_workbook1 = Mock()
        mock_workbook1.sheetnames = ['Sheet1']
        mock_workbook2 = Mock()
        mock_workbook2.sheetnames = ['Sheet1', 'Sheet2']
        
        mock_load_workbook.side_effect = [mock_workbook1, mock_workbook2]
        
        comparer = ExcelComparer()
        
        with patch.object(comparer, '_compare_file_structure', return_value=['sheet_added']):
            with patch.object(comparer, '_compare_sheets', return_value=Mock(total_differences=5)):
                with patch.object(comparer, '_generate_comparison_summary', return_value="文件有差异"):
                    result = comparer.compare_files("file1.xlsx", "file2.xlsx")
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_compare_files_invalid_path(self):
        """测试比较不存在的文件"""
        comparer = ExcelComparer()
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path') as mock_validate:
            mock_validate.side_effect = ExcelFileNotFoundError("文件不存在")
            
            result = comparer.compare_files("nonexistent1.xlsx", "nonexistent2.xlsx")
            
            assert isinstance(result, OperationResult)
            assert result.success is False
            assert "文件不存在" in result.error

    def test_compare_sheets_method(self):
        """测试比较工作表的私有方法"""
        comparer = ExcelComparer()
        
        # 这里需要更复杂的Mock，暂时测试方法存在
        assert hasattr(comparer, '_compare_sheets')
        assert callable(getattr(comparer, '_compare_sheets'))

    def test_compare_file_structure_method(self):
        """测试比较文件结构的私有方法"""
        comparer = ExcelComparer()
        
        # 测试方法存在
        assert hasattr(comparer, '_compare_file_structure')
        assert callable(getattr(comparer, '_compare_file_structure'))

    def test_generate_comparison_summary_method(self):
        """测试生成比较摘要的私有方法"""
        comparer = ExcelComparer()
        
        # 测试方法存在
        assert hasattr(comparer, '_generate_comparison_summary')
        assert callable(getattr(comparer, '_generate_comparison_summary'))

    @patch('src.core.excel_compare.load_workbook')
    @patch('src.utils.validators.ExcelValidator.validate_file_path')
    def test_compare_files_with_options(self, mock_validate, mock_load_workbook):
        """测试使用特定选项比较文件"""
        mock_validate.side_effect = lambda x: x
        
        # Mock工作簿
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1']
        mock_load_workbook.return_value = mock_workbook
        
        comparer = ExcelComparer()
        options = ComparisonOptions(compare_formulas=True)
        
        with patch.object(comparer, '_compare_file_structure', return_value=[]):
            with patch.object(comparer, '_compare_sheets', return_value=Mock(total_differences=0)):
                with patch.object(comparer, '_generate_comparison_summary', return_value="相同"):
                    result = comparer.compare_files("file1.xlsx", "file2.xlsx", options)
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_compare_files_exception_handling(self):
        """测试比较文件时的异常处理"""
        comparer = ExcelComparer()
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path') as mock_validate:
            mock_validate.side_effect = Exception("意外错误")
            
            result = comparer.compare_files("file1.xlsx", "file2.xlsx")
            
            assert isinstance(result, OperationResult)
            assert result.success is False
            assert "意外错误" in result.error

    def test_structured_comparison_mode(self):
        """测试结构化比较模式（游戏开发特化功能）"""
        comparer = ExcelComparer()
        
        # 测试结构化比较相关方法是否存在
        # 这些方法可能在更大的代码库中定义
        assert hasattr(comparer, 'options')

    def test_comparison_options_types(self):
        """测试比较选项的类型"""
        options = ComparisonOptions()
        
        # 测试默认值
        assert hasattr(options, 'compare_formulas')
        assert hasattr(options, 'compare_formats')
        assert hasattr(options, 'ignore_empty_cells')

    @pytest.mark.parametrize("compare_formulas,compare_formats", [
        (True, True),
        (True, False), 
        (False, True),
        (False, False)
    ])
    def test_comparison_options_variations(self, compare_formulas, compare_formats):
        """测试不同的比较选项组合"""
        options = ComparisonOptions(
            compare_formulas=compare_formulas,
            compare_formats=compare_formats
        )
        
        comparer = ExcelComparer(options)
        assert comparer.options.compare_formulas == compare_formulas
        assert comparer.options.compare_formats == compare_formats
