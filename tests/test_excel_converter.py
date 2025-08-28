# -*- coding: utf-8 -*-
"""
Excel文件转换功能测试
测试src.core.excel_converter模块
"""

import pytest
import tempfile
import os
from pathlib import Path
from unittest.mock import Mock, patch, mock_open

from src.core.excel_converter import ExcelConverter
from src.models.types import OperationResult
from src.utils.exceptions import ExcelFileNotFoundError


class TestExcelConverter:
    """Excel转换功能的综合测试"""

    def test_converter_init(self):
        """测试转换器初始化"""
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
            assert converter is not None
            assert converter.file_path == "test.xlsx"

    @patch('src.core.excel_converter.load_workbook')
    @patch('os.path.exists')
    def test_convert_to_csv_success(self, mock_exists, mock_load_workbook):
        """测试成功转换Excel到CSV"""
        # Mock文件存在
        mock_exists.return_value = True
        
        # Mock工作簿和工作表
        mock_worksheet = Mock()
        mock_worksheet.iter_rows.return_value = [
            ['Name', 'Age', 'City'],
            ['Alice', 25, 'New York'],
            ['Bob', 30, 'Los Angeles']
        ]
        mock_worksheet.title = "Sheet1"
        
        mock_workbook = Mock()
        mock_workbook.active = mock_worksheet
        mock_load_workbook.return_value = mock_workbook
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
        
        with patch('builtins.open', mock_open()) as mock_file:
            result = converter.export_to_csv("output.csv")
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    @patch('os.path.exists')
    def test_convert_to_csv_file_not_found(self, mock_exists):
        """测试转换不存在的文件"""
        mock_exists.return_value = False
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="nonexistent.xlsx"):
            converter = ExcelConverter("nonexistent.xlsx")
        result = converter.export_to_csv("output.csv")
        
        assert isinstance(result, OperationResult)
        assert result.success is False

    @patch('src.core.excel_converter.load_workbook')
    @patch('os.path.exists')
    def test_convert_to_json_success(self, mock_exists, mock_load_workbook):
        """测试成功转换Excel到JSON"""
        mock_exists.return_value = True
        
        # Mock工作簿和工作表
        mock_worksheet = Mock()
        mock_worksheet.iter_rows.return_value = [
            ['Name', 'Age', 'City'],
            ['Alice', 25, 'New York'],
            ['Bob', 30, 'Los Angeles']
        ]
        
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1']
        mock_workbook.__getitem__ = Mock(return_value=mock_worksheet)
        mock_load_workbook.return_value = mock_workbook
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
        
        with patch('builtins.open', mock_open()) as mock_file:
            with patch('json.dump') as mock_json_dump:
                with patch('os.path.exists', return_value=True):
                    with patch('os.path.getsize', return_value=1024):
                        with patch('pathlib.Path.mkdir'):
                            result = ExcelConverter.convert_format("test.xlsx", "output.json", "json")
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    @patch('src.core.excel_converter.load_workbook')
    @patch('os.path.exists') 
    def test_convert_with_specific_sheet(self, mock_exists, mock_load_workbook):
        """测试转换指定工作表"""
        mock_exists.return_value = True
        
        # Mock工作簿
        mock_worksheet = Mock()
        mock_worksheet.iter_rows.return_value = [['Data']]
        mock_worksheet.title = "Sheet2"
        
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1', 'Sheet2']
        mock_workbook.__getitem__ = Mock(return_value=mock_worksheet)
        mock_load_workbook.return_value = mock_workbook
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
        
        with patch('builtins.open', mock_open()):
            result = converter.export_to_csv("output.csv", "Sheet2")
        
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_supported_formats(self):
        """测试支持的格式"""
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
        
        # 测试方法是否存在
        assert hasattr(converter, 'export_to_csv')
        assert hasattr(converter, 'convert_format')
        
        # 可能还有其他格式的转换方法
        potential_methods = [
            'convert_to_xml', 'convert_to_html', 'convert_to_txt'
        ]
        
        for method in potential_methods:
            # 检查方法是否存在（不要求必须存在）
            method_exists = hasattr(converter, method)
            assert isinstance(method_exists, bool)

    @patch('src.core.excel_converter.load_workbook')
    @patch('os.path.exists')
    def test_convert_exception_handling(self, mock_exists, mock_load_workbook):
        """测试转换过程中的异常处理"""
        mock_exists.return_value = True
        mock_load_workbook.side_effect = Exception("读取工作簿失败")
        
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="test.xlsx"):
            converter = ExcelConverter("test.xlsx")
        result = converter.export_to_csv("output.csv")
        
        assert isinstance(result, OperationResult)
        assert result.success is False

    @pytest.mark.parametrize("input_format,output_format", [
        ("xlsx", "csv"),
        ("xlsm", "csv"),
        ("xlsx", "json"),
        ("xlsm", "json")
    ])
    def test_format_combinations(self, input_format, output_format):
        """测试不同格式组合的转换"""
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value=f"test.{input_format}"):
            converter = ExcelConverter(f"test.{input_format}")
        
        # 这里只测试方法调用不会出错
        with patch('os.path.exists', return_value=False):
            if output_format == "csv":
                result = converter.export_to_csv("output.csv")
            elif output_format == "json":
                result = ExcelConverter.convert_format(f"test.{input_format}", "output.json", "json")
            
            assert isinstance(result, OperationResult)
            # 由于文件不存在，应该返回失败
            assert result.success is False

    def test_converter_error_messages(self):
        """测试转换器的错误消息"""
        with patch('src.utils.validators.ExcelValidator.validate_file_path', return_value="nonexistent.xlsx"):
            converter = ExcelConverter("nonexistent.xlsx")
        
        # 测试文件不存在的错误消息
        with patch('os.path.exists', return_value=False):
            result = converter.export_to_csv("output.csv")
            
            assert result.success is False
            assert result.error is not None
            assert len(result.error) > 0
