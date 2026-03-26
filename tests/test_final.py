# -*- coding: utf-8 -*-
"""
最后的测试覆盖
"""

import pytest
from openpyxl import Workbook
from src.api.excel_operations import ExcelOperations


class TestFinalCoverage:
    """最终覆盖测试"""

    @pytest.fixture
    def test_file(self, temp_dir):
        file_path = temp_dir / "final.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws['A1'] = "Test"
        wb.save(file_path)
        return str(file_path)

    def test_list_sheets_again(self, test_file):
        """测试列出工作表"""
        result = ExcelOperations.list_sheets(test_file)
        assert result['success'] is True

    def test_get_file_info_again(self, test_file):
        """测试获取文件信息"""
        result = ExcelOperations.get_file_info(test_file)
        assert result['success'] is True

    def test_get_range_again(self, test_file):
        """测试获取范围"""
        result = ExcelOperations.get_range(test_file, "Sheet1!A1")
        assert result['success'] is True

    def test_get_headers_again(self, test_file):
        """测试获取表头"""
        result = ExcelOperations.get_headers(test_file, "Sheet1")
        assert result['success'] is True
