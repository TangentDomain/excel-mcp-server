"""
Tests for ExcelReader class
"""

import pytest
from pathlib import Path
from src.core.excel_reader import ExcelReader
from src.models.types import OperationResult
from src.utils.exceptions import SheetNotFoundError, FileNotFoundError


class TestExcelReader:
    """Test cases for ExcelReader class"""
    
    def test_init_valid_file(self, sample_excel_file):
        """Test initialization with valid file"""
        reader = ExcelReader(sample_excel_file)
        assert reader.file_path == sample_excel_file
    
    def test_init_invalid_file(self):
        """Test initialization with invalid file"""
        with pytest.raises(FileNotFoundError):
            ExcelReader("nonexistent_file.xlsx")
    
    def test_list_sheets(self, sample_excel_file):
        """Test listing sheets"""
        reader = ExcelReader(sample_excel_file)
        result = reader.list_sheets()
        
        assert isinstance(result, OperationResult)
        assert result.success is True
        assert result.data is not None
        assert len(result.data.sheets) == 2
        assert "Sheet1" in result.data.sheets
        assert "Sheet2" in result.data.sheets
        assert result.data.active_sheet == "Sheet1"
    
    def test_get_range_cell_range(self, sample_excel_file):
        """Test getting a cell range"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1:C5")
        
        assert isinstance(result, OperationResult)
        assert result.success is True
        assert result.data is not None
        assert len(result.data.data) == 5  # 5 rows
        assert result.data.data[0][0] == "姓名"  # First cell
        assert result.data.data[4][2] == "人事部"  # Last cell in range
    
    def test_get_range_single_cell(self, sample_excel_file):
        """Test getting a single cell"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1")
        
        assert result.success is True
        assert result.data.data[0][0] == "姓名"
    
    def test_get_range_with_sheet_name(self, sample_excel_file):
        """Test getting range with sheet name"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet2!A1:C3")
        
        assert result.success is True
        assert result.data.data[0][0] == "产品"
        assert result.data.data[2][2] == 30
    
    def test_get_range_entire_row(self, sample_excel_file):
        """Test getting entire row"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("1:1")
        
        assert result.success is True
        assert result.data.data[0][0] == "姓名"
        assert result.data.data[0][1] == "年龄"
    
    def test_get_range_entire_column(self, sample_excel_file):
        """Test getting entire column"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A:A")
        
        assert result.success is True
        assert len(result.data.data) >= 5
        assert result.data.data[0][0] == "姓名"
        assert result.data.data[1][0] == "张三"
    
    def test_get_range_with_formatting(self, sample_excel_file):
        """Test getting range with formatting info"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1:D1", include_formatting=True)
        
        assert result.success is True
        assert result.data.range_info is not None
        # Header cells should have formatting
        assert len(result.data.data) == 1
        assert result.data.data[0][0] == "姓名"
    
    def test_get_range_invalid_sheet(self, sample_excel_file):
        """Test getting range from non-existent sheet"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("NonExistentSheet!A1")
        
        assert result.success is False
        assert "工作表" in result.error
    
    def test_get_range_invalid_range(self, sample_excel_file):
        """Test getting invalid range"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("ZZ999:AAA1000")
        
        # Should handle gracefully, return empty or error
        assert isinstance(result, OperationResult)
    
    def test_get_sheet_dimensions(self, sample_excel_file):
        """Test getting sheet dimensions"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_sheet_dimensions("Sheet1")
        
        assert result.success is True
        assert result.data is not None
        assert result.data.row_count >= 5
        assert result.data.column_count >= 4
    
    def test_get_sheet_info(self, sample_excel_file):
        """Test getting sheet info"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_sheet_info("Sheet1")
        
        assert result.success is True
        assert result.data is not None
        assert result.data.name == "Sheet1"
        assert result.data.row_count >= 5
        assert result.data.column_count >= 4
    
    def test_get_sheet_info_nonexistent(self, sample_excel_file):
        """Test getting info for non-existent sheet"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_sheet_info("NonExistentSheet")
        
        assert result.success is False
        assert "工作表" in result.error
    
    def test_get_cell_value(self, sample_excel_file):
        """Test getting single cell value"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_cell_value("Sheet1", "A1")
        
        assert result.success is True
        assert result.data == "姓名"
    
    def test_get_cell_value_with_formula(self, sample_excel_file):
        """Test getting cell value with formula"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_cell_value("Sheet1", "E2")
        
        assert result.success is True
        # Should return calculated value, not formula
        assert isinstance(result.data, (int, float))
    
    def test_get_cell_value_nonexistent_cell(self, sample_excel_file):
        """Test getting value from non-existent cell"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_cell_value("Sheet1", "ZZ999")
        
        assert result.success is True
        assert result.data is None
    
    def test_list_sheets_empty_file(self, empty_excel_file):
        """Test listing sheets from empty file"""
        reader = ExcelReader(empty_excel_file)
        result = reader.list_sheets()
        
        assert result.success is True
        assert len(result.data.sheets) == 1
        assert "Sheet" in result.data.sheets[0]  # Default sheet name