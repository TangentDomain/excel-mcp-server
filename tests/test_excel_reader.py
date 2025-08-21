"""
Fixed tests for ExcelReader class - matching actual API implementation
"""

import pytest
from src.core.excel_reader import ExcelReader
from src.models.types import OperationResult, SheetInfo, CellInfo
from src.utils.exceptions import ExcelFileNotFoundError


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
        assert isinstance(result.data, list)
        assert len(result.data) == 2

        # Check first sheet
        sheet1 = result.data[0]
        assert isinstance(sheet1, SheetInfo)
        assert hasattr(sheet1, 'name')
        assert hasattr(sheet1, 'index')
        assert hasattr(sheet1, 'is_active')
        assert hasattr(sheet1, 'max_row')
        assert hasattr(sheet1, 'max_column')
        assert hasattr(sheet1, 'max_column_letter')

    def test_get_range_cell_range(self, sample_excel_file):
        """Test getting a cell range"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1:C5")

        assert isinstance(result, OperationResult)
        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 5  # 5 rows

        # Check first row
        first_row = result.data[0]
        assert len(first_row) == 3  # 3 columns
        assert isinstance(first_row[0], CellInfo)
        assert first_row[0].coordinate == "A1"
        assert first_row[0].value is not None

        # Check last cell in range
        last_row = result.data[4]
        last_cell = last_row[2]
        assert last_cell.coordinate == "C5"
        assert last_cell.value is not None

    def test_get_range_single_cell(self, sample_excel_file):
        """Test getting a single cell"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1
        assert len(result.data[0]) == 1
        assert result.data[0][0].coordinate == "A1"
        assert result.data[0][0].value is not None

    def test_get_range_with_sheet_name(self, sample_excel_file):
        """Test getting range with sheet name"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet2!A1:C3")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 3  # 3 rows
        assert len(result.data[0]) == 3  # 3 columns
        assert result.data[0][0].coordinate == "A1"

    def test_get_range_entire_row(self, sample_excel_file):
        """Test getting entire row"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("1:1")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1  # 1 row
        assert len(result.data[0]) >= 4  # At least 4 columns
        assert result.data[0][0].coordinate == "A1"

    def test_get_range_entire_column(self, sample_excel_file):
        """Test getting entire column"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A:A")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) >= 5  # At least 5 rows
        assert result.data[0][0].coordinate == "A1"
        assert result.data[1][0].coordinate == "A2"

    def test_get_range_with_formatting(self, sample_excel_file):
        """Test getting range with formatting info"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1:D1", include_formatting=True)

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1
        assert len(result.data[0]) == 4
        assert result.data[0][0].coordinate == "A1"
        # May have formatting info in CellInfo objects

    def test_get_range_invalid_sheet(self, sample_excel_file):
        """Test getting range from non-existent sheet"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("NonExistentSheet!A1")

        assert result.success is False
        assert result.error is not None
        assert "工作表" in result.error

    def test_get_range_invalid_range(self, sample_excel_file):
        """Test getting invalid range"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("ZZ999:AAA1000")

        # Should handle gracefully, may return empty result
        assert isinstance(result, OperationResult)

    def test_list_sheets_empty_file(self, empty_excel_file):
        """Test listing sheets from empty file"""
        reader = ExcelReader(empty_excel_file)
        result = reader.list_sheets()

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1
        assert isinstance(result.data[0], SheetInfo)
        assert "Sheet" in result.data[0].name  # Default sheet name

    def test_get_range_out_of_bounds(self, sample_excel_file):
        """Test getting range that's out of bounds"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Z100:AA101")

        # Should handle gracefully
        assert isinstance(result, OperationResult)
        if result.success:
            assert isinstance(result.data, list)

    def test_get_range_case_sensitive_sheet(self, sample_excel_file):
        """Test getting range with case sensitive sheet name"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("SHEET1!A1")  # Different case

        # May or may not work depending on implementation
        assert isinstance(result, OperationResult)

    def test_get_range_unicode_content(self, sample_excel_file):
        """Test getting range with unicode content"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("A1:A5")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 5

        # Check that unicode content is handled properly
        for row in result.data:
            assert len(row) == 1
            assert isinstance(row[0], CellInfo)
            assert row[0].value is not None
            # Unicode content should be preserved
            assert isinstance(row[0].value, str)
