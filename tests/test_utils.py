"""
Tests for utility modules
"""

import pytest
from src.utils.validators import ExcelValidator
from src.utils.parsers import RangeParser
from src.utils.exceptions import (
    FileNotFoundError, 
    SheetNotFoundError, 
    DataValidationError,
    InvalidRangeError
)


class TestExcelValidator:
    """Test cases for ExcelValidator class"""
    
    def test_validate_file_path_valid(self, sample_excel_file):
        """Test validating valid file path"""
        result = ExcelValidator.validate_file_path(sample_excel_file)
        assert result == sample_excel_file
    
    def test_validate_file_path_nonexistent(self):
        """Test validating non-existent file"""
        with pytest.raises(FileNotFoundError):
            ExcelValidator.validate_file_path("nonexistent_file.xlsx")
    
    def test_validate_file_path_invalid_extension(self, temp_dir):
        """Test validating file with invalid extension"""
        file_path = temp_dir / "test.txt"
        file_path.touch()  # Create the file
        
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_file_path(str(file_path))
    
    def test_validate_file_for_creation_valid(self, temp_dir):
        """Test validating file path for creation"""
        file_path = temp_dir / "test.xlsx"
        result = ExcelValidator.validate_file_for_creation(str(file_path))
        assert result == str(file_path)
    
    def test_validate_file_for_creation_invalid_extension(self, temp_dir):
        """Test validating invalid extension for creation"""
        file_path = temp_dir / "test.txt"
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_file_for_creation(str(file_path))
    
    def test_validate_sheet_name_valid(self):
        """Test validating valid sheet name"""
        result = ExcelValidator.validate_sheet_name("Sheet1")
        assert result == "Sheet1"
    
    def test_validate_sheet_name_empty(self):
        """Test validating empty sheet name"""
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_sheet_name("")
    
    def test_validate_sheet_name_invalid_chars(self):
        """Test validating sheet name with invalid characters"""
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_sheet_name("Sheet\\1")
    
    def test_validate_range_expression_valid(self):
        """Test validating valid range expression"""
        result = ExcelValidator.validate_range_expression("A1:C10")
        assert result == "A1:C10"
    
    def test_validate_range_expression_invalid(self):
        """Test validating invalid range expression"""
        with pytest.raises(InvalidRangeError):
            ExcelValidator.validate_range_expression("invalid_range")


class TestRangeParser:
    """Test cases for RangeParser class"""
    
    def test_parse_single_cell(self):
        """Test parsing single cell"""
        result = RangeParser.parse("A1")
        assert result.sheet_name is None
        assert result.start_row == 1
        assert result.start_col == 1
        assert result.end_row == 1
        assert result.end_col == 1
        assert result.range_type == "cell"
    
    def test_parse_range_with_sheet(self):
        """Test parsing range with sheet name"""
        result = RangeParser.parse("Sheet1!A1:C10")
        assert result.sheet_name == "Sheet1"
        assert result.start_row == 1
        assert result.start_col == 1
        assert result.end_row == 10
        assert result.end_col == 3
        assert result.range_type == "range"
    
    def test_parse_entire_row(self):
        """Test parsing entire row"""
        result = RangeParser.parse("1:5")
        assert result.start_row == 1
        assert result.end_row == 5
        assert result.range_type == "row"
    
    def test_parse_single_row(self):
        """Test parsing single row"""
        result = RangeParser.parse("3")
        assert result.start_row == 3
        assert result.end_row == 3
        assert result.range_type == "row"
    
    def test_parse_entire_column(self):
        """Test parsing entire column"""
        result = RangeParser.parse("A:C")
        assert result.start_col == 1
        assert result.end_col == 3
        assert result.range_type == "column"
    
    def test_parse_single_column(self):
        """Test parsing single column"""
        result = RangeParser.parse("B")
        assert result.start_col == 2
        assert result.end_col == 2
        assert result.range_type == "column"
    
    def test_parse_invalid_range(self):
        """Test parsing invalid range"""
        with pytest.raises(InvalidRangeError):
            RangeParser.parse("invalid_range")
    
    def test_parse_empty_range(self):
        """Test parsing empty range"""
        with pytest.raises(InvalidRangeError):
            RangeParser.parse("")
    
    def test_column_letter_to_number(self):
        """Test column letter to number conversion"""
        assert RangeParser.column_letter_to_number("A") == 1
        assert RangeParser.column_letter_to_number("Z") == 26
        assert RangeParser.column_letter_to_number("AA") == 27
        assert RangeParser.column_letter_to_number("AZ") == 52
        assert RangeParser.column_letter_to_number("ZZ") == 702
    
    def test_number_to_column_letter(self):
        """Test number to column letter conversion"""
        assert RangeParser.number_to_column_letter(1) == "A"
        assert RangeParser.number_to_column_letter(26) == "Z"
        assert RangeParser.number_to_column_letter(27) == "AA"
        assert RangeParser.number_to_column_letter(52) == "AZ"
        assert RangeParser.number_to_column_letter(702) == "ZZ"
    
    def test_coordinate_to_cell_ref(self):
        """Test coordinate to cell reference conversion"""
        assert RangeParser.coordinate_to_cell_ref(1, 1) == "A1"
        assert RangeParser.coordinate_to_cell_ref(5, 3) == "C5"
        assert RangeParser.coordinate_to_cell_ref(10, 26) == "Z10"
        assert RangeParser.coordinate_to_cell_ref(1, 27) == "AA1"


class TestExceptions:
    """Test cases for custom exceptions"""
    
    def test_file_not_found_exception(self):
        """Test FileNotFoundError"""
        exc = FileNotFoundError("test_file.xlsx")
        assert str(exc) == "test_file.xlsx"
        assert isinstance(exc, Exception)
    
    def test_sheet_not_found_exception(self):
        """Test SheetNotFoundError"""
        exc = SheetNotFoundError("Sheet1")
        assert str(exc) == "Sheet1"
        assert isinstance(exc, Exception)
    
    def test_data_validation_exception(self):
        """Test DataValidationError"""
        exc = DataValidationError("Invalid data")
        assert str(exc) == "Invalid data"
        assert isinstance(exc, Exception)
    
    def test_invalid_range_error(self):
        """Test InvalidRangeError"""
        exc = InvalidRangeError("A1:ZZ999")
        assert str(exc) == "A1:ZZ999"
        assert isinstance(exc, Exception)