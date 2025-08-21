"""
Fixed tests for utility modules - matching actual API implementation
"""

import pytest
from src.utils.exceptions import (
    FileNotFoundError, 
    SheetNotFoundError, 
    DataValidationError
)


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
    
    def test_exceptions_inheritance(self):
        """Test that exceptions inherit from Exception"""
        assert issubclass(FileNotFoundError, Exception)
        assert issubclass(SheetNotFoundError, Exception)
        assert issubclass(DataValidationError, Exception)
    
    def test_exceptions_can_be_raised_and_caught(self):
        """Test that exceptions can be raised and caught"""
        
        def raise_file_not_found():
            raise FileNotFoundError("test.xlsx")
        
        def raise_sheet_not_found():
            raise SheetNotFoundError("Sheet1")
        
        def raise_data_validation():
            raise DataValidationError("Invalid data")
        
        # Test FileNotFoundError
        with pytest.raises(FileNotFoundError) as exc_info:
            raise_file_not_found()
        assert str(exc_info.value) == "test.xlsx"
        
        # Test SheetNotFoundError
        with pytest.raises(SheetNotFoundError) as exc_info:
            raise_sheet_not_found()
        assert str(exc_info.value) == "Sheet1"
        
        # Test DataValidationError
        with pytest.raises(DataValidationError) as exc_info:
            raise_data_validation()
        assert str(exc_info.value) == "Invalid data"
    
    def test_exceptions_with_different_message_types(self):
        """Test exceptions with different message types"""
        # Test with string message
        exc1 = FileNotFoundError("string_message")
        assert str(exc1) == "string_message"
        
        # Test with empty string
        exc2 = SheetNotFoundError("")
        assert str(exc2) == ""
        
        # Test with unicode message
        exc3 = DataValidationError("中文消息")
        assert str(exc3) == "中文消息"
    
    def test_exceptions_are_pickleable(self):
        """Test that exceptions can be pickled (for multiprocessing)"""
        import pickle
        
        exc = FileNotFoundError("test.xlsx")
        pickled = pickle.dumps(exc)
        unpickled = pickle.loads(pickled)
        
        assert isinstance(unpickled, FileNotFoundError)
        assert str(unpickled) == "test.xlsx"
    
    def test_exceptions_have_custom_attributes(self):
        """Test that exceptions can have custom attributes"""
        # Create an exception with custom attributes
        exc = FileNotFoundError("test.xlsx")
        exc.custom_attr = "custom_value"
        
        assert exc.custom_attr == "custom_value"
        assert str(exc) == "test.xlsx"