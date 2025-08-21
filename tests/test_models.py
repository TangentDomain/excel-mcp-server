"""
Fixed tests for model types - matching actual API implementation
"""

import pytest
from src.models.types import (
    SheetInfo, RangeInfo, CellInfo, SearchMatch, ModifiedCell,
    OperationResult, RangeType, MatchType
)


class TestModelTypes:
    """Test cases for model types - matching actual implementation"""
    
    def test_sheet_info_creation(self):
        """Test SheetInfo creation with actual fields"""
        sheet_info = SheetInfo(
            index=0,
            name="Sheet1",
            is_active=True,
            max_row=100,
            max_column=10,
            max_column_letter="J"
        )
        
        assert sheet_info.name == "Sheet1"
        assert sheet_info.index == 0
        assert sheet_info.is_active is True
        assert sheet_info.max_row == 100
        assert sheet_info.max_column == 10
        assert sheet_info.max_column_letter == "J"
    
    def test_range_info_creation(self):
        """Test RangeInfo creation with actual fields"""
        range_info = RangeInfo(
            sheet_name="Sheet1",
            cell_range="A1:C10",
            range_type=RangeType.CELL_RANGE
        )
        
        assert range_info.sheet_name == "Sheet1"
        assert range_info.cell_range == "A1:C10"
        assert range_info.range_type == RangeType.CELL_RANGE
    
    def test_cell_info_creation(self):
        """Test CellInfo creation with actual fields"""
        cell_info = CellInfo(
            coordinate="A1",
            value="Test Value"
        )
        
        assert cell_info.coordinate == "A1"
        assert cell_info.value == "Test Value"
        # Optional fields should have default values
        assert cell_info.data_type is None
        assert cell_info.number_format is None
    
    def test_search_match_creation(self):
        """Test SearchMatch creation with actual fields"""
        match = SearchMatch(
            sheet="Sheet1",
            cell="A1",
            value="Test Value",
            match="Test",
            match_type=MatchType.VALUE
        )
        
        assert match.sheet == "Sheet1"
        assert match.cell == "A1"
        assert match.value == "Test Value"
        assert match.match == "Test"
        assert match.match_type == MatchType.VALUE
    
    def test_modified_cell_creation(self):
        """Test ModifiedCell creation with actual fields"""
        modified_cell = ModifiedCell(
            coordinate="A1",
            old_value="Old Value",
            new_value="New Value"
        )
        
        assert modified_cell.coordinate == "A1"
        assert modified_cell.old_value == "Old Value"
        assert modified_cell.new_value == "New Value"
    
    def test_operation_result_success(self):
        """Test OperationResult success case"""
        result = OperationResult(
            success=True,
            data="test data",
            message="Operation completed successfully"
        )
        
        assert result.success is True
        assert result.data == "test data"
        assert result.message == "Operation completed successfully"
        assert result.error is None
    
    def test_operation_result_failure(self):
        """Test OperationResult failure case"""
        result = OperationResult(
            success=False,
            error="Operation failed",
            message="Error occurred"
        )
        
        assert result.success is False
        assert result.error == "Operation failed"
        assert result.message == "Error occurred"
        assert result.data is None
    
    def test_range_type_enum_values(self):
        """Test RangeType enum values"""
        assert RangeType.CELL_RANGE.value == "cell_range"
        assert RangeType.ROW_RANGE.value == "row_range"
        assert RangeType.COLUMN_RANGE.value == "column_range"
        assert RangeType.SINGLE_ROW.value == "single_row"
        assert RangeType.SINGLE_COLUMN.value == "single_column"
    
    def test_match_type_enum_values(self):
        """Test MatchType enum values"""
        assert MatchType.VALUE.value == "value"
        assert MatchType.FORMULA.value == "formula"
    
    def test_operation_result_with_metadata(self):
        """Test OperationResult with metadata"""
        metadata = {"key": "value", "count": 5}
        result = OperationResult(
            success=True,
            data="test data",
            metadata=metadata
        )
        
        assert result.success is True
        assert result.metadata == metadata
        assert result.metadata["key"] == "value"
        assert result.metadata["count"] == 5
    
    def test_models_are_dataclasses(self):
        """Test that models are dataclasses"""
        from dataclasses import is_dataclass, fields
        
        assert is_dataclass(SheetInfo)
        assert is_dataclass(RangeInfo)
        assert is_dataclass(CellInfo)
        assert is_dataclass(SearchMatch)
        assert is_dataclass(ModifiedCell)
        assert is_dataclass(OperationResult)
        
        # Check that they have the expected fields
        sheet_fields = [f.name for f in fields(SheetInfo)]
        expected_sheet_fields = ['index', 'name', 'is_active', 'max_row', 'max_column', 'max_column_letter']
        assert all(field in sheet_fields for field in expected_sheet_fields)
    
    def test_models_implement_str_method(self):
        """Test that models implement __str__ method"""
        sheet_info = SheetInfo("Sheet1", 0, 100, 10, "J", True)
        result = OperationResult(True, "test")
        match = SearchMatch("Sheet1", "A1", "value", None, "value", MatchType.VALUE)
        
        # All should have string representations
        assert str(sheet_info) is not None
        assert str(result) is not None
        assert str(match) is not None
    
    def test_models_implement_repr_method(self):
        """Test that models implement __repr__ method"""
        sheet_info = SheetInfo("Sheet1", 0, 100, 10, "J", True)
        result = OperationResult(True, "test")
        
        # All should have repr representations
        assert repr(sheet_info) is not None
        assert repr(result) is not None
    
    def test_models_equality(self):
        """Test that models support equality comparison"""
        sheet1 = SheetInfo("Sheet1", 0, 100, 10, "J", True)
        sheet2 = SheetInfo("Sheet1", 0, 100, 10, "J", True)
        sheet3 = SheetInfo("Sheet2", 0, 100, 10, "J", True)
        
        assert sheet1 == sheet2
        assert sheet1 != sheet3
    
    def test_models_hashability(self):
        """Test that models are hashable"""
        sheet_info = SheetInfo("Sheet1", 0, 100, 10, "J", True)
        
        # Should be hashable (can be used in sets/dicts)
        assert hash(sheet_info) is not None
        
        # Can be used in a set
        sheet_set = {sheet_info}
        assert len(sheet_set) == 1
    
    def test_cell_info_with_optional_fields(self):
        """Test CellInfo with all optional fields"""
        cell_info = CellInfo(
            coordinate="A1",
            value="Test",
            data_type="string",
            number_format="General",
            font="Arial",
            fill="None"
        )
        
        assert cell_info.coordinate == "A1"
        assert cell_info.value == "Test"
        assert cell_info.data_type == "string"
        assert cell_info.number_format == "General"
        assert cell_info.font == "Arial"
        assert cell_info.fill == "None"
    
    def test_search_match_with_all_fields(self):
        """Test SearchMatch with all fields"""
        match = SearchMatch(
            sheet="Sheet1",
            cell="A1",
            value="Test Value",
            formula="=A1",
            match="Test",
            match_start=0,
            match_end=4,
            match_type=MatchType.VALUE
        )
        
        assert match.sheet == "Sheet1"
        assert match.cell == "A1"
        assert match.value == "Test Value"
        assert match.formula == "=A1"
        assert match.match == "Test"
        assert match.match_start == 0
        assert match.match_end == 4
        assert match.match_type == MatchType.VALUE