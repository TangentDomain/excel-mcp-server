"""
Tests for model types
"""

import pytest
from src.models.types import (
    SheetInfo, RangeInfo, CellInfo, ExcelData, ExcelDimensions,
    OperationResult, SearchMatch, MatchType, RangeType, ModifiedCell
)


class TestModelTypes:
    """Test cases for model types"""
    
    def test_sheet_info_creation(self):
        """Test SheetInfo creation"""
        sheet_info = SheetInfo(
            name="Sheet1",
            index=0,
            row_count=100,
            column_count=10,
            is_visible=True,
            is_active=True
        )
        
        assert sheet_info.name == "Sheet1"
        assert sheet_info.index == 0
        assert sheet_info.row_count == 100
        assert sheet_info.column_count == 10
        assert sheet_info.is_visible is True
        assert sheet_info.is_active is True
    
    def test_range_info_creation(self):
        """Test RangeInfo creation"""
        range_info = RangeInfo(
            sheet_name="Sheet1",
            start_row=1,
            start_col=1,
            end_row=5,
            end_col=3,
            range_type=RangeType.CELL_RANGE,
            cell_count=15
        )
        
        assert range_info.sheet_name == "Sheet1"
        assert range_info.start_row == 1
        assert range_info.start_col == 1
        assert range_info.end_row == 5
        assert range_info.end_col == 3
        assert range_info.range_type == RangeType.CELL_RANGE
        assert range_info.cell_count == 15
    
    def test_cell_info_creation(self):
        """Test CellInfo creation"""
        cell_info = CellInfo(
            coordinate="A1",
            row=1,
            column=1,
            value="Test Value",
            formula="=A1",
            data_type="string",
            is_merged=False,
            merge_range=None
        )
        
        assert cell_info.coordinate == "A1"
        assert cell_info.row == 1
        assert cell_info.column == 1
        assert cell_info.value == "Test Value"
        assert cell_info.formula == "=A1"
        assert cell_info.data_type == "string"
        assert cell_info.is_merged is False
        assert cell_info.merge_range is None
    
    def test_excel_data_creation(self):
        """Test ExcelData creation"""
        data = [[1, 2, 3], [4, 5, 6]]
        excel_data = ExcelData(
            data=data,
            range_info=RangeInfo(
                sheet_name="Sheet1",
                start_row=1,
                start_col=1,
                end_row=2,
                end_col=3,
                range_type=RangeType.CELL_RANGE,
                cell_count=6
            )
        )
        
        assert excel_data.data == data
        assert excel_data.range_info.sheet_name == "Sheet1"
        assert excel_data.range_info.cell_count == 6
    
    def test_excel_dimensions_creation(self):
        """Test ExcelDimensions creation"""
        dimensions = ExcelDimensions(
            row_count=100,
            column_count=10,
            used_range="A1:J100"
        )
        
        assert dimensions.row_count == 100
        assert dimensions.column_count == 10
        assert dimensions.used_range == "A1:J100"
    
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
    
    def test_search_match_creation(self):
        """Test SearchMatch creation"""
        match = SearchMatch(
            coordinate="A1",
            sheet_name="Sheet1",
            value="Test Value",
            formula="=A1",
            matched_text="Test",
            match_type=MatchType.VALUE,
            row=1,
            column=1
        )
        
        assert match.coordinate == "A1"
        assert match.sheet_name == "Sheet1"
        assert match.value == "Test Value"
        assert match.formula == "=A1"
        assert match.matched_text == "Test"
        assert match.match_type == MatchType.VALUE
        assert match.row == 1
        assert match.column == 1
    
    def test_modified_cell_creation(self):
        """Test ModifiedCell creation"""
        modified_cell = ModifiedCell(
            coordinate="A1",
            old_value="Old Value",
            new_value="New Value",
            row=1,
            column=1
        )
        
        assert modified_cell.coordinate == "A1"
        assert modified_cell.old_value == "Old Value"
        assert modified_cell.new_value == "New Value"
        assert modified_cell.row == 1
        assert modified_cell.column == 1
    
    def test_range_type_enum(self):
        """Test RangeType enum values"""
        assert RangeType.CELL == "cell"
        assert RangeType.CELL_RANGE == "range"
        assert RangeType.ROW == "row"
        assert RangeType.COLUMN == "column"
        assert RangeType.WORKSHEET == "worksheet"
    
    def test_match_type_enum(self):
        """Test MatchType enum values"""
        assert MatchType.VALUE == "value"
        assert MatchType.FORMULA == "formula"
        assert MatchType.FORMAT == "format"
        assert MatchType.COMMENT == "comment"
    
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
    
    def test_models_implement_str_method(self):
        """Test that models implement __str__ method"""
        sheet_info = SheetInfo("Sheet1", 0, 100, 10)
        result = OperationResult(True, "test")
        match = SearchMatch("A1", "Sheet1", "value", None, "value", MatchType.VALUE, 1, 1)
        
        # All should have string representations
        assert str(sheet_info) is not None
        assert str(result) is not None
        assert str(match) is not None
    
    def test_models_implement_repr_method(self):
        """Test that models implement __repr__ method"""
        sheet_info = SheetInfo("Sheet1", 0, 100, 10)
        result = OperationResult(True, "test")
        
        # All should have repr representations
        assert repr(sheet_info) is not None
        assert repr(result) is not None
    
    def test_models_are_dataclasses(self):
        """Test that models are dataclasses"""
        from dataclasses import is_dataclass
        
        assert is_dataclass(SheetInfo)
        assert is_dataclass(RangeInfo)
        assert is_dataclass(CellInfo)
        assert is_dataclass(ExcelData)
        assert is_dataclass(ExcelDimensions)
        assert is_dataclass(OperationResult)
        assert is_dataclass(SearchMatch)
        assert is_dataclass(ModifiedCell)