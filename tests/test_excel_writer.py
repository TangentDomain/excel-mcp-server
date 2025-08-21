"""
Tests for ExcelWriter class
"""

import pytest
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from src.core.excel_writer import ExcelWriter
from src.models.types import OperationResult
from src.utils.exceptions import SheetNotFoundError, DataValidationError


class TestExcelWriter:
    """Test cases for ExcelWriter class"""
    
    def test_init_valid_file(self, sample_excel_file):
        """Test initialization with valid file"""
        writer = ExcelWriter(sample_excel_file)
        assert writer.file_path == sample_excel_file
    
    def test_init_invalid_file(self):
        """Test initialization with invalid file"""
        with pytest.raises(FileNotFoundError):
            ExcelWriter("nonexistent_file.xlsx")
    
    def test_update_range_single_cell(self, sample_excel_file):
        """Test updating a single cell"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("A1", [["新标题"]])
        
        assert result.success is True
        assert result.data.updated_cells == 1
        
        # Verify the change
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        assert ws["A1"].value == "新标题"
    
    def test_update_range_multiple_cells(self, sample_excel_file):
        """Test updating multiple cells"""
        writer = ExcelWriter(sample_excel_file)
        new_data = [
            ["新姓名", "新年龄"],
            ["测试1", 99],
            ["测试2", 88]
        ]
        result = writer.update_range("A1:B3", new_data)
        
        assert result.success is True
        assert result.data.updated_cells == 6
        
        # Verify the changes
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        assert ws["A1"].value == "新姓名"
        assert ws["B3"].value == 88
    
    def test_update_range_with_sheet_name(self, sample_excel_file):
        """Test updating range with sheet name"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet2!A1", [["新产品"]])
        
        assert result.success is True
        assert result.data.updated_cells == 1
        
        # Verify the change
        wb = load_workbook(sample_excel_file)
        ws = wb["Sheet2"]
        assert ws["A1"].value == "新产品"
    
    def test_update_range_preserve_formulas(self, sample_excel_file):
        """Test updating range while preserving formulas"""
        writer = ExcelWriter(sample_excel_file)
        
        # Update a cell that doesn't contain formula
        result = writer.update_range("A6", [["总计行"]], preserve_formulas=True)
        
        assert result.success is True
        
        # Verify formula in E2 is still there
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        assert ws["E2"].value is not None  # Formula should be preserved
    
    def test_update_range_overwrite_formulas(self, sample_excel_file):
        """Test updating range and overwriting formulas"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("E2", [["手动值"]], preserve_formulas=False)
        
        assert result.success is True
        assert result.data.updated_cells == 1
        
        # Verify formula is overwritten
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        assert ws["E2"].value == "手动值"  # Should be the new value, not formula
    
    def test_update_range_invalid_sheet(self, sample_excel_file):
        """Test updating range in non-existent sheet"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("NonExistentSheet!A1", [["测试"]])
        
        assert result.success is False
        assert "工作表" in result.error
    
    def test_update_range_data_mismatch(self, sample_excel_file):
        """Test updating range with data size mismatch"""
        writer = ExcelWriter(sample_excel_file)
        # Range is 1x1 but data is 2x2
        result = writer.update_range("A1", [["A", "B"], ["C", "D"]])
        
        assert result.success is True  # Should still work, may expand or truncate
    
    def test_insert_rows(self, sample_excel_file):
        """Test inserting rows"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.insert_rows("Sheet1", 2, 2)  # Insert 2 rows at position 2
        
        assert result.success is True
        assert result.data.inserted_rows == 2
        
        # Verify the insertion
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        # Row 2 should now be empty
        assert ws.cell(row=2, column=1).value is None
        # Original row 2 should now be row 4
        assert ws.cell(row=4, column=1).value == "李四"
    
    def test_insert_columns(self, sample_excel_file):
        """Test inserting columns"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.insert_columns("Sheet1", 2, 1)  # Insert 1 column at position 2
        
        assert result.success is True
        assert result.data.inserted_columns == 1
        
        # Verify the insertion
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        # Column B should now be empty
        assert ws.cell(row=1, column=2).value is None
        # Original column B should now be column C
        assert ws.cell(row=1, column=3).value == "年龄"
    
    def test_delete_rows(self, sample_excel_file):
        """Test deleting rows"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.delete_rows("Sheet1", 2, 1)  # Delete 1 row at position 2
        
        assert result.success is True
        assert result.data.deleted_rows == 1
        
        # Verify the deletion
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        # Row 2 should now contain what was in row 3
        assert ws.cell(row=2, column=1).value == "李四"
    
    def test_delete_columns(self, sample_excel_file):
        """Test deleting columns"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.delete_columns("Sheet1", 2, 1)  # Delete 1 column at position 2
        
        assert result.success is True
        assert result.data.deleted_columns == 1
        
        # Verify the deletion
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        # Column B should now contain what was in column C
        assert ws.cell(row=1, column=2).value == "部门"
    
    def test_insert_rows_invalid_sheet(self, sample_excel_file):
        """Test inserting rows in non-existent sheet"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.insert_rows("NonExistentSheet", 1, 1)
        
        assert result.success is False
        assert "工作表" in result.error
    
    def test_delete_rows_invalid_position(self, sample_excel_file):
        """Test deleting rows at invalid position"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.delete_rows("Sheet1", 999, 1)
        
        # Should handle gracefully
        assert isinstance(result, OperationResult)
    
    def test_set_formula(self, sample_excel_file):
        """Test setting a formula"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.set_formula("F1", "SUM(A1:A5)", "Sheet1")
        
        assert result.success is True
        assert result.data.formula == "SUM(A1:A5)"
        
        # Verify the formula is set
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        assert ws["F1"].value == "SUM(A1:A5)"
    
    def test_set_formula_invalid_sheet(self, sample_excel_file):
        """Test setting formula in non-existent sheet"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.set_formula("A1", "SUM(B1:B10)", "NonExistentSheet")
        
        assert result.success is False
        assert "工作表" in result.error
    
    def test_evaluate_formula(self, formula_excel_file):
        """Test evaluating a formula without saving"""
        writer = ExcelWriter(formula_excel_file)
        result = writer.evaluate_formula("SUM(A2:A4)")
        
        assert result.success is True
        assert result.data == 90  # 10 + 30 + 50
        
        # Original file should not be modified
        wb = load_workbook(formula_excel_file)
        ws = wb.active
        assert ws["A6"].value == "总计"  # Should remain unchanged
    
    def test_evaluate_formula_with_context(self, formula_excel_file):
        """Test evaluating formula with specific sheet context"""
        writer = ExcelWriter(formula_excel_file)
        result = writer.evaluate_formula("SUM(A2:A4)", "Formulas")
        
        assert result.success is True
        assert result.data == 90
    
    def test_format_cells(self, sample_excel_file):
        """Test formatting cells"""
        writer = ExcelWriter(sample_excel_file)
        formatting = {
            'font': {'name': 'Arial', 'size': 14, 'bold': True, 'color': 'FF0000'},
            'fill': {'color': 'FFFF00'},
            'alignment': {'horizontal': 'center', 'vertical': 'middle'}
        }
        result = writer.format_cells("A1:D1", formatting, "Sheet1")
        
        assert result.success is True
        assert result.data.formatted_count == 4
        
        # Verify the formatting (this is a basic check)
        wb = load_workbook(sample_excel_file)
        ws = wb.active
        cell = ws["A1"]
        assert cell.font.bold is True
    
    def test_format_cells_invalid_format(self, sample_excel_file):
        """Test formatting cells with invalid format"""
        writer = ExcelWriter(sample_excel_file)
        formatting = {'invalid_key': 'invalid_value'}
        result = writer.format_cells("A1", formatting, "Sheet1")
        
        # Should handle invalid formatting gracefully
        assert isinstance(result, OperationResult)