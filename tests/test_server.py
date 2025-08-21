"""
Tests for Server MCP interfaces
"""

import pytest
import tempfile
from pathlib import Path
from src.server import (
    excel_list_sheets,
    excel_get_range,
    excel_update_range,
    excel_create_file,
    excel_create_sheet,
    excel_delete_sheet,
    excel_rename_sheet,
    excel_insert_rows,
    excel_insert_columns,
    excel_delete_rows,
    excel_delete_columns,
    excel_set_formula,
    excel_evaluate_formula,
    excel_format_cells,
    excel_regex_search
)


class TestServerInterfaces:
    """Test cases for Server MCP interfaces"""
    
    def test_excel_list_sheets(self, sample_excel_file):
        """Test excel_list_sheets interface"""
        result = excel_list_sheets(sample_excel_file)
        
        assert result['success'] is True
        assert 'sheets' in result
        assert 'active_sheet' in result
        assert len(result['sheets']) == 2
        assert "Sheet1" in result['sheets']
        assert "Sheet2" in result['sheets']
    
    def test_excel_list_sheets_invalid_file(self):
        """Test excel_list_sheets with invalid file"""
        result = excel_list_sheets("nonexistent_file.xlsx")
        
        assert result['success'] is False
        assert 'error' in result
        assert 'file_path' in result
    
    def test_excel_get_range(self, sample_excel_file):
        """Test excel_get_range interface"""
        result = excel_get_range(sample_excel_file, "A1:C5")
        
        assert result['success'] is True
        assert 'data' in result
        assert len(result['data']) == 5  # 5 rows
        assert result['data'][0][0] == "姓名"
        assert result['data'][4][2] == "人事部"
    
    def test_excel_get_range_with_formatting(self, sample_excel_file):
        """Test excel_get_range with formatting"""
        result = excel_get_range(sample_excel_file, "A1:D1", include_formatting=True)
        
        assert result['success'] is True
        assert 'data' in result
        assert 'range_info' in result
    
    def test_excel_get_range_with_sheet_name(self, sample_excel_file):
        """Test excel_get_range with sheet name"""
        result = excel_get_range(sample_excel_file, "Sheet2!A1:C3")
        
        assert result['success'] is True
        assert result['data'][0][0] == "产品"
        assert result['data'][2][2] == 30
    
    def test_excel_get_range_invalid_sheet(self, sample_excel_file):
        """Test excel_get_range with invalid sheet"""
        result = excel_get_range(sample_excel_file, "NonExistentSheet!A1")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_update_range(self, sample_excel_file):
        """Test excel_update_range interface"""
        data = [["新姓名", "新年龄"], ["测试1", 99]]
        result = excel_update_range(sample_excel_file, "A1:B2", data)
        
        assert result['success'] is True
        assert 'updated_cells' in result
        assert result['updated_cells'] == 4
        assert 'message' in result
    
    def test_excel_update_range_preserve_formulas(self, sample_excel_file):
        """Test excel_update_range with preserve_formulas"""
        data = [["总计行"]]
        result = excel_update_range(sample_excel_file, "A6", data, preserve_formulas=True)
        
        assert result['success'] is True
        assert 'updated_cells' in result
    
    def test_excel_update_range_invalid_sheet(self, sample_excel_file):
        """Test excel_update_range with invalid sheet"""
        data = [["测试"]]
        result = excel_update_range(sample_excel_file, "NonExistentSheet!A1", data)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_create_file(self, temp_dir):
        """Test excel_create_file interface"""
        file_path = temp_dir / "test_create.xlsx"
        result = excel_create_file(str(file_path))
        
        assert result['success'] is True
        assert 'file_path' in result
        assert 'sheets' in result
        assert len(result['sheets']) == 1
        
        # Verify file was created
        assert file_path.exists()
    
    def test_excel_create_file_with_sheets(self, temp_dir):
        """Test excel_create_file with custom sheets"""
        file_path = temp_dir / "test_create_sheets.xlsx"
        sheet_names = ["数据", "图表", "汇总"]
        result = excel_create_file(str(file_path), sheet_names)
        
        assert result['success'] is True
        assert len(result['sheets']) == 3
        assert "数据" in result['sheets']
        assert "图表" in result['sheets']
        assert "汇总" in result['sheets']
    
    def test_excel_create_sheet(self, sample_excel_file):
        """Test excel_create_sheet interface"""
        result = excel_create_sheet(sample_excel_file, "新工作表")
        
        assert result['success'] is True
        assert 'sheet_name' in result
        assert result['sheet_name'] == "新工作表"
        assert 'total_sheets' in result
        assert result['total_sheets'] == 3
    
    def test_excel_create_sheet_at_position(self, sample_excel_file):
        """Test excel_create_sheet at specific position"""
        result = excel_create_sheet(sample_excel_file, "首页", 0)
        
        assert result['success'] is True
        assert result['sheet_name'] == "首页"
    
    def test_excel_create_sheet_duplicate_name(self, sample_excel_file):
        """Test excel_create_sheet with duplicate name"""
        result = excel_create_sheet(sample_excel_file, "Sheet1")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_delete_sheet(self, sample_excel_file):
        """Test excel_delete_sheet interface"""
        result = excel_delete_sheet(sample_excel_file, "Sheet2")
        
        assert result['success'] is True
        assert 'deleted_sheet' in result
        assert result['deleted_sheet'] == "Sheet2"
        assert 'remaining_sheets' in result
        assert len(result['remaining_sheets']) == 1
        assert "Sheet1" in result['remaining_sheets']
    
    def test_excel_delete_sheet_nonexistent(self, sample_excel_file):
        """Test excel_delete_sheet with non-existent sheet"""
        result = excel_delete_sheet(sample_excel_file, "NonExistentSheet")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_rename_sheet(self, sample_excel_file):
        """Test excel_rename_sheet interface"""
        result = excel_rename_sheet(sample_excel_file, "Sheet1", "数据表")
        
        assert result['success'] is True
        assert 'old_name' in result
        assert result['old_name'] == "Sheet1"
        assert 'new_name' in result
        assert result['new_name'] == "数据表"
    
    def test_excel_rename_sheet_nonexistent(self, sample_excel_file):
        """Test excel_rename_sheet with non-existent sheet"""
        result = excel_rename_sheet(sample_excel_file, "NonExistentSheet", "新名称")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_rename_sheet_duplicate_name(self, sample_excel_file):
        """Test excel_rename_sheet with duplicate name"""
        result = excel_rename_sheet(sample_excel_file, "Sheet1", "Sheet2")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_insert_rows(self, sample_excel_file):
        """Test excel_insert_rows interface"""
        result = excel_insert_rows(sample_excel_file, "Sheet1", 2, 2)
        
        assert result['success'] is True
        assert 'inserted_rows' in result
        assert result['inserted_rows'] == 2
        assert 'message' in result
    
    def test_excel_insert_rows_single(self, sample_excel_file):
        """Test excel_insert_rows single row"""
        result = excel_insert_rows(sample_excel_file, "Sheet1", 3)
        
        assert result['success'] is True
        assert result['inserted_rows'] == 1
    
    def test_excel_insert_rows_invalid_sheet(self, sample_excel_file):
        """Test excel_insert_rows with invalid sheet"""
        result = excel_insert_rows(sample_excel_file, "NonExistentSheet", 1)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_insert_columns(self, sample_excel_file):
        """Test excel_insert_columns interface"""
        result = excel_insert_columns(sample_excel_file, "Sheet1", 2, 1)
        
        assert result['success'] is True
        assert 'inserted_columns' in result
        assert result['inserted_columns'] == 1
        assert 'message' in result
    
    def test_excel_insert_columns_multiple(self, sample_excel_file):
        """Test excel_insert_columns multiple columns"""
        result = excel_insert_columns(sample_excel_file, "Sheet1", 1, 3)
        
        assert result['success'] is True
        assert result['inserted_columns'] == 3
    
    def test_excel_insert_columns_invalid_sheet(self, sample_excel_file):
        """Test excel_insert_columns with invalid sheet"""
        result = excel_insert_columns(sample_excel_file, "NonExistentSheet", 1)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_delete_rows(self, sample_excel_file):
        """Test excel_delete_rows interface"""
        result = excel_delete_rows(sample_excel_file, "Sheet1", 2, 1)
        
        assert result['success'] is True
        assert 'deleted_rows' in result
        assert result['deleted_rows'] == 1
        assert 'message' in result
    
    def test_excel_delete_rows_multiple(self, sample_excel_file):
        """Test excel_delete_rows multiple rows"""
        result = excel_delete_rows(sample_excel_file, "Sheet1", 3, 2)
        
        assert result['success'] is True
        assert result['deleted_rows'] == 2
    
    def test_excel_delete_rows_invalid_sheet(self, sample_excel_file):
        """Test excel_delete_rows with invalid sheet"""
        result = excel_delete_rows(sample_excel_file, "NonExistentSheet", 1)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_delete_columns(self, sample_excel_file):
        """Test excel_delete_columns interface"""
        result = excel_delete_columns(sample_excel_file, "Sheet1", 2, 1)
        
        assert result['success'] is True
        assert 'deleted_columns' in result
        assert result['deleted_columns'] == 1
        assert 'message' in result
    
    def test_excel_delete_columns_multiple(self, sample_excel_file):
        """Test excel_delete_columns multiple columns"""
        result = excel_delete_columns(sample_excel_file, "Sheet1", 1, 2)
        
        assert result['success'] is True
        assert result['deleted_columns'] == 2
    
    def test_excel_delete_columns_invalid_sheet(self, sample_excel_file):
        """Test excel_delete_columns with invalid sheet"""
        result = excel_delete_columns(sample_excel_file, "NonExistentSheet", 1)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_set_formula(self, sample_excel_file):
        """Test excel_set_formula interface"""
        result = excel_set_formula(sample_excel_file, "Sheet1", "F1", "SUM(A1:A5)")
        
        assert result['success'] is True
        assert 'formula' in result
        assert result['formula'] == "SUM(A1:A5)"
        assert 'calculated_value' in result
        assert 'message' in result
    
    def test_excel_set_formula_invalid_sheet(self, sample_excel_file):
        """Test excel_set_formula with invalid sheet"""
        result = excel_set_formula(sample_excel_file, "NonExistentSheet", "A1", "SUM(B1:B10)")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_evaluate_formula(self, formula_excel_file):
        """Test excel_evaluate_formula interface"""
        result = excel_evaluate_formula(formula_excel_file, "SUM(A2:A4)")
        
        assert result['success'] is True
        assert 'formula' in result
        assert result['formula'] == "SUM(A2:A4)"
        assert 'result' in result
        assert result['result'] == 90  # 10 + 30 + 50
        assert 'result_type' in result
        assert 'execution_time_ms' in result
        assert 'message' in result
    
    def test_excel_evaluate_formula_with_context(self, formula_excel_file):
        """Test excel_evaluate_formula with context sheet"""
        result = excel_evaluate_formula(formula_excel_file, "SUM(A2:A4)", "Formulas")
        
        assert result['success'] is True
        assert result['result'] == 90
        assert 'context_sheet' in result
        assert result['context_sheet'] == "Formulas"
    
    def test_excel_evaluate_formula_invalid_file(self):
        """Test excel_evaluate_formula with invalid file"""
        result = excel_evaluate_formula("nonexistent_file.xlsx", "SUM(A1:A10)")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_format_cells(self, sample_excel_file):
        """Test excel_format_cells interface"""
        formatting = {
            'font': {'name': 'Arial', 'size': 14, 'bold': True, 'color': 'FF0000'},
            'fill': {'color': 'FFFF00'},
            'alignment': {'horizontal': 'center', 'vertical': 'middle'}
        }
        result = excel_format_cells(sample_excel_file, "Sheet1", "A1:D1", formatting)
        
        assert result['success'] is True
        assert 'formatted_count' in result
        assert result['formatted_count'] == 4
        assert 'message' in result
    
    def test_excel_format_cells_invalid_sheet(self, sample_excel_file):
        """Test excel_format_cells with invalid sheet"""
        formatting = {'font': {'bold': True}}
        result = excel_format_cells(sample_excel_file, "NonExistentSheet", "A1", formatting)
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_excel_regex_search(self, sample_excel_file):
        """Test excel_regex_search interface"""
        result = excel_regex_search(sample_excel_file, r"张三")
        
        assert result['success'] is True
        assert 'matches' in result
        assert 'match_count' in result
        assert result['match_count'] > 0
        assert 'searched_sheets' in result
        assert 'message' in result
        
        # Check match structure
        match = result['matches'][0]
        assert 'coordinate' in match
        assert 'sheet_name' in match
        assert 'value' in match
        assert 'matched_text' in match
    
    def test_excel_regex_search_with_flags(self, sample_excel_file):
        """Test excel_regex_search with flags"""
        result = excel_regex_search(sample_excel_file, r"技术部", flags="i")
        
        assert result['success'] is True
        assert result['match_count'] > 0
    
    def test_excel_regex_search_formulas_only(self, formula_excel_file):
        """Test excel_regex_search formulas only"""
        result = excel_regex_search(
            formula_excel_file, 
            r"SUM", 
            search_values=False, 
            search_formulas=True
        )
        
        assert result['success'] is True
        assert result['match_count'] > 0
    
    def test_excel_regex_search_invalid_file(self):
        """Test excel_regex_search with invalid file"""
        result = excel_regex_search("nonexistent_file.xlsx", r"test")
        
        assert result['success'] is False
        assert 'error' in result
    
    def test_all_interfaces_return_consistent_structure(self, sample_excel_file):
        """Test that all interfaces return consistent response structure"""
        # Test a few key interfaces
        interfaces = [
            lambda: excel_list_sheets(sample_excel_file),
            lambda: excel_get_range(sample_excel_file, "A1"),
            lambda: excel_create_sheet(sample_excel_file, "TestSheet"),
            lambda: excel_regex_search(sample_excel_file, r"test")
        ]
        
        for interface in interfaces:
            result = interface()
            
            # All should have success boolean
            assert 'success' in result
            assert isinstance(result['success'], bool)
            
            # If successful, should have appropriate data
            if result['success']:
                assert 'error' not in result
            else:
                assert 'error' in result
                assert isinstance(result['error'], str)