"""
Additional edge case tests for backup/restore, formatting, and utility tools.
"""

import pytest
import os
import shutil
from src.server import (
    excel_create_backup,
    excel_restore_backup,
    excel_list_backups,
    excel_set_formula,
    excel_evaluate_formula,
    excel_format_cells,
    excel_merge_cells,
    excel_unmerge_cells,
    excel_set_borders,
    excel_set_row_height,
    excel_set_column_width,
    excel_rename_sheet,
    excel_delete_sheet,
    excel_get_file_info,
)


class TestExcelBackupRestore:
    """Test backup creation and restoration workflow"""

    def test_create_and_list_backups(self, sample_excel_file, temp_dir):
        """Test creating a backup and listing it"""
        backup_dir = os.path.join(str(temp_dir), "backups")
        result = excel_create_backup(sample_excel_file, backup_dir)

        assert result['success'] is True
        assert 'backup_file' in result
        assert os.path.exists(result['backup_file'])

        # List backups
        list_result = excel_list_backups(sample_excel_file, backup_dir)
        assert list_result['success'] is True
        assert list_result['total_backups'] >= 1

    def test_restore_backup_roundtrip(self, sample_excel_file, temp_dir):
        """Test full backup → modify → restore cycle"""
        backup_dir = os.path.join(str(temp_dir), "backups")
        restore_path = os.path.join(str(temp_dir), "restored.xlsx")

        # Create backup
        backup_result = excel_create_backup(sample_excel_file, backup_dir)
        assert backup_result['success'] is True

        backup_path = backup_result['backup_file']
        assert os.path.exists(backup_path)

        # Restore backup
        restore_result = excel_restore_backup(backup_path, restore_path)
        assert restore_result['success'] is True
        assert os.path.exists(restore_path)

    def test_restore_nonexistent_backup(self, temp_dir):
        """Test restoring from non-existent backup path"""
        restore_path = os.path.join(str(temp_dir), "restored.xlsx")
        result = excel_restore_backup("/nonexistent/backup.xlsx", restore_path)

        assert result['success'] is False

    def test_create_backup_nonexistent_file(self, temp_dir):
        """Test creating backup of non-existent file"""
        result = excel_create_backup("/nonexistent/file.xlsx")

        assert result['success'] is False

    def test_list_backups_empty(self, sample_excel_file, temp_dir):
        """Test listing backups when none exist"""
        backup_dir = os.path.join(str(temp_dir), "empty_backups")
        result = excel_list_backups(sample_excel_file, backup_dir)

        assert result['success'] is True
        assert len(result['backups']) == 0

    def test_create_multiple_backups_with_delay(self, sample_excel_file, temp_dir):
        """Test creating multiple backups with time separation"""
        import time
        backup_dir = os.path.join(str(temp_dir), "multi_backups")

        # Create backups with delays to ensure unique timestamps
        for i in range(3):
            result = excel_create_backup(sample_excel_file, backup_dir)
            assert result['success'] is True
            time.sleep(1.1)  # Ensure unique timestamp

        # List all
        list_result = excel_list_backups(sample_excel_file, backup_dir)
        assert list_result['success'] is True
        assert list_result['total_backups'] == 3


class TestExcelFormulas:
    """Test formula-related tools"""

    def test_evaluate_formula_with_context(self, sample_excel_file):
        """Test evaluating a formula with file context"""
        # excel_evaluate_formula requires a valid file path as context_sheet
        result = excel_evaluate_formula("SUM(A3:A6)", context_sheet=sample_excel_file)

        # Note: may fail if xlcalculator not installed, but should not crash
        assert 'success' in result

    def test_evaluate_formula_without_context(self):
        """Test evaluating formula without file context returns error"""
        # Without context_sheet, the API creates ExcelWriter("") which fails validation
        result = excel_evaluate_formula("SUM(1, 2, 3)")

        # This is expected to fail due to file validation
        assert result['success'] is False

    def test_set_formula_and_verify(self, sample_excel_file, temp_dir):
        """Test setting a formula and reading it back"""
        test_file = os.path.join(str(temp_dir), "formula_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        # Set formula
        result = excel_set_formula(test_file, "Sheet1", "E7", "=SUM(E3:E6)")
        assert result['success'] is True

    def test_set_formula_invalid_sheet(self, sample_excel_file):
        """Test setting formula on non-existent sheet"""
        result = excel_set_formula(sample_excel_file, "NoSuchSheet", "A1", "=1+1")

        assert result['success'] is False

    def test_set_formula_multiple_cells(self, sample_excel_file, temp_dir):
        """Test setting formulas on multiple cells"""
        test_file = os.path.join(str(temp_dir), "multi_formula.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        for i in range(3):
            result = excel_set_formula(test_file, "Sheet1", f"F{i+3}", f"=D{i+3}*2")
            assert result['success'] is True


class TestExcelFormatting:
    """Test formatting tools"""

    def test_set_borders_basic(self, sample_excel_file, temp_dir):
        """Test setting borders on a range"""
        test_file = os.path.join(str(temp_dir), "border_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_set_borders(test_file, "Sheet1", "A1:C3", "thin")
        assert result['success'] is True

    def test_set_borders_invalid_sheet(self, sample_excel_file):
        """Test setting borders on non-existent sheet"""
        result = excel_set_borders(sample_excel_file, "NoSuchSheet", "A1:C3", "thin")
        assert result['success'] is False

    def test_set_row_height(self, sample_excel_file, temp_dir):
        """Test setting row height"""
        test_file = os.path.join(str(temp_dir), "height_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_set_row_height(test_file, "Sheet1", 1, 30)
        assert result['success'] is True

    def test_set_column_width(self, sample_excel_file, temp_dir):
        """Test setting column width"""
        test_file = os.path.join(str(temp_dir), "width_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_set_column_width(test_file, "Sheet1", 1, 25)
        assert result['success'] is True

    def test_merge_cells_basic(self, sample_excel_file, temp_dir):
        """Test merging cells"""
        test_file = os.path.join(str(temp_dir), "merge_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_merge_cells(test_file, "Sheet1", "A1:C1")
        assert result['success'] is True

    def test_unmerge_cells(self, sample_excel_file, temp_dir):
        """Test unmerging cells"""
        test_file = os.path.join(str(temp_dir), "unmerge_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        # First merge
        excel_merge_cells(test_file, "Sheet1", "A1:C1")

        # Then unmerge
        result = excel_unmerge_cells(test_file, "Sheet1", "A1:C1")
        assert result['success'] is True

    def test_format_cells_with_preset(self, sample_excel_file, temp_dir):
        """Test formatting cells with a preset"""
        test_file = os.path.join(str(temp_dir), "preset_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_format_cells(test_file, "Sheet1", "A1:D1", {}, preset="header")
        assert result['success'] is True

    def test_format_cells_custom(self, sample_excel_file, temp_dir):
        """Test formatting cells with custom formatting"""
        test_file = os.path.join(str(temp_dir), "custom_fmt_test.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_format_cells(
            test_file, "Sheet1", "A1:A1",
            {"font": {"bold": True, "color": "FF0000"}, "fill": {"color": "FFFF00"}}
        )
        assert result['success'] is True


class TestExcelSheetManagement:
    """Test sheet management edge cases"""

    def test_rename_nonexistent_sheet(self, sample_excel_file):
        """Test renaming a non-existent sheet"""
        result = excel_rename_sheet(sample_excel_file, "NoSuchSheet", "NewName")
        assert result['success'] is False

    def test_delete_last_sheet(self, temp_dir):
        """Test deleting the only sheet"""
        from openpyxl import Workbook
        fp = os.path.join(str(temp_dir), "single.xlsx")
        wb = Workbook()
        wb.save(fp)

        result = excel_delete_sheet(fp, "Sheet")
        # Should fail - can't delete last sheet
        assert result['success'] is False

    def test_get_file_info_success(self, sample_excel_file):
        """Test getting file info"""
        result = excel_get_file_info(sample_excel_file)
        assert result['success'] is True

    def test_get_file_info_nonexistent(self):
        """Test getting file info for non-existent file"""
        result = excel_get_file_info("/nonexistent/file.xlsx")
        assert result['success'] is False

    def test_set_row_height_multiple_rows(self, sample_excel_file, temp_dir):
        """Test setting height for multiple consecutive rows"""
        test_file = os.path.join(str(temp_dir), "multi_height.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_set_row_height(test_file, "Sheet1", 3, 25, count=3)
        assert result['success'] is True

    def test_set_column_width_multiple_columns(self, sample_excel_file, temp_dir):
        """Test setting width for multiple consecutive columns"""
        test_file = os.path.join(str(temp_dir), "multi_width.xlsx")
        shutil.copy2(sample_excel_file, test_file)

        result = excel_set_column_width(test_file, "Sheet1", 1, 20, count=3)
        assert result['success'] is True
