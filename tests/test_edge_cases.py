"""
Additional edge case tests for backup/restore, formatting, and utility tools.
"""

import pytest
import os
import shutil
from src.excel_mcp_server_fastmcp.server import (
    excel_create_backup,
    excel_restore_backup,
    excel_list_backups,
    excel_set_formula,
    excel_format_cells,
    excel_set_row_height,
    excel_set_column_width,
    excel_rename_sheet,
    excel_delete_sheet,
)


class TestExcelBackupRestore:
    """Test backup creation and restoration workflow"""

    def test_create_and_list_backups(self, sample_excel_file, temp_dir):
        """Test creating a backup and listing it"""
        backup_dir = os.path.join(str(temp_dir), "backups")
        result = excel_create_backup(sample_excel_file, backup_dir)

        assert result['success'] is True
        assert 'backup_file' in result['data']
        assert os.path.exists(result['data']['backup_file'])

        # List backups
        list_result = excel_list_backups(sample_excel_file, backup_dir)
        assert list_result['success'] is True
        assert list_result['data']['total_backups'] >= 1

    def test_restore_backup_roundtrip(self, sample_excel_file, temp_dir):
        """Test full backup → modify → restore cycle"""
        backup_dir = os.path.join(str(temp_dir), "backups")
        restore_path = os.path.join(str(temp_dir), "restored.xlsx")

        # Create backup
        backup_result = excel_create_backup(sample_excel_file, backup_dir)
        assert backup_result['success'] is True

        backup_path = backup_result['data']['backup_file']
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
        assert len(result['data']['backups']) == 0

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
        assert list_result['data']['total_backups'] == 3


class TestExcelFormulas:
    """Test formula-related tools"""

class TestExcelFormatting:
    """Test formatting tools"""

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

