"""
Tests for previously untested tools:
- excel_compare_files
- excel_compare_sheets
- excel_convert_format
- excel_export_to_csv
- excel_import_from_csv
- excel_merge_files
- excel_search_directory
"""

import pytest
import os
from src.excel_mcp_server_fastmcp.server import (
    excel_compare_files,
    excel_compare_sheets,
    excel_convert_format,
    excel_export_to_csv,
    excel_import_from_csv,
    excel_merge_files,
    excel_search_directory,
    excel_create_file,
)


class TestExcelExportImportCSV:
    """Test CSV export/import functionality"""

    def test_export_to_csv_default(self, sample_excel_file, temp_dir):
        """Test exporting to CSV with defaults"""
        output_path = os.path.join(str(temp_dir), "export.csv")
        result = excel_export_to_csv(sample_excel_file, output_path)

        assert result['success'] is True
        assert result['data']['output_path'] == output_path
        assert result['data']['row_count'] > 0
        assert os.path.exists(output_path)

        # Verify CSV content
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert 'name' in content or '姓名' in content

    def test_export_to_csv_specific_sheet(self, sample_excel_file, temp_dir):
        """Test exporting a specific sheet to CSV"""
        output_path = os.path.join(str(temp_dir), "sheet2.csv")
        result = excel_export_to_csv(sample_excel_file, output_path, sheet_name="Sheet2")

        assert result['success'] is True
        assert os.path.exists(output_path)

    def test_export_to_csv_nonexistent_sheet(self, sample_excel_file, temp_dir):
        """Test exporting a non-existent sheet returns error"""
        output_path = os.path.join(str(temp_dir), "fail.csv")
        result = excel_export_to_csv(sample_excel_file, output_path, sheet_name="NoSuchSheet")

        assert result['success'] is False

    def test_export_to_csv_invalid_file(self, temp_dir):
        """Test exporting a non-existent file returns error"""
        output_path = os.path.join(str(temp_dir), "fail.csv")
        result = excel_export_to_csv("/nonexistent/file.xlsx", output_path)

        assert result['success'] is False

    def test_import_from_csv(self, sample_excel_file, temp_dir):
        """Test creating Excel from CSV"""
        # First export to CSV
        csv_path = os.path.join(str(temp_dir), "data.csv")
        excel_export_to_csv(sample_excel_file, csv_path)

        # Then import back
        output_path = os.path.join(str(temp_dir), "imported.xlsx")
        result = excel_import_from_csv(csv_path, output_path, sheet_name="Imported")

        assert result['success'] is True
        assert result['data']['sheet_name'] == "Imported"
        assert result['data']['row_count'] > 0
        assert os.path.exists(output_path)

    def test_import_from_csv_no_header(self, temp_dir):
        """Test importing CSV without header row"""
        csv_path = os.path.join(str(temp_dir), "no_header.csv")
        with open(csv_path, 'w', encoding='utf-8') as f:
            f.write("100,200,300\n")
            f.write("400,500,600\n")

        output_path = os.path.join(str(temp_dir), "no_header.xlsx")
        result = excel_import_from_csv(csv_path, output_path, has_header=False)

        assert result['success'] is True
        assert result['data']['row_count'] == 2

    def test_import_from_csv_invalid_path(self, temp_dir):
        """Test importing from non-existent CSV"""
        output_path = os.path.join(str(temp_dir), "fail.xlsx")
        result = excel_import_from_csv("/nonexistent.csv", output_path)

        assert result['success'] is False


class TestExcelConvertFormat:
    """Test format conversion functionality"""

    def test_convert_xlsx_to_json(self, sample_excel_file, temp_dir):
        """Test converting Excel to JSON"""
        output_path = os.path.join(str(temp_dir), "data.json")
        result = excel_convert_format(sample_excel_file, output_path, "json")

        assert result['success'] is True
        assert os.path.exists(output_path)

        import json
        with open(output_path, 'r') as f:
            data = json.load(f)
        assert isinstance(data, dict)

    def test_convert_invalid_format(self, sample_excel_file, temp_dir):
        """Test converting to unsupported format"""
        output_path = os.path.join(str(temp_dir), "data.pdf")
        result = excel_convert_format(sample_excel_file, output_path, "pdf")

        assert result['success'] is False

    def test_convert_nonexistent_file(self, temp_dir):
        """Test converting non-existent file"""
        output_path = os.path.join(str(temp_dir), "out.json")
        result = excel_convert_format("/nonexistent.xlsx", output_path, "json")

        assert result['success'] is False


class TestExcelMergeFiles:
    """Test file merging functionality"""

    def test_merge_files_as_sheets(self, temp_dir):
        """Test merging multiple files as separate sheets"""
        files = []
        for i in range(3):
            path = os.path.join(str(temp_dir), f"file{i}.xlsx")
            excel_create_file(path, sheet_names=[f"Data{i}"])
            files.append(path)

        output_path = os.path.join(str(temp_dir), "merged.xlsx")
        result = excel_merge_files(files, output_path, merge_mode="sheets")

        assert result['success'] is True
        assert result['data']['merged_files'] == 3
        assert os.path.exists(output_path)

    def test_merge_files_empty_list(self, temp_dir):
        """Test merging with empty file list"""
        output_path = os.path.join(str(temp_dir), "empty.xlsx")
        result = excel_merge_files([], output_path)

        assert result['success'] is False

    def test_merge_files_nonexistent(self, temp_dir):
        """Test merging with non-existent files"""
        output_path = os.path.join(str(temp_dir), "fail.xlsx")
        result = excel_merge_files(["/nonexistent1.xlsx", "/nonexistent2.xlsx"], output_path)

        assert result['success'] is False


class TestExcelSearchDirectory:
    """Test directory-wide search functionality"""

    def test_search_directory_basic(self, temp_dir_with_excel_files):
        """Test basic directory search"""
        result = excel_search_directory(temp_dir_with_excel_files, "标题")

        assert result['success'] is True
        assert 'data' in result
        assert isinstance(result['data'], list)
        assert result['metadata']['total_matches'] > 0

    def test_search_directory_case_sensitive(self, temp_dir_with_excel_files):
        """Test case-sensitive search"""
        result = excel_search_directory(temp_dir_with_excel_files, "标题", case_sensitive=True)

        assert result['success'] is True

    def test_search_directory_no_match(self, temp_dir_with_excel_files):
        """Test search with no matches"""
        result = excel_search_directory(temp_dir_with_excel_files, "ZZZNONEXISTENT123")

        assert result['success'] is True
        assert result['metadata']['total_matches'] == 0

    def test_search_directory_regex(self, temp_dir_with_excel_files):
        """Test regex search"""
        result = excel_search_directory(temp_dir_with_excel_files, r"标题\d", use_regex=True)

        assert result['success'] is True
        assert isinstance(result['data'], list)

    def test_search_directory_nonexistent(self, temp_dir):
        """Test searching non-existent directory"""
        result = excel_search_directory("/nonexistent/path", "test")

        assert result['success'] is False

    def test_search_directory_whole_word(self, temp_dir_with_excel_files):
        """Test whole word matching"""
        result = excel_search_directory(temp_dir_with_excel_files, "标题", whole_word=True)

        assert result['success'] is True

    def test_search_directory_file_extensions(self, temp_dir):
        """Test search with file extension filter"""
        from openpyxl import Workbook
        path = os.path.join(str(temp_dir), "test.xlsx")
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "searchterm"
        wb.save(path)

        result = excel_search_directory(
            str(temp_dir), "searchterm",
            file_extensions=[".xlsx"]
        )
        assert result['success'] is True
        assert result['metadata']['total_matches'] > 0


class TestExcelCompareFiles:
    """Test file comparison functionality"""

    def test_compare_identical_files(self, sample_excel_file, temp_dir):
        """Test comparing identical files"""
        import shutil
        copy_path = os.path.join(str(temp_dir), "copy.xlsx")
        shutil.copy2(sample_excel_file, copy_path)

        result = excel_compare_files(sample_excel_file, copy_path)

        assert result['success'] is True
        assert result['data']['identical'] is True
        assert result['data']['total_differences'] == 0

    def test_compare_different_files(self, sample_excel_file, temp_dir):
        """Test comparing different files"""
        from openpyxl import Workbook
        diff_path = os.path.join(str(temp_dir), "different.xlsx")
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Different"
        wb.save(diff_path)

        result = excel_compare_files(sample_excel_file, diff_path)

        assert result['success'] is True
        assert result['data']['identical'] is False

    def test_compare_nonexistent_files(self, temp_dir):
        """Test comparing non-existent files"""
        result = excel_compare_files("/nonexistent1.xlsx", "/nonexistent2.xlsx")

        assert result['success'] is False


class TestExcelCompareSheets:
    """Test sheet comparison functionality"""

    def test_compare_identical_sheets(self, sample_excel_file, temp_dir):
        """Test comparing identical sheets"""
        import shutil
        copy_path = os.path.join(str(temp_dir), "copy.xlsx")
        shutil.copy2(sample_excel_file, copy_path)

        result = excel_compare_sheets(sample_excel_file, "Sheet1", copy_path, "Sheet1")

        assert result['success'] is True
        assert result['data']['total_differences'] == 0

    def test_compare_different_sheets(self, sample_excel_file, temp_dir):
        """Test comparing sheets with different data"""
        import shutil
        from openpyxl import load_workbook
        copy_path = os.path.join(str(temp_dir), "modified.xlsx")
        shutil.copy2(sample_excel_file, copy_path)

        # Modify the copy
        wb = load_workbook(copy_path)
        ws = wb["Sheet1"]
        ws['A3'] = "MODIFIED_VALUE"
        wb.save(copy_path)
        wb.close()

        result = excel_compare_sheets(sample_excel_file, "Sheet1", copy_path, "Sheet1")

        assert result['success'] is True
        assert result['data']['total_differences'] > 0

    def test_compare_nonexistent_sheet(self, sample_excel_file, temp_dir):
        """Test comparing non-existent sheet"""
        import shutil
        copy_path = os.path.join(str(temp_dir), "copy.xlsx")
        shutil.copy2(sample_excel_file, copy_path)

        result = excel_compare_sheets(sample_excel_file, "NoSuchSheet", copy_path, "Sheet1")

        assert result['success'] is False
