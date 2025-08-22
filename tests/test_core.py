# -*- coding: utf-8 -*-
"""
核心Excel操作功能测试
合并了ExcelReader, ExcelWriter, ExcelManager的测试
这个文件替代了原本分散在多个文件中的核心功能测试
"""

import pytest
import tempfile
from pathlib import Path

from src.core.excel_reader import ExcelReader
from src.core.excel_writer import ExcelWriter
from src.core.excel_manager import ExcelManager
from src.models.types import OperationResult, SheetInfo, CellInfo, ModifiedCell
from src.utils.exceptions import ExcelFileNotFoundError, WorksheetNotFoundError


class TestExcelCore:
    """核心Excel操作功能的综合测试"""

    # ==================== Excel Reader 测试 ====================

    def test_reader_init_valid_file(self, sample_excel_file):
        """Test ExcelReader initialization with valid file"""
        reader = ExcelReader(sample_excel_file)
        assert reader.file_path == sample_excel_file

    def test_reader_init_invalid_file(self):
        """Test ExcelReader initialization with invalid file"""
        with pytest.raises(FileNotFoundError):
            ExcelReader("nonexistent_file.xlsx")

    def test_reader_list_sheets(self, sample_excel_file):
        """Test listing sheets"""
        reader = ExcelReader(sample_excel_file)
        result = reader.list_sheets()

        assert isinstance(result, OperationResult)
        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) >= 1

        # Check first sheet
        sheet1 = result.data[0]
        assert isinstance(sheet1, SheetInfo)
        assert hasattr(sheet1, 'name')
        assert hasattr(sheet1, 'index')

    def test_reader_get_range_cell_range(self, sample_excel_file):
        """Test getting a range of cells"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1:C5")

        assert isinstance(result, OperationResult)
        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 5
        assert len(result.data[0]) == 3

    def test_reader_get_range_single_cell(self, sample_excel_file):
        """Test getting a single cell"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1
        assert len(result.data[0]) == 1

    def test_reader_get_range_with_sheet_name(self, sample_excel_file):
        """Test getting range with explicit sheet name"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1:B2")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 2
        assert len(result.data[0]) == 2

    def test_reader_get_range_invalid_sheet(self, sample_excel_file):
        """Test getting range from non-existent sheet"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("NonExistentSheet!A1:A5")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_reader_get_range_with_formatting(self, sample_excel_file):
        """Test getting range with formatting information"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1:B2", include_formatting=True)

        assert result.success is True
        assert isinstance(result.data, list)

    def test_reader_get_range_unicode_content(self, sample_excel_file):
        """Test getting range with unicode content"""
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1:A5")

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 5

        # Check that unicode content is handled properly
        for row in result.data:
            assert len(row) == 1

    # ==================== Excel Writer 测试 ====================

    def test_writer_init_valid_file(self, sample_excel_file):
        """Test ExcelWriter initialization with valid file"""
        writer = ExcelWriter(sample_excel_file)
        assert writer.file_path == sample_excel_file

    def test_writer_init_invalid_file(self):
        """Test ExcelWriter initialization with invalid file"""
        with pytest.raises(ExcelFileNotFoundError):
            ExcelWriter("nonexistent_file.xlsx")

    def test_writer_update_range_single_cell(self, sample_excel_file):
        """Test updating a single cell"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet1!A1", [["新标题"]])

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 1
        assert isinstance(result.data[0], ModifiedCell)

    def test_writer_update_range_multiple_cells(self, sample_excel_file):
        """Test updating multiple cells"""
        writer = ExcelWriter(sample_excel_file)
        new_data = [
            ["新产品", "新价格"],
            ["产品A", 100],
            ["产品B", 200]
        ]
        result = writer.update_range("Sheet1!A1:B3", new_data)

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) == 6  # 3 rows * 2 columns

    def test_writer_update_range_with_sheet_name(self, sample_excel_file):
        """Test updating with explicit sheet name"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet2!A1", [["新产品"]])

        assert result.success is True
        assert len(result.data) == 1

    def test_writer_update_range_preserve_formulas(self, sample_excel_file):
        """Test updating while preserving formulas"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet1!A6", [["总计行"]], preserve_formulas=True)

        assert result.success is True
        assert len(result.data) >= 1

    def test_writer_update_range_overwrite_formulas(self, sample_excel_file):
        """Test updating and overwriting formulas"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("Sheet1!E2", [["手动值"]], preserve_formulas=False)

        assert result.success is True
        assert len(result.data) >= 1

    def test_writer_update_range_invalid_sheet(self, sample_excel_file):
        """Test updating non-existent sheet"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("NonExistentSheet!A1", [["测试"]])

        assert result.success is False
        assert "工作表" in result.error

    def test_writer_insert_rows(self, sample_excel_file):
        """Test inserting rows"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.insert_rows("Sheet1", 2, 2)

        assert result.success is True
        assert 'inserted_count' in result.metadata

    def test_writer_insert_columns(self, sample_excel_file):
        """Test inserting columns"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.insert_columns("Sheet1", 2, 1)

        assert result.success is True
        assert 'inserted_count' in result.metadata

    def test_writer_delete_rows(self, sample_excel_file):
        """Test deleting rows"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.delete_rows("Sheet1", 3, 1)

        assert result.success is True
        assert 'deleted_count' in result.metadata

    def test_writer_delete_columns(self, sample_excel_file):
        """Test deleting columns"""
        writer = ExcelWriter(sample_excel_file)
        result = writer.delete_columns("Sheet1", 3, 1)

        assert result.success is True
        assert 'deleted_count' in result.metadata

    def test_writer_format_cells(self, sample_excel_file):
        """Test formatting cells"""
        writer = ExcelWriter(sample_excel_file)
        formatting = {
            'font': {'name': 'Arial', 'size': 14, 'bold': True}
        }
        result = writer.format_cells("Sheet1", "A1:D1", formatting)

        assert result.success is True
        assert 'formatted_count' in result.metadata

    def test_writer_update_range_mixed_data_types(self, sample_excel_file):
        """Test updating with mixed data types"""
        writer = ExcelWriter(sample_excel_file)
        mixed_data = [
            ["文本", 123, 45.67, True],
            ["更多文本", 456, 78.90, False]
        ]
        result = writer.update_range("Sheet1!A1:D2", mixed_data)

        assert result.success is True
        assert len(result.data) == 8  # 2 rows * 4 columns

    # ==================== Excel Manager 测试 ====================

    def test_manager_init_valid_file(self, sample_excel_file):
        """Test ExcelManager initialization"""
        manager = ExcelManager(sample_excel_file)
        assert manager.file_path == sample_excel_file

    def test_manager_create_file(self, temp_dir):
        """Test creating a new Excel file"""
        file_path = temp_dir / "new_test_file.xlsx"
        result = ExcelManager.create_file(str(file_path), ["Sheet1", "Sheet2"])

        assert result.success is True
        assert file_path.exists()
        assert "created successfully" in result.message

    def test_manager_create_file_default_sheets(self, temp_dir):
        """Test creating file with default sheets"""
        file_path = temp_dir / "default_sheets.xlsx"
        result = ExcelManager.create_file(str(file_path))

        assert result.success is True
        assert file_path.exists()

    def test_manager_create_sheet(self, sample_excel_file):
        """Test creating a new sheet"""
        manager = ExcelManager(sample_excel_file)
        result = manager.create_sheet("新工作表")

        assert result.success is True
        assert result.data.name == "新工作表"

    def test_manager_create_sheet_duplicate_name(self, sample_excel_file):
        """Test creating sheet with duplicate name"""
        manager = ExcelManager(sample_excel_file)
        result = manager.create_sheet("Sheet1")  # Already exists

        assert result.success is False
        assert "已存在" in result.error or "exist" in result.error.lower()

    def test_manager_delete_sheet(self, sample_excel_file):
        """Test deleting a sheet"""
        manager = ExcelManager(sample_excel_file)
        # First create a sheet to delete
        manager.create_sheet("临时工作表")

        result = manager.delete_sheet("临时工作表")
        assert result.success is True

    def test_manager_delete_sheet_nonexistent(self, sample_excel_file):
        """Test deleting non-existent sheet"""
        manager = ExcelManager(sample_excel_file)
        result = manager.delete_sheet("不存在的工作表")

        assert result.success is False
        assert "不存在" in result.error or "not found" in result.error.lower()

    def test_manager_rename_sheet(self, sample_excel_file):
        """Test renaming a sheet"""
        manager = ExcelManager(sample_excel_file)
        result = manager.rename_sheet("Sheet1", "重命名工作表")

        assert result.success is True
        assert result.data.name == "重命名工作表"

    def test_manager_rename_sheet_nonexistent(self, sample_excel_file):
        """Test renaming non-existent sheet"""
        manager = ExcelManager(sample_excel_file)
        result = manager.rename_sheet("不存在的工作表", "新名称")

        assert result.success is False
        assert "不存在" in result.error or "not found" in result.error.lower()

    def test_manager_list_sheets(self, sample_excel_file):
        """Test listing all sheets through manager"""
        manager = ExcelManager(sample_excel_file)
        result = manager.list_sheets()

        assert result.success is True
        assert isinstance(result.data, list)
        assert len(result.data) >= 1

    # ==================== 综合测试 ====================

    def test_core_workflow_integration(self, temp_dir):
        """Test integrated workflow: create -> write -> read -> manage"""
        # 1. Create file
        file_path = temp_dir / "integration_test.xlsx"
        result = ExcelManager.create_file(str(file_path), ["数据表", "汇总表"])
        assert result.success is True

        # 2. Write data
        writer = ExcelWriter(str(file_path))
        test_data = [
            ["项目", "金额", "状态"],
            ["项目A", 1000, "完成"],
            ["项目B", 2000, "进行中"]
        ]
        result = writer.update_range("数据表!A1:C3", test_data)
        assert result.success is True

        # 3. Read data back
        reader = ExcelReader(str(file_path))
        result = reader.get_range("数据表!A1:C3")
        assert result.success is True
        assert len(result.data) == 3

        # 4. Manage sheets
        manager = ExcelManager(str(file_path))
        result = manager.create_sheet("临时表")
        assert result.success is True

    def test_core_error_handling_consistency(self, sample_excel_file):
        """Test that all core components handle errors consistently"""
        # Reader error handling
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("不存在的工作表!A1:A1")
        assert result.success is False
        assert isinstance(result.error, str)

        # Writer error handling
        writer = ExcelWriter(sample_excel_file)
        result = writer.update_range("不存在的工作表!A1", [["测试"]])
        assert result.success is False
        assert isinstance(result.error, str)

        # Manager error handling
        manager = ExcelManager(sample_excel_file)
        result = manager.delete_sheet("不存在的工作表")
        assert result.success is False
        assert isinstance(result.error, str)

    def test_core_chinese_support(self, sample_excel_file):
        """Test Chinese character support across all core components"""
        # Test Chinese data writing
        writer = ExcelWriter(sample_excel_file)
        chinese_data = [["中文标题", "数值"], ["产品名称", 100]]
        result = writer.update_range("Sheet1!A1:B2", chinese_data)
        assert result.success is True

        # Test Chinese data reading
        reader = ExcelReader(sample_excel_file)
        result = reader.get_range("Sheet1!A1:B2")
        assert result.success is True
        assert len(result.data) == 2

        # Test Chinese sheet management
        manager = ExcelManager(sample_excel_file)
        result = manager.create_sheet("中文工作表名称")
        assert result.success is True
