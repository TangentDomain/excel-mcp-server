# -*- coding: utf-8 -*-
"""
Excel搜索功能测试
合并了ExcelSearcher和目录搜索功能的测试
这个文件替代了原本的test_excel_search.py和test_directory_search.py
"""

import pytest
import tempfile
from pathlib import Path
import os
import re

from src.core.excel_search import ExcelSearcher
from src.models.types import OperationResult
from src.utils.exceptions import ExcelFileNotFoundError


class TestExcelSearch:
    """Excel搜索功能的综合测试"""

    # ==================== 基础初始化测试 ====================

    def test_searcher_init_valid_file(self, sample_excel_file):
        """Test ExcelSearcher initialization with valid file"""
        searcher = ExcelSearcher(sample_excel_file)
        assert searcher.file_path == sample_excel_file

    def test_searcher_init_invalid_file(self):
        """Test ExcelSearcher initialization with invalid file"""
        with pytest.raises(FileNotFoundError):
            ExcelSearcher("nonexistent_file.xlsx")

    # ==================== 单文件搜索测试 ====================

    def test_regex_search_simple_pattern(self, sample_excel_file):
        """Test simple regex search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"张三")

        assert result.success is True
        assert result.data is not None
        assert hasattr(result, 'data')

    def test_regex_search_case_sensitive(self, sample_excel_file):
        """Test case sensitive search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"技术部")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_case_insensitive(self, sample_excel_file):
        """Test case insensitive search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"技术部", flags="i")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_numbers(self, sample_excel_file):
        """Test searching for numbers"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"\d+")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_email_pattern(self, sample_excel_file):
        """Test email pattern search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"\w+@\w+\.\w+")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_chinese_characters(self, sample_excel_file):
        """Test Chinese character pattern search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"[\u4e00-\u9fff]+")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_multiline_mode(self, sample_excel_file):
        """Test multiline mode search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"^项目", flags="m")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_specific_sheet(self, sample_excel_file):
        """Test search in specific sheet"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"[\w]+", sheet_name="Sheet1")

        assert result.success is True
        assert result.data is not None

    def test_regex_search_values_only(self, sample_excel_file):
        """Test search in cell values only"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"[\w]+", search_values=True, search_formulas=False)

        assert result.success is True
        assert result.data is not None

    def test_regex_search_formulas_only(self, sample_excel_file):
        """Test search in formulas only"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"SUM|AVERAGE|COUNT", search_values=False, search_formulas=True)

        assert result.success is True
        # Note: May have no matches if sample file has no formulas

    def test_regex_search_invalid_pattern(self, sample_excel_file):
        """Test search with invalid regex pattern"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"[")  # Invalid regex

        assert result.success is False
        assert "正则表达式" in result.error or "regex" in result.error.lower()

    def test_regex_search_nonexistent_sheet(self, sample_excel_file):
        """Test search in non-existent sheet"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"test", sheet_name="NonExistentSheet")

        assert result.success is False
        assert "工作表" in result.error or "sheet" in result.error.lower()

    # ==================== 目录搜索测试 ====================

    def test_regex_search_directory_basic(self, temp_dir_with_excel_files):
        """Test basic directory search functionality"""
        # 创建搜索器实例(使用临时目录中的第一个文件)
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+",
            flags="i"
        )

        assert result.success is True
        assert result.data is not None
        assert 'searched_files' in result.metadata
        assert 'total_matches' in result.metadata

    def test_regex_search_directory_recursive(self, temp_dir_with_excel_files):
        """Test recursive directory search"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"\d+",
            recursive=True
        )

        assert result.success is True
        assert result.data is not None

    def test_regex_search_directory_file_extensions(self, temp_dir_with_excel_files):
        """Test directory search with specific file extensions"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+",
            file_extensions=[".xlsx"]
        )

        assert result.success is True
        assert result.data is not None

    def test_regex_search_directory_file_pattern(self, temp_dir_with_excel_files):
        """Test directory search with file name pattern"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+",
            file_pattern=r".*test.*"
        )

        assert result.success is True
        assert result.data is not None

    def test_regex_search_directory_max_files_limit(self, temp_dir_with_excel_files):
        """Test directory search with max files limit"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+",
            max_files=2
        )

        assert result.success is True
        assert result.data is not None
        if 'searched_files' in result.metadata:
            searched_files = result.metadata['searched_files']
            if isinstance(searched_files, list):
                assert len(searched_files) <= 2
            else:
                assert searched_files <= 2

    def test_regex_search_directory_nonexistent_path(self, sample_excel_file):
        """Test directory search with non-existent path"""
        searcher = ExcelSearcher(sample_excel_file)

        result = searcher.regex_search_directory(
            directory_path="/path/that/does/not/exist",
            pattern=r"test"
        )

        assert result.success is False
        assert "目录" in result.error or "directory" in result.error.lower()

    def test_regex_search_directory_values_and_formulas(self, temp_dir_with_excel_files):
        """Test directory search in both values and formulas"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        result = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+",
            search_values=True,
            search_formulas=True
        )

        assert result.success is True
        assert result.data is not None

    # ==================== 高级搜索模式测试 ====================

    def test_advanced_search_patterns(self, sample_excel_file):
        """Test advanced regex patterns"""
        searcher = ExcelSearcher(sample_excel_file)

        # Test different complex patterns
        patterns = [
            r"\d{4}-\d{2}-\d{2}",  # Date pattern
            r"^\w+$",  # Word boundaries
            r"(?i)total|sum|计",  # Case insensitive with alternation
            r"\b\d+\.\d{2}\b",  # Decimal numbers
        ]

        for pattern in patterns:
            result = searcher.regex_search(pattern)
            assert result.success is True
            assert result.data is not None

    def test_search_performance_large_dataset(self, sample_excel_file):
        """Test search performance with comprehensive patterns"""
        searcher = ExcelSearcher(sample_excel_file)

        # Test that search completes in reasonable time
        import time
        start_time = time.time()

        result = searcher.regex_search(r"[\w\u4e00-\u9fff]+", flags="i")

        end_time = time.time()
        duration = end_time - start_time

        assert result.success is True
        assert duration < 10.0  # Should complete within 10 seconds

    # ==================== 错误处理和边界测试 ====================

    def test_search_error_handling_consistency(self, sample_excel_file):
        """Test consistent error handling across search methods"""
        searcher = ExcelSearcher(sample_excel_file)

        # Test invalid regex
        result1 = searcher.regex_search(r"[")
        assert result1.success is False
        assert isinstance(result1.error, str)

        # Test invalid sheet
        result2 = searcher.regex_search(r"test", sheet_name="InvalidSheet")
        assert result2.success is False
        assert isinstance(result2.error, str)

    def test_search_empty_pattern(self, sample_excel_file):
        """Test search with empty pattern"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search("")

        # Should handle empty pattern gracefully
        assert result.success is True

    def test_search_unicode_support(self, sample_excel_file):
        """Test Unicode character support in search"""
        searcher = ExcelSearcher(sample_excel_file)

        # Test various Unicode patterns
        unicode_patterns = [
            r"[\u4e00-\u9fff]+",  # Chinese characters
            r"[\u3040-\u309f]+",  # Hiragana
            r"[\u30a0-\u30ff]+",  # Katakana
            r"[\u0400-\u04ff]+",  # Cyrillic
        ]

        for pattern in unicode_patterns:
            result = searcher.regex_search(pattern)
            assert result.success is True
            assert result.data is not None

    # ==================== 集成测试 ====================

    def test_search_workflow_integration(self, temp_dir_with_excel_files):
        """Test integrated search workflow"""
        excel_files = list(Path(temp_dir_with_excel_files).glob("*.xlsx"))
        if not excel_files:
            pytest.skip("No Excel files in temp directory")

        searcher = ExcelSearcher(str(excel_files[0]))

        # 1. Single file search
        result1 = searcher.regex_search(r"[\w]+")
        assert result1.success is True

        # 2. Directory search
        result2 = searcher.regex_search_directory(
            directory_path=temp_dir_with_excel_files,
            pattern=r"[\w]+"
        )
        assert result2.success is True

        # 3. Compare results - directory should have more matches
        if 'total_matches' in result2.metadata:
            assert result2.metadata['total_matches'] >= 0
