"""
Tests for ExcelSearcher class
"""

import pytest
import tempfile
from pathlib import Path
from openpyxl import Workbook
from src.core.excel_search import ExcelSearcher
from src.models.types import OperationResult
from src.utils.exceptions import FileNotFoundError


class TestExcelSearcher:
    """Test cases for ExcelSearcher class"""
    
    def test_init_valid_file(self, sample_excel_file):
        """Test initialization with valid file"""
        searcher = ExcelSearcher(sample_excel_file)
        assert searcher.file_path == sample_excel_file
    
    def test_init_invalid_file(self):
        """Test initialization with invalid file"""
        with pytest.raises(FileNotFoundError):
            ExcelSearcher("nonexistent_file.xlsx")
    
    def test_regex_search_simple_pattern(self, sample_excel_file):
        """Test simple regex search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"张三")
        
        assert result.success is True
        assert result.data.match_count > 0
        assert len(result.data.matches) > 0
        
        # Check first match
        match = result.data.matches[0]
        assert match.matched_text == "张三"
        assert "张三" in match.value
    
    def test_regex_search_case_insensitive(self, sample_excel_file):
        """Test case insensitive search"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"技术部", flags="i")
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should match "技术部"
        matches = [m for m in result.data.matches if "技术部" in m.value]
        assert len(matches) > 0
    
    def test_regex_search_numbers(self, sample_excel_file):
        """Test searching for numbers"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"\d+")
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should match ages and salaries
        numbers = [int(m.matched_text) for m in result.data.matches if m.matched_text.isdigit()]
        assert len(numbers) > 0
        assert all(n > 0 for n in numbers)
    
    def test_regex_search_email_pattern(self, temp_dir):
        """Test searching for email pattern"""
        # Create test file with email addresses
        file_path = temp_dir / "test_emails.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["姓名", "邮箱"])
        ws.append(["张三", "zhangsan@example.com"])
        ws.append(["李四", "lisi@test.org"])
        ws.append(["王五", "wangwu@demo.net"])
        wb.save(file_path)
        
        searcher = ExcelSearcher(str(file_path))
        result = searcher.regex_search(r"\w+@\w+\.\w+")
        
        assert result.success is True
        assert result.data.match_count == 3
        
        # Check that all matches are email-like
        for match in result.data.matches:
            assert "@" in match.matched_text
            assert "." in match.matched_text
    
    def test_regex_search_formulas_only(self, formula_excel_file):
        """Test searching only in formulas"""
        searcher = ExcelSearcher(formula_excel_file)
        result = searcher.regex_search(r"SUM", search_values=False, search_formulas=True)
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should find SUM in formulas
        sum_matches = [m for m in result.data.matches if "SUM" in m.formula]
        assert len(sum_matches) > 0
    
    def test_regex_search_values_only(self, formula_excel_file):
        """Test searching only in values"""
        searcher = ExcelSearcher(formula_excel_file)
        result = searcher.regex_search(r"总计", search_values=True, search_formulas=False)
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should find "总计" in cell values
        total_matches = [m for m in result.data.matches if m.matched_text == "总计"]
        assert len(total_matches) > 0
    
    def test_regex_search_both_values_and_formulas(self, formula_excel_file):
        """Test searching in both values and formulas"""
        searcher = ExcelSearcher(formula_excel_file)
        result = searcher.regex_search(r"AVERAGE", search_values=True, search_formulas=True)
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should find AVERAGE in both values and formulas
        value_matches = [m for m in result.data.matches if "AVERAGE" in str(m.value)]
        formula_matches = [m for m in result.data.matches if "AVERAGE" in str(m.formula)]
        assert len(value_matches) + len(formula_matches) > 0
    
    def test_regex_search_multiline_pattern(self, temp_dir):
        """Test multiline regex pattern"""
        file_path = temp_dir / "test_multiline.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["文本"])
        ws.append(["第一行\n第二行"])
        ws.append(["单行文本"])
        wb.save(file_path)
        
        searcher = ExcelSearcher(str(file_path))
        result = searcher.regex_search(r"第一行.*第二行", flags="s")  # dot matches newline
        
        assert result.success is True
        assert result.data.match_count > 0
    
    def test_regex_search_word_boundaries(self, sample_excel_file):
        """Test regex with word boundaries"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"\b技术部\b")
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should match exact word "技术部"
        matches = [m for m in result.data.matches if m.matched_text == "技术部"]
        assert len(matches) > 0
    
    def test_regex_search_date_pattern(self, temp_dir):
        """Test searching for date pattern"""
        file_path = temp_dir / "test_dates.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["日期"])
        ws.append(["2024-01-15"])
        ws.append(["2024/02/20"])
        ws.append(["2024.03.25"])
        wb.save(file_path)
        
        searcher = ExcelSearcher(str(file_path))
        result = searcher.regex_search(r"\d{4}[-/\.]\d{2}[-/\.]\d{2}")
        
        assert result.success is True
        assert result.data.match_count == 3
    
    def test_regex_search_no_matches(self, sample_excel_file):
        """Test search with no matches"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"不存在的文本")
        
        assert result.success is True
        assert result.data.match_count == 0
        assert len(result.data.matches) == 0
    
    def test_regex_search_empty_pattern(self, sample_excel_file):
        """Test search with empty pattern"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search("")
        
        # Should handle empty pattern gracefully
        assert isinstance(result, OperationResult)
    
    def test_regex_search_special_regex_chars(self, temp_dir):
        """Test searching for special regex characters"""
        file_path = temp_dir / "test_special.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["文本"])
        ws.append(["[特殊]字符"])
        ws.append(["(括号)内容"])
        ws.append(["点.星*"])
        wb.save(file_path)
        
        searcher = ExcelSearcher(str(file_path))
        result = searcher.regex_search(r"\[特殊\]")  # Escaped brackets
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should find "[特殊]"
        matches = [m for m in result.data.matches if "[特殊]" in m.matched_text]
        assert len(matches) > 0
    
    def test_regex_search_multiple_sheets(self, multi_sheet_excel_file):
        """Test search across multiple sheets"""
        searcher = ExcelSearcher(multi_sheet_excel_file)
        result = searcher.regex_search(r"测试数据")
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Should find matches in multiple sheets
        sheet_names = set(m.sheet_name for m in result.data.matches)
        assert len(sheet_names) > 1
    
    def test_regex_search_searched_sheets_info(self, sample_excel_file):
        """Test that searched sheets info is included"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"技术部")
        
        assert result.success is True
        assert hasattr(result.data, 'searched_sheets')
        assert len(result.data.searched_sheets) > 0
        assert "Sheet1" in result.data.searched_sheets
    
    def test_regex_search_coordinates(self, sample_excel_file):
        """Test that match coordinates are correct"""
        searcher = ExcelSearcher(sample_excel_file)
        result = searcher.regex_search(r"张三")
        
        assert result.success is True
        assert result.data.match_count > 0
        
        # Check that coordinates are valid Excel references
        for match in result.data.matches:
            assert isinstance(match.coordinate, str)
            assert len(match.coordinate) >= 2
            assert match.coordinate[0].isalpha()
            assert match.coordinate[1:].isdigit()
    
    def test_regex_search_large_file_performance(self, temp_dir):
        """Test search performance on larger file"""
        file_path = temp_dir / "test_large.xlsx"
        wb = Workbook()
        ws = wb.active
        
        # Add many rows
        for i in range(1000):
            ws.append([f"数据{i}", f"值{i}", "测试"])
        
        wb.save(file_path)
        
        searcher = ExcelSearcher(str(file_path))
        result = searcher.regex_search(r"测试")
        
        assert result.success is True
        assert result.data.match_count == 1000  # Should find "测试" in every row