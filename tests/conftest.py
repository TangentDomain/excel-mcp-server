"""
Test configuration and fixtures for Excel MCP Server tests
"""

import pytest
import tempfile
import shutil
import uuid
import time
import logging
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Set up logging to suppress warnings
logging.getLogger('openpyxl').setLevel(logging.ERROR)


def safe_rmtree(path, max_retries=3, delay=0.1):
    """Safely remove directory tree with retry mechanism for Windows file locking"""
    for attempt in range(max_retries):
        try:
            shutil.rmtree(path)
            return
        except PermissionError as e:
            if attempt == max_retries - 1:
                # Last attempt failed, log warning but don't fail test
                logging.warning(f"Could not remove temp directory {path}: {e}")
                return
            time.sleep(delay)
            delay *= 2  # Exponential backoff


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files"""
    temp_path = Path(tempfile.mkdtemp())
    yield temp_path
    safe_rmtree(temp_path)


@pytest.fixture
def sample_excel_file(temp_dir, request):
    """Create a sample Excel file for testing with unique name"""
    # Generate unique filename for each test
    test_id = str(uuid.uuid4())[:8]
    test_name = request.node.name
    file_path = temp_dir / f"test_sample_{test_name}_{test_id}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Add sample data
    data = [
        ["姓名", "年龄", "部门", "薪资"],
        ["张三", 25, "技术部", 8000],
        ["李四", 30, "市场部", 9000],
        ["王五", 28, "技术部", 8500],
        ["赵六", 35, "人事部", 9500]
    ]

    for row in data:
        ws.append(row)

    # Add some formulas
    ws["E1"] = "总计"
    ws["E2"] = "=SUM(D2:D5)"

    # Add formatting
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Create second sheet
    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["产品", "销量", "单价"])
    ws2.append(["A", 100, 50])
    ws2.append(["B", 200, 30])

    wb.save(file_path)
    return str(file_path)


@pytest.fixture
def empty_excel_file(temp_dir, request):
    """Create an empty Excel file for testing with unique name"""
    # Generate unique filename for each test
    test_id = str(uuid.uuid4())[:8]
    test_name = request.node.name
    file_path = temp_dir / f"test_empty_{test_name}_{test_id}.xlsx"

    wb = Workbook()
    wb.save(file_path)
    return str(file_path)


@pytest.fixture
def multi_sheet_excel_file(temp_dir, request):
    """Create an Excel file with multiple sheets for testing with unique name"""
    # Generate unique filename for each test
    test_id = str(uuid.uuid4())[:8]
    test_name = request.node.name
    file_path = temp_dir / f"test_multi_sheet_{test_name}_{test_id}.xlsx"

    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Create multiple sheets
    sheet_names = ["数据", "图表", "汇总", "分析"]
    for name in sheet_names:
        ws = wb.create_sheet(name)
        ws.append(["测试数据", "值"])
        ws.append(["项目1", 100])
        ws.append(["项目2", 200])

    wb.save(file_path)
    return str(file_path)


@pytest.fixture
def formula_excel_file(temp_dir):
    """Create an Excel file with various formulas for testing"""
    file_path = temp_dir / "test_formulas.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Formulas"

    # Add data for formulas
    data = [
        ["数值A", "数值B", "和", "积", "平均"],
        [10, 20, "=A2+B2", "=A2*B2", "=AVERAGE(A2:B2)"],
        [30, 40, "=A3+B3", "=A3*B3", "=AVERAGE(A3:B3)"],
        [50, 60, "=A4+B4", "=A4*B4", "=AVERAGE(A4:B4)"]
    ]

    for row in data:
        ws.append(row)

    # Add summary formulas
    ws["A6"] = "总计"
    ws["B6"] = "=SUM(A2:A4)"
    ws["C6"] = "=SUM(C2:C4)"

    wb.save(file_path)
    return str(file_path)
