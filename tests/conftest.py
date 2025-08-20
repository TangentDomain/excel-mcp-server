"""
Excel MCP Server - 测试配置

提供测试用的公共配置、fixture和工具函数
"""

import pytest
import tempfile
import os
from pathlib import Path
from openpyxl import Workbook


@pytest.fixture
def temp_dir():
    """提供临时目录"""
    with tempfile.TemporaryDirectory() as temp_dir:
        yield Path(temp_dir)


@pytest.fixture
def sample_xlsx_file(temp_dir):
    """创建简单的测试Excel文件"""
    file_path = temp_dir / "test_sample.xlsx"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"

    # 添加测试数据
    test_data = [
        ["Name", "Age", "Email"],
        ["Alice", 25, "alice@example.com"],
        ["Bob", 30, "bob@test.org"],
        ["Charlie", 35, "charlie@demo.net"],
        ["David", 28, "david@sample.com"]
    ]

    for row_idx, row_data in enumerate(test_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    workbook.save(file_path)
    return str(file_path)


@pytest.fixture
def multi_sheet_xlsx_file(temp_dir):
    """创建多工作表的测试Excel文件"""
    file_path = temp_dir / "test_multi_sheet.xlsx"

    workbook = Workbook()

    # 删除默认工作表
    workbook.remove(workbook.active)

    # 创建第一个工作表
    sheet1 = workbook.create_sheet(title="Data")
    sheet1['A1'] = "ID"
    sheet1['B1'] = "Value"
    sheet1['A2'] = 1
    sheet1['B2'] = "First"
    sheet1['A3'] = 2
    sheet1['B3'] = "Second"

    # 创建第二个工作表
    sheet2 = workbook.create_sheet(title="Summary")
    sheet2['A1'] = "Total"
    sheet2['B1'] = "=COUNT(Data!A:A)"

    workbook.active = sheet1
    workbook.save(file_path)
    return str(file_path)


@pytest.fixture
def empty_xlsx_file(temp_dir):
    """创建空的Excel文件"""
    file_path = temp_dir / "test_empty.xlsx"

    workbook = Workbook()
    workbook.save(file_path)
    return str(file_path)


@pytest.fixture
def nonexistent_file_path(temp_dir):
    """提供不存在的文件路径"""
    return str(temp_dir / "nonexistent.xlsx")


@pytest.fixture
def invalid_format_file(temp_dir):
    """创建无效格式的文件"""
    file_path = temp_dir / "invalid.txt"
    file_path.write_text("This is not an Excel file")
    return str(file_path)


class TestHelpers:
    """测试辅助工具类"""

    @staticmethod
    def create_test_excel_with_data(file_path: str, data: list, sheet_name: str = "Sheet1"):
        """创建包含指定数据的Excel文件"""
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = sheet_name

        for row_idx, row_data in enumerate(data, 1):
            for col_idx, value in enumerate(row_data, 1):
                sheet.cell(row=row_idx, column=col_idx, value=value)

        workbook.save(file_path)
        return file_path
