"""Tests for new practical tools (batch ops, charts, data validation, conditional formatting)."""
import os
import pytest
from openpyxl import Workbook

def _create_workbook(path, sheet_title, headers, rows):
    """Create a test workbook and save to path."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(headers)
    for row in rows:
        ws.append(row)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path

@pytest.fixture
def workbook(tmp_path):
    """Create a test workbook with sample game data."""
    path = str(tmp_path / "practical_tools_test.xlsx")
    _create_workbook(path, "技能配置",
                     ["skill_id", "skill_name", "skill_type", "damage", "cooldown"],
                     [
                         [1001, "火球术", "法师", 150, 3.0],
                         [1002, "冰冻术", "法师", 120, 4.0],
                         [1003, "斩击", "战士", 200, 2.0],
                         [1004, "治疗术", "牧师", 80, 5.0],
                     ])
    return path

@pytest.fixture
def chart_workbook(tmp_path):
    """Create a test workbook with chart data."""
    path = str(tmp_path / "chart_test.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "角色属性"
    ws.append(["角色名", "等级", "生命值", "魔法值", "攻击力", "防御力"])
    ws.append(["战士", 10, 1200, 100, 85, 60])
    ws.append(["法师", 10, 800, 300, 70, 40])
    ws.append(["牧师", 10, 900, 250, 55, 45])
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path

class TestBatchOperations:
    """Tests for batch operations."""

    def test_batch_update_ranges(self, workbook):
        """Batch update multiple ranges in a single file."""
        from src.excel_mcp_server_fastmcp.server import excel_batch_update_ranges

        updates = [
            {"range": "B2:B3", "data": [["超级火球术"], ["超级冰冻术"]], "sheet": "技能配置"},
            {"range": "D2:D3", "data": [[200], [150]], "sheet": "技能配置"},
        ]

        result = excel_batch_update_ranges(workbook, updates)

        assert result["success"] is True
        assert result["data"]["success_count"] == 2
        assert result["data"]["error_count"] == 0

        # Verify persisted values
        from openpyxl import load_workbook
        wb2 = load_workbook(workbook)
        ws = wb2["技能配置"]
        assert ws["B2"].value == "超级火球术"
        assert ws["D2"].value == 200

    def test_batch_update_partial_failure(self, workbook):
        """Failed range should not block others."""
        from src.excel_mcp_server_fastmcp.server import excel_batch_update_ranges

        updates = [
            {"range": "B2:B2", "data": [["新名称"]], "sheet": "技能配置"},
            {"range": "X99:Z99", "data": [["a", "b", "c"]], "sheet": "不存在的表"},
        ]

        result = excel_batch_update_ranges(workbook, updates)
        assert result["success"] is True
        assert result["data"]["success_count"] == 1
        assert result["data"]["error_count"] == 1

    def test_merge_multiple_files_append(self, tmp_path):
        """Merge two files with append mode."""
        from src.excel_mcp_server_fastmcp.server import excel_merge_multiple_files

        file1 = str(tmp_path / "source1.xlsx")
        _create_workbook(file1, "技能表", ["id", "name"], [[1, "火球术"], [2, "冰冻术"]])

        file2 = str(tmp_path / "source2.xlsx")
        _create_workbook(file2, "装备表", ["id", "name"], [[1, "长剑"], [2, "法杖"]])

        target = str(tmp_path / "merged.xlsx")
        result = excel_merge_multiple_files([file1, file2], target, "append")

        assert result["success"] is True
        assert result["data"]["source_files_count"] == 2
        assert result["data"]["merged_sheets_count"] == 2

        from openpyxl import load_workbook
        wb = load_workbook(target)
        assert "技能表" in wb.sheetnames
        assert "装备表" in wb.sheetnames

class TestChartOperations:
    """Tests for chart creation and listing."""

    def test_create_column_chart(self, chart_workbook):
        """Create a column chart on a sheet."""
        from src.excel_mcp_server_fastmcp.server import excel_create_chart

        result = excel_create_chart(
            chart_workbook, "角色属性", "column", "B1:F4",
            title="角色属性对比图", position="H2"
        )
        assert result["success"] is True
        assert result["data"]["chart_type"] == "column"

    def test_create_pie_chart(self, chart_workbook):
        """Create a pie chart."""
        from src.excel_mcp_server_fastmcp.server import excel_create_chart

        result = excel_create_chart(
            chart_workbook, "角色属性", "pie", "B1:B4", title="角色占比"
        )
        assert result["success"] is True
        assert result["data"]["chart_type"] == "pie"

    def test_invalid_chart_type(self, chart_workbook):
        """Reject unsupported chart types."""
        from src.excel_mcp_server_fastmcp.server import excel_create_chart

        result = excel_create_chart(
            chart_workbook, "角色属性", "radar", "A1:C5"
        )
        assert result["success"] is False
        assert "不支持的图表类型" in result["message"]

class TestDataValidation:
    """Tests for data validation set / clear."""

    def test_set_list_validation(self, workbook):
        """Set a list-type data validation."""
        from src.excel_mcp_server_fastmcp.server import excel_set_data_validation

        result = excel_set_data_validation(
            workbook, "技能配置", "C2:C10", "list",
            "法师,战士,牧师,刺客", "选择职业", "请选择正确的职业类型"
        )
        assert result["success"] is True
        assert result["data"]["validation_type"] == "list"

    def test_set_validation_nonexistent_sheet(self, workbook):
        """Reject non-existent sheet."""
        from src.excel_mcp_server_fastmcp.server import excel_set_data_validation

        result = excel_set_data_validation(
            workbook, "不存在", "A1:A10", "list", "test"
        )
        assert result["success"] is False

    def test_clear_validation(self, workbook):
        """Clear all validations on a sheet."""
        from src.excel_mcp_server_fastmcp.server import (
            excel_set_data_validation,
            excel_clear_validation,
        )

        excel_set_data_validation(
            workbook, "技能配置", "C2:C10", "list", "法师,战士,牧师"
        )

        result = excel_clear_validation(workbook, "技能配置")
        assert result["success"] is True
        assert result["data"]["cleared_count"] >= 1

class TestConditionalFormatting:
    """Tests for conditional formatting add / clear."""

    def test_add_cell_value_format(self, workbook):
        """Add a cell-value conditional format."""
        from src.excel_mcp_server_fastmcp.server import excel_add_conditional_format

        result = excel_add_conditional_format(
            workbook, "技能配置", "D2:D5", "cellValue",
            ">=150", "lightGreen"
        )
        assert result["success"] is True
        assert result["data"]["format_type"] == "cellValue"

    def test_add_format_nonexistent_sheet(self, workbook):
        """Reject non-existent sheet."""
        from src.excel_mcp_server_fastmcp.server import excel_add_conditional_format

        result = excel_add_conditional_format(
            workbook, "不存在", "A1:A10", "cellValue", ">=0"
        )
        assert result["success"] is False

    def test_clear_conditional_format(self, workbook):
        """Clear conditional formatting on a sheet."""
        from src.excel_mcp_server_fastmcp.server import (
            excel_add_conditional_format,
            excel_clear_conditional_format,
        )

        excel_add_conditional_format(
            workbook, "技能配置", "D2:D5", "cellValue", ">=150"
        )

        result = excel_clear_conditional_format(workbook, "技能配置")
        assert result["success"] is True
        assert result["data"]["cleared_count"] >= 1