"""MCP tool layer tests for excel_run_python.

Tests the server.py tool wrapper: validation, decorator chain,
and result format. Does not re-test script_runner internals."""

import pytest

from excel_mcp_server_fastmcp.server import excel_run_python


class TestRunPythonValidation:
    """Test input validation at the MCP tool layer."""

    def test_empty_code_rejected(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "")
        assert result["success"] is False
        assert "code" in result["message"].lower() or "不能为空" in result["message"]

    def test_whitespace_code_rejected(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "   \n\t  ")
        assert result["success"] is False

    def test_nonexistent_file_allowed(self):
        """Nonexistent paths are valid (user script may create the file)."""
        result = excel_run_python("/tmp/_test_nonexistent_xyz.xlsx", "result = 1")
        assert result["success"] is True

    def test_path_traversal_rejected(self, sample_excel_file):
        result = excel_run_python("../../etc/passwd", "print(1)")
        assert result["success"] is False

    def test_timeout_clamped_to_max(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "result = 42", timeout=999)
        assert result["success"] is True
        assert result["data"]["result"] == "42"

    def test_timeout_clamped_to_min(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "result = 1", timeout=0)
        assert result["success"] is True

    def test_invalid_timeout_type(self, sample_excel_file):
        """String timeout should be rejected gracefully, not crash."""
        result = excel_run_python(sample_excel_file, "result = 1", timeout="invalid")
        assert result["success"] is True  # falls back to default 30


class TestRunPythonExecution:
    """Test normal execution through the MCP tool layer."""

    def test_eval_mode(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "2 + 2")
        assert result["success"] is True
        assert result["data"]["result"] == "4"

    def test_exec_mode(self, sample_excel_file):
        code = 'result = "hello"\nprint("world")'
        result = excel_run_python(sample_excel_file, code)
        assert result["success"] is True
        assert result["data"]["result"] == "'hello'"
        assert "world" in result["data"]["stdout"]

    def test_sql_query(self, sample_excel_file):
        code = 'rows = query("SELECT name, age FROM Sheet1")\nresult = rows[1]'
        result = excel_run_python(sample_excel_file, code)
        assert result["success"] is True
        # rows[0] is headers, rows[1] is first data row
        assert result["data"]["result"] is not None

    def test_sql_query_failure_raises(self, sample_excel_file):
        code = 'result = query("SELECT FROM nonexistent")'
        result = excel_run_python(sample_excel_file, code)
        assert result["success"] is False
        assert "SQL" in result["message"] or "sql" in result["message"].lower()

    def test_meta_contains_file_path(self, sample_excel_file):
        result = excel_run_python(sample_excel_file, "42")
        assert result["success"] is True
        assert result["meta"]["file_path"] == sample_excel_file
