"""安全验证测试 - 路径穿越、文件大小、公式注入"""

import os
import tempfile
import pytest
from unittest.mock import patch
from src.excel_mcp_server_fastmcp.server import SecurityValidator, _validate_path


class TestPathTraversal:
    """路径穿越防护"""

    def test_parent_directory_traversal(self):
        """拒绝 ../ 路径穿越"""
        result = _validate_path("../../etc/passwd")
        assert result is not None
        assert 'success' in result
        assert result['success'] is False
        assert '路径穿越' in result['message']

    def test_absolute_path_allowed(self):
        """绝对路径允许（在self-evolution环境中）"""
        # /tmp 下的文件应该允许
        result = _validate_path("/tmp/test.xlsx")
        assert result is None  # 通过验证

    def test_relative_path_allowed(self):
        """普通相对路径允许"""
        result = _validate_path("data.xlsx")
        assert result is None

    def test_nested_traversal(self):
        """深层路径穿越"""
        result = _validate_path("foo/bar/../../../etc/shadow")
        assert result is not None
        assert result['success'] is False
        assert '路径穿越' in result['message']

    def test_empty_path_rejected(self):
        """空路径拒绝"""
        result = _validate_path("")
        assert result is not None
        assert result['success'] is False
        assert '不能为空' in result['message']

    def test_none_path(self):
        """None路径拒绝"""
        result = _validate_path(None)
        assert result is not None
        assert result['success'] is False


class TestFileExtensions:
    """文件扩展名验证"""

    def test_xlsx_allowed(self):
        """xlsx允许"""
        assert _validate_path("data.xlsx") is None

    def test_xlsm_allowed(self):
        """xlsm允许"""
        assert _validate_path("data.xlsm") is None

    def test_csv_allowed(self):
        """csv允许"""
        assert _validate_path("data.csv") is None

    def test_json_allowed(self):
        """json允许"""
        assert _validate_path("data.json") is None

    def test_exe_rejected(self):
        """exe拒绝"""
        result = _validate_path("malware.exe")
        assert result is not None
        assert result['success'] is False
        assert '不支持的文件格式' in result['message']

    def test_py_rejected(self):
        """py拒绝"""
        result = _validate_path("script.py")
        assert result is not None
        assert result['success'] is False

    def test_hidden_file_rejected(self):
        """隐藏文件拒绝"""
        result = _validate_path(".env")
        assert result is not None
        assert result['success'] is False
        assert '隐藏文件' in result['message']

    def test_bak_file_allowed(self):
        """bak备份文件允许"""
        result = _validate_path("data.xlsx.bak")
        assert result is None

    def test_no_extension_allowed(self):
        """无扩展名允许（create_file场景）"""
        result = _validate_path("newfile")
        assert result is None


class TestFileSize:
    """文件大小验证"""

    def test_small_file_allowed(self):
        """小文件允许"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            f.write(b'PK' * 100)  # 200 bytes
            path = f.name
        try:
            result = _validate_path(path)
            assert result is None
        finally:
            os.unlink(path)

    def test_size_check_only_for_existing_files(self):
        """只有已存在的文件才检查大小"""
        result = _validate_path("/tmp/nonexistent_12345.xlsx")
        assert result is None  # 文件不存在，不检查大小


class TestFormulaInjection:
    """公式注入防护"""

    def test_normal_formula_allowed(self):
        """普通公式允许"""
        result = SecurityValidator.validate_formula("SUM(A1:A10)")
        assert result['valid'] is True

    def test_average_formula_allowed(self):
        """AVERAGE公式允许"""
        result = SecurityValidator.validate_formula("AVERAGE(B1:B20)")
        assert result['valid'] is True

    def test_dd_formula_rejected(self):
        """DDE攻击公式拒绝"""
        result = SecurityValidator.validate_formula("DDE(\"cmd\",\"/c calc\",\"\")")
        assert result['valid'] is False
        assert '危险公式' in result['error']

    def test_cmd_formula_rejected(self):
        """CMD命令执行拒绝"""
        result = SecurityValidator.validate_formula('CMD("/c calc")')
        assert result['valid'] is False

    def test_shell_formula_rejected(self):
        """SHELL命令执行拒绝"""
        result = SecurityValidator.validate_formula('SHELL("explorer.exe")')
        assert result['valid'] is False

    def test_power_formula_allowed(self):
        """POWER公式允许（合法Excel函数）"""
        result = SecurityValidator.validate_formula("POWER(2,10)")
        assert result['valid'] is True

    def test_register_formula_rejected(self):
        """REGISTER函数拒绝"""
        result = SecurityValidator.validate_formula('REGISTER("kernel32","...")')
        assert result['valid'] is False

    def test_pipe_format_rejected(self):
        """管道格式拒绝（DDE链接）"""
        result = SecurityValidator.validate_formula('|cmd|"/c calc"!')
        assert result['valid'] is False

    def test_if_formula_allowed(self):
        """IF公式允许"""
        result = SecurityValidator.validate_formula("IF(A1>10,B1,C1)")
        assert result['valid'] is True

    def test_vlookup_allowed(self):
        """VLOOKUP允许"""
        result = SecurityValidator.validate_formula("VLOOKUP(D1,A1:C10,3,FALSE)")
        assert result['valid'] is True

    def test_countif_allowed(self):
        """COUNTIF允许"""
        result = SecurityValidator.validate_formula('COUNTIF(A1:A10,">5")')
        assert result['valid'] is True


class TestSymlink:
    """符号链接防护"""

    def test_symlink_rejected(self):
        """符号链接拒绝"""
        # 创建临时文件和符号链接
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            f.write(b'PK' * 10)
            real_path = f.name
        link_path = real_path + ".link"
        try:
            os.symlink(real_path, link_path)
            result = _validate_path(link_path)
            assert result is not None
            assert result['success'] is False
            assert '符号链接' in result['message']
        finally:
            try:
                os.unlink(link_path)
            except OSError:
                pass
            os.unlink(real_path)


class TestTempFileCleanup:
    """临时文件清理"""

    def test_cleanup_orphan_temp_files(self):
        """清理超过1小时的孤儿临时文件"""
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            # 创建一个"旧"临时文件
            old_file = os.path.join(tmpdir, 'old_backup.xlsx.bak')
            with open(old_file, 'wb') as f:
                f.write(b'PK' * 10)
            # 修改时间设为2小时前
            import time
            two_hours_ago = time.time() - 7200
            os.utime(old_file, (two_hours_ago, two_hours_ago))

            cleaned = SecurityValidator.cleanup_orphan_temp_files(tmpdir)
            assert cleaned == 1
            assert not os.path.exists(old_file)

    def test_cleanup_preserves_recent_files(self):
        """保留最近的临时文件（不到1小时）"""
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            recent_file = os.path.join(tmpdir, 'recent_backup.xlsx.bak')
            with open(recent_file, 'wb') as f:
                f.write(b'PK' * 10)

            cleaned = SecurityValidator.cleanup_orphan_temp_files(tmpdir)
            assert cleaned == 0
            assert os.path.exists(recent_file)

    def test_cleanup_empty_dir(self):
        """空目录不报错"""
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            cleaned = SecurityValidator.cleanup_orphan_temp_files(tmpdir)
            assert cleaned == 0
