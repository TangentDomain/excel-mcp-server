"""
安全功能测试套件

测试所有安全改进功能，确保数据安全和操作可靠性
"""

import pytest
import tempfile
import os
from unittest.mock import patch, MagicMock
from datetime import datetime

from src.excel_mcp_server_fastmcp.server import (
    OperationLogger,
    excel_create_backup,
    excel_restore_backup,
    excel_list_backups,
)
from src.excel_mcp_server_fastmcp.utils.validators import ExcelValidator, DataValidationError
from src.excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations

class TestOperationLogger:
    """测试操作日志记录器"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.logger = OperationLogger()
        self.test_file = "test_operations.xlsx"

    def test_start_session(self):
        """测试开始操作会话"""
        self.logger.start_session(self.test_file)

        assert self.logger.current_session is not None
        assert len(self.logger.current_session) == 1
        assert self.logger.current_session[0]['file_path'] == self.test_file
        assert 'session_id' in self.logger.current_session[0]
        assert 'operations' in self.logger.current_session[0]

    def test_log_operation(self):
        """测试记录操作"""
        self.logger.start_session(self.test_file)

        operation_details = {
            "range": "Sheet1!A1:C10",
            "data_rows": 10
        }

        self.logger.log_operation("update_range", operation_details)

        operations = self.logger.current_session[0]['operations']
        assert len(operations) == 1
        assert operations[0]['operation'] == "update_range"
        assert operations[0]['details'] == operation_details
        assert 'timestamp' in operations[0]

    def test_get_recent_operations(self):
        """测试获取最近操作"""
        self.logger.start_session(self.test_file)

        # 记录多个操作
        for i in range(5):
            self.logger.log_operation(f"operation_{i}", {"step": i})

        recent_ops = self.logger.get_recent_operations(3)
        assert len(recent_ops) == 3
        assert recent_ops[-1]['operation'] == "operation_4"

    def test_get_recent_operations_empty_session(self):
        """测试空会话时获取操作"""
        recent_ops = self.logger.get_recent_operations()
        assert recent_ops == []

class TestEnhancedRangeValidation:
    """测试增强的范围验证功能"""

    def test_validate_valid_ranges(self):
        """测试有效范围格式"""
        valid_ranges = [
            "Sheet1!A1:C10",
            "Data!1:100",
            "Report!A:Z",
            "Summary!5",
            "Analysis!B",
            "技能配置表!A1:Z100"
        ]

        for range_expr in valid_ranges:
            result = ExcelValidator.validate_range_expression(range_expr)
            assert result['success'] is True
            assert 'sheet_name' in result
            assert 'range_info' in result
            assert 'normalized_range' in result

    def test_validate_invalid_ranges(self):
        """测试无效范围格式"""
        invalid_ranges = [
            "A1:C10",  # 缺少工作表名
            "",        # 空字符串
            None,      # None值
            "Sheet1!", # 缺少范围部分
            "Sheet1!INVALID_FORMAT"
        ]

        for range_expr in invalid_ranges:
            with pytest.raises(DataValidationError):
                ExcelValidator.validate_range_expression(range_expr)

    def test_operation_scale_validation(self):
        """测试操作规模验证"""
        # 低风险
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 5, 'start_row': 1, 'end_row': 10}
        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['risk_level'] == "LOW"

        # 中等风险
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 20, 'start_row': 1, 'end_row': 100}
        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['risk_level'] == "MEDIUM"

        # 高风险
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 50, 'start_row': 1, 'end_row': 300}
        result = ExcelValidator.validate_operation_scale(range_info)
        assert result['risk_level'] == "HIGH"

    def test_scale_limits_enforcement(self):
        """测试规模限制强制执行"""
        # 超过行数限制
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 10, 'start_row': 1, 'end_row': 2000}
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_operation_scale(range_info)

        # 超过列数限制
        range_info = {'type': 'cell_range', 'start_col': 1, 'end_col': 200, 'start_row': 1, 'end_row': 10}
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_operation_scale(range_info)

class TestBackupSystem:
    """测试备份系统功能"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_data.xlsx")

    def teardown_method(self):
        """每个测试方法后的清理"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_create_backup_success(self):
        """测试成功创建备份"""
        # 创建测试文件
        with open(self.test_file, 'w') as f:
            f.write("test content")

        result = excel_create_backup(self.test_file)

        assert result['success'] is True
        assert 'backup_file' in result['data']
        assert 'timestamp' in result['data']
        assert os.path.exists(result['data']['backup_file'])
        assert result['data']['file_size']['original'] == result['data']['file_size']['backup']

    def test_list_backups(self):
        """测试列出备份文件"""
        import time

        # 创建测试文件
        with open(self.test_file, 'w') as f:
            f.write("test content")

        # 创建多个备份，确保时间戳不同
        excel_create_backup(self.test_file)
        time.sleep(0.1)  # 确保时间戳不同
        excel_create_backup(self.test_file)

        result = excel_list_backups(self.test_file)

        assert result['success'] is True
        assert result['data']['total_backups'] >= 1  # 至少有一个备份
        assert len(result['data']['backups']) >= 1
        assert all('filename' in backup for backup in result['data']['backups'])

    def test_list_backups_no_backups(self):
        """测试列出不存在的备份"""
        result = excel_list_backups(self.test_file)

        assert result['success'] is True
        assert result['data'].get('total_backups', 0) == 0
        assert result['data'].get('backups', []) == []

class TestSecurityPerformance:
    """测试安全功能的性能"""

    def test_large_scale_validation_performance(self):
        """测试大规模验证的性能"""
        import time

        # 测试各种范围格式的验证性能
        test_ranges = [
            "Sheet1!A1:Z1000",
            "Data!1:10000",
            "Report!A:ZZ",
            "Analysis!1:100000"
        ]

        start_time = time.time()

        for range_expr in test_ranges:
            try:
                result = ExcelValidator.validate_range_expression(range_expr)
                if result['success']:
                    scale_info = ExcelValidator.validate_operation_scale(result['range_info'])
            except DataValidationError:
                pass  # 某些范围可能无效，这是正常的

        end_time = time.time()

        # 验证性能要求（所有验证应在1秒内完成）
        assert (end_time - start_time) < 1.0

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
