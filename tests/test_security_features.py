"""
安全功能测试套件

测试所有安全改进功能，确保数据安全和操作可靠性
"""

import pytest
import tempfile
import os
from unittest.mock import patch, MagicMock
from datetime import datetime

from src.server import (
    OperationLogger,
    excel_preview_operation,
    excel_assess_data_impact,
    excel_create_backup,
    excel_restore_backup,
    excel_list_backups,
    excel_get_operation_history,
    _analyze_current_data,
    _assess_operation_risk,
    _generate_safety_recommendations,
    _predict_operation_result
)
from src.utils.validators import ExcelValidator, DataValidationError
from src.api.excel_operations import ExcelOperations


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


class TestDataAnalysis:
    """测试数据分析功能"""

    def test_analyze_empty_data(self):
        """测试分析空数据"""
        result = _analyze_current_data([])

        assert result['row_count'] == 0
        assert result['column_count'] == 0
        assert result['non_empty_cell_count'] == 0
        assert result['completeness_rate'] == 0.0

    def test_analyze_numeric_data(self):
        """测试分析数值数据"""
        data = [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]

        result = _analyze_current_data(data)

        assert result['row_count'] == 3
        assert result['column_count'] == 3
        assert result['non_empty_cell_count'] == 9
        assert result['has_numeric_data'] is True
        assert result['has_text_data'] is False
        assert result['completeness_rate'] == 100.0
        assert result['data_types']['numeric'] == 9

    def test_analyze_mixed_data(self):
        """测试分析混合类型数据"""
        data = [
            [1, "Text", "=SUM(A1)"],
            [None, "", "Another Text"],
            [3.14, "中文", 42]
        ]

        result = _analyze_current_data(data)

        assert result['row_count'] == 3
        assert result['column_count'] == 3
        assert result['non_empty_cell_count'] == 7  # 1, Text, =SUM, Another Text, 3.14, 中文, 42
        assert result['has_numeric_data'] is True
        assert result['has_text_data'] is True
        assert result['has_formulas'] is True
        assert result['data_types']['numeric'] == 3  # 1, 3.14, 42
        assert result['data_types']['text'] == 3  # Text, Another Text, 中文
        assert result['data_types']['formulas'] == 1  # =SUM

    def test_analyze_data_completeness(self):
        """测试数据完整性分析"""
        # 部分填充的数据
        data = [
            [1, None, 3],
            [None, 5, None],
            [7, None, 9]
        ]

        result = _analyze_current_data(data)

        assert result['total_cells'] == 9
        assert result['non_empty_cell_count'] == 5
        assert result['empty_cell_count'] == 4
        assert result['completeness_rate'] == (5/9) * 100


class TestRiskAssessment:
    """测试风险评估功能"""

    def test_assess_low_risk_operation(self):
        """测试低风险操作评估"""
        data_analysis = {
            'non_empty_cell_count': 0,
            'has_formulas': False,
            'completeness_rate': 0.0
        }
        scale_info = {'total_cells': 10}

        result = _assess_operation_risk("update", data_analysis, scale_info)

        assert result['overall_risk'] == "LOW"
        assert result['risk_score'] < 30
        assert result['requires_backup'] is False
        assert result['requires_confirmation'] is False

    def test_assess_high_risk_operation(self):
        """测试高风险操作评估"""
        data_analysis = {
            'non_empty_cell_count': 100,
            'has_formulas': True,
            'completeness_rate': 90.0
        }
        scale_info = {'total_cells': 15000}

        result = _assess_operation_risk("delete", data_analysis, scale_info)

        assert result['overall_risk'] == "HIGH"
        assert result['risk_score'] >= 60
        assert result['requires_backup'] is True
        assert result['requires_confirmation'] is True
        assert "删除操作不可逆" in result['risk_factors']
        assert "包含公式数据" in result['risk_factors']

    def test_safety_recommendations_generation(self):
        """测试安全建议生成"""
        data_analysis = {
            'has_formulas': True,
            'completeness_rate': 85.0,
            'non_empty_cell_count': 50
        }
        risk_assessment = {
            'requires_backup': True,
            'requires_confirmation': True,
            'overall_risk': "HIGH"
        }
        scale_info = {'total_cells': 5000}

        recommendations = _generate_safety_recommendations(
            "update", data_analysis, risk_assessment, scale_info
        )

        assert len(recommendations) > 0
        assert any("建议在操作前创建备份" in rec for rec in recommendations)
        assert any("高风险操作" in rec for rec in recommendations)
        assert any("公式数据" in rec for rec in recommendations)

    def test_operation_result_prediction(self):
        """测试操作结果预测"""
        current_data = [[1, 2], [3, 4]]
        new_data = [[5, 6], [7, 8]]
        scale_info = {'total_cells': 4}

        result = _predict_operation_result("update", current_data, new_data, scale_info)

        assert result['affected_cells'] == 4
        assert result['data_overwrite_count'] == 4
        assert result['data_insert_count'] == 4
        assert result['estimated_time'] == "minimal"


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
        assert 'backup_file' in result
        assert 'timestamp' in result
        assert os.path.exists(result['backup_file'])
        assert result['file_size']['original'] == result['file_size']['backup']

    def test_create_backup_nonexistent_file(self):
        """测试创建不存在文件的备份"""
        result = excel_create_backup("nonexistent.xlsx")

        assert result['success'] is False
        assert result['error'] == 'FILE_NOT_FOUND'
        assert "不存在" in result['message']

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
        assert result['total_backups'] >= 1  # 至少有一个备份
        assert len(result['backups']) >= 1
        assert all('filename' in backup for backup in result['backups'])

    def test_list_backups_no_backups(self):
        """测试列出不存在的备份"""
        result = excel_list_backups(self.test_file)

        assert result['success'] is True
        assert result.get('total_backups', 0) == 0
        assert result.get('backups', []) == []


class TestIntegrationSecurityWorkflow:
    """测试完整的安全工作流程集成"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "security_test.xlsx")

    def teardown_method(self):
        """每个测试方法后的清理"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    @patch('src.server.ExcelOperations.get_range')
    def test_complete_safe_update_workflow(self, mock_get_range):
        """测试完整的安全更新工作流程"""
        # 模拟现有数据
        mock_get_range.return_value = {
            'success': True,
            'data': [[1, 2, 3], [4, 5, 6]]
        }

        # 1. 数据影响评估
        assessment_result = excel_assess_data_impact(
            self.test_file,
            "Sheet1!A1:C3",
            "update",
            [[7, 8, 9], [10, 11, 12]]
        )

        assert assessment_result['success'] is True
        assert 'risk_assessment' in assessment_result
        assert 'safety_recommendations' in assessment_result

        # 2. 操作预览
        preview_result = excel_preview_operation(
            self.test_file,
            "Sheet1!A1:C3",
            "update"
        )

        assert preview_result['success'] is True
        assert 'impact_assessment' in preview_result
        assert 'safety_warning' in preview_result

        # 3. 验证风险评估一致性
        risk_level = assessment_result['risk_assessment']['overall_risk']
        assert risk_level in ["LOW", "MEDIUM", "HIGH"]

    def test_validation_prevents_dangerous_operations(self):
        """测试验证机制阻止危险操作"""
        # 测试无效范围格式 - 应该返回错误而不是抛出异常
        result = excel_assess_data_impact(
            self.test_file,
            "invalid_range",  # 无效格式
            "update"
        )
        assert result['success'] is False
        assert result['error'] == 'VALIDATION_FAILED'

        # 测试范围验证 - 这个会抛出异常
        with pytest.raises(DataValidationError):
            ExcelValidator.validate_range_expression("A1:C10")  # 缺少工作表名

    @patch('src.server.ExcelOperations.get_range')
    def test_high_risk_operation_detection(self, mock_get_range):
        """测试高风险操作检测"""
        # 模拟大量数据
        large_data = [[f"cell_{i}_{j}" for j in range(100)] for i in range(100)]
        mock_get_range.return_value = {
            'success': True,
            'data': large_data
        }

        # 大范围操作应该被识别为高风险
        assessment_result = excel_assess_data_impact(
            self.test_file,
            "Sheet1!A1:CV100",  # 100x100 = 10000 cells
            "delete"
        )

        assert assessment_result['success'] is True
        assert assessment_result['risk_assessment']['overall_risk'] in ["HIGH", "MEDIUM"]
        assert assessment_result['risk_assessment']['requires_backup'] is True

    def test_operation_history_tracking(self):
        """测试操作历史跟踪"""
        # 创建临时文件用于测试
        with open(self.test_file, 'w') as f:
            f.write("test")

        # 获取操作历史
        history_result = excel_get_operation_history(self.test_file, 10)

        assert history_result['success'] is True
        assert 'operations' in history_result
        assert 'statistics' in history_result

        # 验证统计信息结构
        stats = history_result['statistics']
        assert 'total_operations' in stats
        assert 'operation_types' in stats
        assert 'success_rate' in stats


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

    def test_data_analysis_performance(self):
        """测试数据分析的性能"""
        import time

        # 创建大型数据集
        large_data = [[f"data_{i}_{j}" for j in range(50)] for i in range(100)]

        start_time = time.time()
        result = _analyze_current_data(large_data)
        end_time = time.time()

        # 验证分析结果正确性
        assert result['row_count'] == 100
        assert result['column_count'] == 50
        assert result['non_empty_cell_count'] == 5000

        # 验证性能要求（大型数据分析应在0.5秒内完成）
        assert (end_time - start_time) < 0.5


if __name__ == "__main__":
    pytest.main([__file__, "-v"])