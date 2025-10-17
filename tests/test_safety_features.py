"""
Excel MCP 服务器安全功能测试用例

本文件包含全面的安全功能测试，验证各种误操作场景的保护机制。
"""

import pytest
import os
import tempfile
import time
import json
from pathlib import Path
from unittest.mock import patch, MagicMock
import pandas as pd

# 导入要测试的模块
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from src.api.excel_operations import ExcelOperations
from src.utils.exceptions import ExcelFileError, SecurityError, OperationCancelledError


class TestSafetyFeatures:
    """安全功能测试类"""

    @pytest.fixture
    def temp_excel_file(self):
        """创建临时Excel文件用于测试"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file_path = tmp.name

        # 创建基本的Excel文件
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        ws.append(["ID", "名称", "类型", "数值", "公式"])
        ws.append([1, "测试1", "A", 100, "=C2*D2"])
        ws.append([2, "测试2", "B", 200, "=C3*D3"])
        ws.append([3, "测试3", "C", 300, "=C4*D4"])

        # 添加一些空行
        for i in range(5, 15):
            ws.append([None, None, None, None, None])

        wb.save(file_path)
        wb.close()

        yield file_path

        # 清理
        if os.path.exists(file_path):
            os.unlink(file_path)

    @pytest.fixture
    def large_temp_excel_file(self):
        """创建大型临时Excel文件用于测试"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file_path = tmp.name

        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "LargeSheet"

        # 添加大量测试数据
        ws.append(["ID", "名称", "类型", "数值", "描述"])
        for i in range(1, 1000):
            ws.append([i, f"项目{i}", chr(65 + (i % 26)), i * 10, f"这是第{i}个测试项目"])

        wb.save(file_path)
        wb.close()

        yield file_path

        # 清理
        if os.path.exists(file_path):
            os.unlink(file_path)

    def test_assess_operation_impact_low_risk(self, temp_excel_file):
        """测试低风险操作的影响评估"""
        result = ExcelOperations.assess_operation_impact(
            file_path=temp_excel_file,
            range_expression="TestSheet!A1:C1",
            operation_type="read",
            preview_data=None
        )

        assert result['success'] is True
        assert result['impact_analysis']['operation_risk_level'] == 'low'
        assert result['impact_analysis']['total_cells'] == 3
        assert result['impact_analysis']['affected_rows'] == 1
        assert result['impact_analysis']['affected_columns'] == 3

    def test_assess_operation_impact_medium_risk(self, temp_excel_file):
        """测试中等风险操作的影响评估"""
        new_data = [["新数据1", "新数据2", "新数据3"]]
        result = ExcelOperations.assess_operation_impact(
            file_path=temp_excel_file,
            range_expression="TestSheet!A10:C15",
            operation_type="update",
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['impact_analysis']['operation_risk_level'] in ['low', 'medium']
        assert result['impact_analysis']['total_cells'] == 18  # 6行 × 3列

    def test_assess_operation_impact_high_risk(self, temp_excel_file):
        """测试高风险操作的影响评估"""
        new_data = [[f"数据{i}" for i in range(1, 11)] for _ in range(50)]
        result = ExcelOperations.assess_operation_impact(
            file_path=temp_excel_file,
            range_expression="TestSheet!A1:J50",
            operation_type="update",
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['impact_analysis']['operation_risk_level'] in ['high', 'critical']
        assert result['impact_analysis']['total_cells'] == 500
        assert result['impact_analysis']['existing_data_count'] > 0

    def test_assess_operation_impact_critical_risk(self, large_temp_excel_file):
        """测试极高风险操作的影响评估"""
        new_data = [[f"新数据{i}" for i in range(1, 11)] for _ in range(900)]
        result = ExcelOperations.assess_operation_impact(
            file_path=large_temp_excel_file,
            range_expression="LargeSheet!A1:J900",
            operation_type="update",
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['impact_analysis']['operation_risk_level'] == 'critical'
        assert result['impact_analysis']['total_cells'] == 9000
        assert result['warnings']  # 应该有警告信息

    def test_check_file_status_normal_file(self, temp_excel_file):
        """测试正常文件的状态检查"""
        result = ExcelOperations.check_file_status(temp_excel_file)

        assert result['success'] is True
        assert result['file_status']['exists'] is True
        assert result['file_status']['readable'] is True
        assert result['file_status']['writable'] is True
        assert result['file_status']['locked'] is False
        assert result['file_status']['file_size'] > 0

    def test_check_file_status_nonexistent_file(self):
        """测试不存在文件的检查"""
        result = ExcelOperations.check_file_status("不存在的文件.xlsx")

        assert result['success'] is False
        assert "文件不存在" in result['error']

    def test_check_file_status_invalid_extension(self):
        """测试无效扩展名文件的检查"""
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            tmp.write(b"test content")
            tmp_path = tmp.name

        try:
            result = ExcelOperations.check_file_status(tmp_path)

            assert result['success'] is False
            assert "不是有效的Excel文件" in result['error']
        finally:
            os.unlink(tmp_path)

    def test_generate_safety_warnings_low_risk(self):
        """测试低风险操作的安全警告生成"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'low',
                'total_cells': 10,
                'existing_data_count': 0
            }
        }

        warnings = ExcelOperations._generate_safety_warnings(impact)

        assert len(warnings) == 1
        assert warnings[0]['level'] == 'info'
        assert "低风险" in warnings[0]['message']

    def test_generate_safety_warnings_high_risk(self):
        """测试高风险操作的安全警告生成"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'high',
                'total_cells': 1000,
                'existing_data_count': 500,
                'formula_count': 50
            }
        }

        warnings = ExcelOperations._generate_safety_warnings(impact)

        assert len(warnings) > 1
        # 应该有高风险警告
        high_risk_warnings = [w for w in warnings if w['level'] == 'error']
        assert len(high_risk_warnings) > 0

    def test_generate_operation_visualization_small_range(self):
        """测试小范围操作的可视化生成"""
        impact = {
            'impact_analysis': {
                'total_cells': 25,
                'affected_rows': 5,
                'affected_columns': 5
            },
            'current_data_preview': [
                ['A1', 'B1', 'C1', 'D1', 'E1'],
                ['A2', 'B2', 'C2', 'D2', 'E2'],
            ]
        }

        viz = ExcelOperations._generate_operation_visualization(impact)

        assert viz['type'] == 'text_grid'
        assert 'grid' in viz
        assert len(viz['grid']['rows']) > 0

    def test_generate_operation_visualization_large_range(self):
        """测试大范围操作的可视化生成"""
        impact = {
            'impact_analysis': {
                'total_cells': 5000,
                'affected_rows': 100,
                'affected_columns': 50
            }
        }

        viz = ExcelOperations._generate_operation_visualization(impact)

        assert viz['type'] == 'statistics'
        assert 'statistics' in viz
        assert viz['statistics']['total_cells'] == 5000

    def test_confirm_operation_low_risk_auto_approve(self):
        """测试低风险操作自动确认"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'low'
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C1",
            operation_type="read",
            impact_analysis=impact
        )

        assert result['can_proceed'] is True
        assert result['confirmation_required'] is False

    def test_confirm_operation_high_risk_requires_confirmation(self):
        """测试高风险操作需要确认"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'high',
                'total_cells': 1000,
                'existing_data_count': 800
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:Z100",
            operation_type="update",
            impact_analysis=impact
        )

        assert result['confirmation_required'] is True
        assert 'confirmation_token' in result
        assert len(result['safety_steps']) > 0

    def test_confirm_operation_with_valid_token(self):
        """测试使用有效确认令牌"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        # 首先获取确认令牌
        confirm_result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:Z100",
            operation_type="update",
            impact_analysis=impact
        )

        token = confirm_result['confirmation_token']

        # 使用令牌确认操作
        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:Z100",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token
        )

        assert result['can_proceed'] is True
        assert result['confirmed'] is True

    def test_confirm_operation_with_invalid_token(self):
        """测试使用无效确认令牌"""
        impact = {
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:Z100",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token="invalid_token"
        )

        assert result['can_proceed'] is False
        assert "无效的确认令牌" in result['message']

    def test_operation_manager_singleton(self):
        """测试操作管理器的单例模式"""
        manager1 = ExcelOperations.get_operation_manager()
        manager2 = ExcelOperations.get_operation_manager()

        assert manager1 is manager2

    def test_operation_manager_start_operation(self):
        """测试启动操作"""
        manager = ExcelOperations.get_operation_manager()

        operation_id = manager.start_operation(
            operation_type="update",
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C10"
        )

        assert operation_id is not None
        assert manager.is_operation_running(operation_id) is True

        # 清理
        manager.cancel_operation(operation_id)

    def test_operation_manager_cancel_operation(self):
        """测试取消操作"""
        manager = ExcelOperations.get_operation_manager()

        operation_id = manager.start_operation(
            operation_type="update",
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C10"
        )

        # 取消操作
        result = manager.cancel_operation(operation_id)

        assert result['success'] is True
        assert manager.is_operation_running(operation_id) is False

    def test_operation_manager_get_operation_status(self):
        """测试获取操作状态"""
        manager = ExcelOperations.get_operation_manager()

        operation_id = manager.start_operation(
            operation_type="update",
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C10"
        )

        status = manager.get_operation_status(operation_id)

        assert status['operation_id'] == operation_id
        assert status['operation_type'] == "update"
        assert status['file_path'] == "test.xlsx"
        assert status['status'] == 'running'

        # 清理
        manager.cancel_operation(operation_id)

    def test_auto_backup_creation(self, temp_excel_file):
        """测试自动备份创建"""
        result = ExcelOperations.create_auto_backup(
            file_path=temp_excel_file,
            backup_name="test_backup",
            backup_reason="测试备份",
            user_id="test_user"
        )

        assert result['success'] is True
        assert 'backup_path' in result
        assert 'backup_id' in result
        assert result['backup_reason'] == "测试备份"

        # 验证备份文件存在
        assert os.path.exists(result['backup_path'])

        # 清理备份文件
        if os.path.exists(result['backup_path']):
            os.unlink(result['backup_path'])

    def test_auto_backup_with_checksum(self, temp_excel_file):
        """测试带校验和的自动备份"""
        result = ExcelOperations.create_auto_backup(
            file_path=temp_excel_file,
            backup_name="checksum_test",
            backup_reason="校验和测试",
            user_id="test_user"
        )

        assert result['success'] is True
        assert 'checksum' in result
        assert len(result['checksum']) == 64  # SHA-256 长度

        # 清理备份文件
        if os.path.exists(result['backup_path']):
            os.unlink(result['backup_path'])

    def test_update_range_with_safety_checks_low_risk(self, temp_excel_file):
        """测试低风险更新操作的安全检查"""
        new_data = [["新数据", "新数据2", "新数据3"]]

        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="TestSheet!A10:C10",
            data=new_data,
            insert_mode=True,
            require_confirmation=False,
            skip_safety_checks=False
        )

        assert result['success'] is True

    def test_update_range_with_safety_checks_high_risk_requires_confirmation(self, temp_excel_file):
        """测试高风险更新操作需要确认"""
        new_data = [[f"数据{i}" for i in range(1, 11)] for _ in range(50)]

        # 这应该失败，因为高风险操作需要确认
        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="TestSheet!A1:J50",
            data=new_data,
            insert_mode=True,
            require_confirmation=True,  # 需要确认
            skip_safety_checks=False
        )

        assert result['success'] is False
        assert "需要用户确认" in result['error']

    def test_update_range_skip_safety_checks(self, temp_excel_file):
        """测试跳过安全检查"""
        new_data = [["快速数据"]]

        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="TestSheet!A20:A20",
            data=new_data,
            insert_mode=True,
            require_confirmation=False,
            skip_safety_checks=True  # 跳过安全检查
        )

        assert result['success'] is True

    def test_get_safe_operation_guidance(self):
        """测试获取安全操作指导"""
        guidance = ExcelOperations.get_safe_operation_guidance(
            operation_goal="更新技能表数据",
            file_path="skills.xlsx"
        )

        assert guidance['success'] is True
        assert 'recommended_steps' in guidance
        assert 'safety_considerations' in guidance
        assert len(guidance['recommended_steps']) > 0

    def test_validate_range_expression_valid(self):
        """测试有效范围表达式验证"""
        result = ExcelOperations._validate_range_expression("Sheet1!A1:C10")

        assert result['valid'] is True
        assert result['sheet_name'] == "Sheet1"
        assert result['range'] == "A1:C10"

    def test_validate_range_expression_invalid(self):
        """测试无效范围表达式验证"""
        result = ExcelOperations._validate_range_expression("invalid_range")

        assert result['valid'] is False
        assert "不是有效的范围表达式" in result['error']

    def test_validate_range_expression_missing_sheet(self):
        """测试缺少工作表名的范围表达式"""
        result = ExcelOperations._validate_range_expression("A1:C10")

        assert result['valid'] is False
        assert "必须包含工作表名" in result['error']

    def test_comprehensive_safety_workflow(self, temp_excel_file):
        """测试完整的安全工作流程"""
        # 1. 获取安全指导
        guidance = ExcelOperations.get_safe_operation_guidance(
            operation_goal="更新测试数据",
            file_path=temp_excel_file
        )
        assert guidance['success'] is True

        # 2. 检查文件状态
        file_status = ExcelOperations.check_file_status(temp_excel_file)
        assert file_status['success'] is True

        # 3. 评估操作影响（小范围操作）
        new_data = [["测试数据1", "测试数据2", "测试数据3"]]
        impact = ExcelOperations.assess_operation_impact(
            file_path=temp_excel_file,
            range_expression="TestSheet!A15:C15",
            operation_type="update",
            preview_data=new_data
        )
        assert impact['success'] is True

        # 4. 确认操作（低风险应该自动通过）
        confirmation = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestSheet!A15:C15",
            operation_type="update",
            impact_analysis=impact
        )
        assert confirmation['can_proceed'] is True

        # 5. 执行操作
        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="TestSheet!A15:C15",
            data=new_data,
            insert_mode=True,
            require_confirmation=False,
            skip_safety_checks=False
        )
        assert result['success'] is True

        # 6. 验证结果
        verification = ExcelOperations.get_range(
            file_path=temp_excel_file,
            range="TestSheet!A15:C15"
        )
        assert verification['success'] is True
        assert verification['data'] == new_data


class TestErrorHandling:
    """错误处理测试类"""

    def test_file_not_found_error(self):
        """测试文件不存在错误"""
        with pytest.raises(ExcelFileError):
            ExcelOperations.check_file_status("不存在的文件.xlsx")

    def test_invalid_range_error(self, temp_excel_file):
        """测试无效范围错误"""
        result = ExcelOperations._validate_range_expression("invalid_range")
        assert result['valid'] is False

    def test_security_error_high_risk_without_confirmation(self, temp_excel_file):
        """测试高风险操作无确认的安全错误"""
        new_data = [[f"数据{i}" for i in range(1, 11)] for _ in range(50)]

        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="TestSheet!A1:J50",
            data=new_data,
            insert_mode=True,
            require_confirmation=True,  # 需要确认但不提供
            skip_safety_checks=False
        )

        assert result['success'] is False
        assert "需要用户确认" in result['error']

    def test_operation_cancelled_error(self):
        """测试操作取消错误"""
        manager = ExcelOperations.get_operation_manager()

        operation_id = manager.start_operation(
            operation_type="update",
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C10"
        )

        # 取消操作
        manager.cancel_operation(operation_id)

        # 尝试获取已取消操作的状态
        status = manager.get_operation_status(operation_id)
        assert status['status'] == 'cancelled'


if __name__ == "__main__":
    # 运行测试
    pytest.main([__file__, "-v"])