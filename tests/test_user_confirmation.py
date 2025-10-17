"""
Excel MCP 服务器用户确认流程测试

本文件包含用户确认机制的各种场景测试，验证确认流程的正确性和安全性。
"""

import pytest
import os
import tempfile
import time
import json
from unittest.mock import patch, MagicMock
import uuid

# 导入要测试的模块
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from src.api.excel_operations import ExcelOperations
from src.utils.exceptions import SecurityError, OperationCancelledError


class TestConfirmationWorkflow:
    """确认工作流程测试类"""

    @pytest.fixture
    def temp_excel_file(self):
        """创建临时Excel文件用于测试"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file_path = tmp.name

        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "TestData"

        # 添加测试数据
        ws.append(["ID", "名称", "类型", "数值", "公式"])
        ws.append([1, "项目1", "A", 100, "=C2*D2"])
        ws.append([2, "项目2", "B", 200, "=C3*D3"])
        ws.append([3, "项目3", "C", 300, "=C4*D4"])

        wb.save(file_path)
        wb.close()

        yield file_path

        # 清理
        if os.path.exists(file_path):
            os.unlink(file_path)

    def test_low_risk_operation_auto_approval(self, temp_excel_file):
        """测试低风险操作自动批准"""
        # 低风险操作的影响评估
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'low',
                'total_cells': 3,
                'affected_rows': 1,
                'affected_columns': 3,
                'existing_data_count': 0,
                'formula_count': 0
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A10:C10",
            operation_type="update",
            impact_analysis=impact
        )

        assert result['success'] is True
        assert result['can_proceed'] is True
        assert result['confirmation_required'] is False
        assert 'confirmation_token' not in result

    def test_medium_risk_operation_recommend_preview(self, temp_excel_file):
        """测试中等风险操作建议预览"""
        new_data = [["数据1", "数据2", "数据3", "数据4", "数据5"]]

        # 中等风险操作的影响评估
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'medium',
                'total_cells': 20,
                'affected_rows': 4,
                'affected_columns': 5,
                'existing_data_count': 5,
                'formula_count': 1
            },
            'current_data_preview': [
                ["现有1", "现有2", "现有3", "现有4", "现有5"]
            ],
            'preview_data': new_data
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A5:E8",
            operation_type="update",
            impact_analysis=impact,
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['can_proceed'] is True  # 中等风险可以直接执行
        assert result['confirmation_required'] is False
        assert 'recommendations' in result
        assert any("预览" in rec for rec in result['recommendations'])

    def test_high_risk_operation_requires_confirmation(self, temp_excel_file):
        """测试高风险操作需要确认"""
        new_data = [[f"数据{i}" for i in range(1, 11)] for _ in range(20)]

        # 高风险操作的影响评估
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high',
                'total_cells': 200,
                'affected_rows': 20,
                'affected_columns': 10,
                'existing_data_count': 100,
                'formula_count': 20
            },
            'current_data_preview': [
                ["现有数据1", "现有数据2", "现有数据3"]
            ],
            'warnings': [
                {
                    'level': 'warning',
                    'message': '操作将覆盖大量现有数据'
                }
            ]
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J20",
            operation_type="update",
            impact_analysis=impact,
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['can_proceed'] is False  # 需要确认
        assert result['confirmation_required'] is True
        assert 'confirmation_token' in result
        assert 'safety_steps' in result
        assert len(result['safety_steps']) > 0

    def test_critical_risk_operation_multiple_confirmations(self, temp_excel_file):
        """测试极高风险操作需要多重确认"""
        new_data = [[f"数据{i}" for i in range(1, 21)] for _ in range(100)]

        # 极高风险操作的影响评估
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'critical',
                'total_cells': 2000,
                'affected_rows': 100,
                'affected_columns': 20,
                'existing_data_count': 1500,
                'formula_count': 200
            },
            'current_data_preview': [
                ["重要数据1", "重要数据2", "重要数据3"]
            ],
            'warnings': [
                {
                    'level': 'error',
                    'message': '极高风险操作：可能导致数据永久丢失'
                },
                {
                    'level': 'warning',
                    'message': '操作将覆盖大量包含公式的重要数据'
                }
            ]
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:T100",
            operation_type="update",
            impact_analysis=impact,
            preview_data=new_data
        )

        assert result['success'] is True
        assert result['can_proceed'] is False
        assert result['confirmation_required'] is True
        assert result['risk_level'] == 'critical'
        assert len(result['safety_steps']) >= 3  # 多重确认步骤
        assert 'manual_backup_required' in result
        assert result['manual_backup_required'] is True

    def test_confirmation_token_validation(self, temp_excel_file):
        """测试确认令牌验证"""
        # 创建需要确认的操作
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high',
                'total_cells': 100,
                'existing_data_count': 50
            }
        }

        # 第一次确认：生成令牌
        result1 = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J10",
            operation_type="update",
            impact_analysis=impact
        )

        assert result1['success'] is True
        assert result1['confirmation_required'] is True
        token = result1['confirmation_token']

        # 第二次确认：使用有效令牌
        result2 = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J10",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token
        )

        assert result2['success'] is True
        assert result2['can_proceed'] is True
        assert result2['confirmed'] is True

    def test_invalid_confirmation_token(self, temp_excel_file):
        """测试无效确认令牌"""
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J10",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token="invalid_token_12345"
        )

        assert result['success'] is False
        assert result['can_proceed'] is False
        assert "无效的确认令牌" in result['error']

    def test_expired_confirmation_token(self, temp_excel_file):
        """测试过期确认令牌"""
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        # 生成令牌
        result1 = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J10",
            operation_type="update",
            impact_analysis=impact
        )

        token = result1['confirmation_token']

        # 模拟令牌过期（修改令牌创建时间）
        with patch('time.time', return_value=time.time() + 1800):  # 30分钟后
            result2 = ExcelOperations.confirm_operation(
                file_path=temp_excel_file,
                range_expression="TestData!A1:J10",
                operation_type="update",
                impact_analysis=impact,
                confirmation_token=token
            )

            assert result2['success'] is False
            assert result2['can_proceed'] is False
            assert "已过期" in result2['error']

    def test_confirmation_operation_mismatch(self, temp_excel_file):
        """测试确认操作不匹配"""
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        # 为操作A生成令牌
        result1 = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:J10",
            operation_type="update",
            impact_analysis=impact
        )

        token = result1['confirmation_token']

        # 尝试对操作B使用相同的令牌
        result2 = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!K1:T10",  # 不同的范围
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token
        )

        assert result2['success'] is False
        assert result2['can_proceed'] is False
        assert "不匹配" in result2['error']

    def test_confirmation_with_different_risk_levels(self, temp_excel_file):
        """测试不同风险级别的确认流程"""
        test_cases = [
            ('low', 'TestData!A10:C10', [["A", "B", "C"]], False),
            ('medium', 'TestData!A5:E8', [["A" for _ in range(5)] for _ in range(4)], False),
            ('high', 'TestData!A1:J10', [["A" for _ in range(10)] for _ in range(10)], True),
            ('critical', 'TestData!A1:T50', [["A" for _ in range(20)] for _ in range(50)], True)
        ]

        for risk_level, range_expr, data, requires_confirmation in test_cases:
            impact = {
                'success': True,
                'impact_analysis': {
                    'operation_risk_level': risk_level,
                    'total_cells': len(data) * len(data[0]) if data else 0,
                    'existing_data_count': 10 if risk_level in ['high', 'critical'] else 0,
                    'formula_count': 5 if risk_level == 'critical' else 0
                }
            }

            result = ExcelOperations.confirm_operation(
                file_path=temp_excel_file,
                range_expression=range_expr,
                operation_type="update",
                impact_analysis=impact,
                preview_data=data
            )

            assert result['success'] is True
            assert result['confirmation_required'] == requires_confirmation
            assert result['risk_level'] == risk_level

    def test_confirmation_safety_steps_content(self, temp_excel_file):
        """测试确认安全步骤内容"""
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'critical',
                'total_cells': 1000,
                'existing_data_count': 800,
                'formula_count': 100
            },
            'warnings': [
                {'level': 'error', 'message': '极高风险警告'},
                {'level': 'warning', 'message': '数据覆盖警告'}
            ]
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A1:Z50",
            operation_type="update",
            impact_analysis=impact
        )

        assert result['success'] is True
        assert 'safety_steps' in result
        safety_steps = result['safety_steps']

        # 验证安全步骤内容
        step_types = [step['type'] for step in safety_steps]
        assert 'manual_backup' in step_types
        assert 'data_review' in step_types
        assert 'final_confirmation' in step_types

        # 验证每个步骤都有必要的字段
        for step in safety_steps:
            assert 'type' in step
            assert 'description' in step
            assert 'required' in step
            assert 'completed' in step

    def test_confirmation_with_preview_data(self, temp_excel_file):
        """测试带预览数据的确认"""
        new_data = [
            ["新ID1", "新名称1", "新类型1", "新数值1"],
            ["新ID2", "新名称2", "新类型2", "新数值2"]
        ]

        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'medium',
                'total_cells': 8,
                'existing_data_count': 4
            },
            'current_data_preview': [
                ["旧ID1", "旧名称1", "旧类型1", "旧数值1"],
                ["旧ID2", "旧名称2", "旧类型2", "旧数值2"]
            ]
        }

        result = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="TestData!A5:D6",
            operation_type="update",
            impact_analysis=impact,
            preview_data=new_data
        )

        assert result['success'] is True
        assert 'preview_comparison' in result
        comparison = result['preview_comparison']

        assert 'current_data' in comparison
        assert 'new_data' in comparison
        assert comparison['current_data'] == impact['current_data_preview']
        assert comparison['new_data'] == new_data

    def test_confirmation_error_handling(self):
        """测试确认错误处理"""
        # 测试无效的影响分析
        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C1",
            operation_type="update",
            impact_analysis=None
        )

        assert result['success'] is False
        assert "影响分析" in result['error']

        # 测试无效的操作类型
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        result = ExcelOperations.confirm_operation(
            file_path="test.xlsx",
            range_expression="Sheet1!A1:C1",
            operation_type="invalid_operation",
            impact_analysis=impact
        )

        assert result['success'] is False
        assert "无效的操作类型" in result['error']


class TestConfirmationIntegration:
    """确认集成测试类"""

    @pytest.fixture
    def temp_excel_file(self):
        """创建临时Excel文件用于测试"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file_path = tmp.name

        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "GameData"

        ws.append(["技能ID", "技能名称", "技能类型", "技能伤害", "冷却时间"])
        ws.append([1001, "火球术", "火系", 150, 3.0])
        ws.append([1002, "冰箭", "冰系", 120, 2.5])

        wb.save(file_path)
        wb.close()

        yield file_path

        # 清理
        if os.path.exists(file_path):
            os.unlink(file_path)

    def test_full_confirmation_workflow_high_risk(self, temp_excel_file):
        """测试高风险操作的完整确认工作流程"""
        new_skills = [
            [1003, "雷击", "雷系", 180, 4.0],
            [1004, "风刃", "风系", 140, 2.8],
            [1005, "土盾", "土系", 80, 5.0]
        ]

        # 1. 评估操作影响
        impact = ExcelOperations.assess_operation_impact(
            file_path=temp_excel_file,
            range_expression="GameData!A3:E5",
            operation_type="update",
            preview_data=new_skills
        )

        assert impact['success'] is True

        # 2. 请求确认
        confirmation = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A3:E5",
            operation_type="update",
            impact_analysis=impact,
            preview_data=new_skills
        )

        assert confirmation['success'] is True
        assert confirmation['confirmation_required'] is True

        # 3. 使用令牌确认
        token = confirmation['confirmation_token']
        final_confirmation = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A3:E5",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token
        )

        assert final_confirmation['success'] is True
        assert final_confirmation['can_proceed'] is True
        assert final_confirmation['confirmed'] is True

    def test_confirmation_cancellation(self, temp_excel_file):
        """测试确认取消流程"""
        # 1. 创建高风险操作
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high',
                'total_cells': 100,
                'existing_data_count': 50
            }
        }

        # 2. 请求确认
        confirmation = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A1:J10",
            operation_type="update",
            impact_analysis=impact
        )

        assert confirmation['success'] is True
        assert confirmation['confirmation_required'] is True

        # 3. 不提供令牌，直接尝试执行（应该失败）
        result = ExcelOperations.update_range(
            file_path=temp_excel_file,
            range="GameData!A1:J10",
            data=[["test"]],
            insert_mode=True,
            require_confirmation=True,
            skip_safety_checks=False
        )

        assert result['success'] is False
        assert "需要用户确认" in result['error']

    def test_confirmation_persistence_across_operations(self, temp_excel_file):
        """测试确认在操作间的持久性"""
        # 1. 为操作A生成确认
        impact = {
            'success': True,
            'impact_analysis': {
                'operation_risk_level': 'high'
            }
        }

        confirmation_a = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A1:E1",
            operation_type="update",
            impact_analysis=impact
        )

        token_a = confirmation_a['confirmation_token']

        # 2. 为操作B生成确认
        confirmation_b = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A2:E2",
            operation_type="update",
            impact_analysis=impact
        )

        token_b = confirmation_b['confirmation_token']

        # 3. 验证两个令牌不同
        assert token_a != token_b

        # 4. 验证令牌只能用于对应的操作
        result_a = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A1:E1",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token_a
        )

        result_b_wrong = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression="GameData!A1:E1",
            operation_type="update",
            impact_analysis=impact,
            confirmation_token=token_b  # 使用错误的令牌
        )

        assert result_a['success'] is True
        assert result_a['confirmed'] is True
        assert result_b_wrong['success'] is False

    def test_batch_operation_confirmation(self, temp_excel_file):
        """测试批量操作确认"""
        batch_operations = [
            {
                'range': 'GameData!A3:E3',
                'data': [[1003, "雷击", "雷系", 180, 4.0]],
                'impact': {
                    'success': True,
                    'impact_analysis': {'operation_risk_level': 'medium'}
                }
            },
            {
                'range': 'GameData!A4:E4',
                'data': [[1004, "风刃", "风系", 140, 2.8]],
                'impact': {
                    'success': True,
                    'impact_analysis': {'operation_risk_level': 'medium'}
                }
            },
            {
                'range': 'GameData!A5:E5',
                'data': [[1005, "土盾", "土系", 80, 5.0]],
                'impact': {
                    'success': True,
                    'impact_analysis': {'operation_risk_level': 'high'}
                }
            }
        ]

        # 批量确认应该为高风险操作请求确认
        confirmations = []
        for i, op in enumerate(batch_operations):
            confirmation = ExcelOperations.confirm_operation(
                file_path=temp_excel_file,
                range_expression=op['range'],
                operation_type="update",
                impact_analysis=op['impact'],
                preview_data=op['data']
            )
            confirmations.append(confirmation)

        # 验证确认结果
        assert len(confirmations) == 3
        assert confirmations[0]['confirmation_required'] is False  # 中等风险
        assert confirmations[1]['confirmation_required'] is False  # 中等风险
        assert confirmations[2]['confirmation_required'] is True   # 高风险

        # 只有高风险操作需要令牌确认
        high_risk_confirmation = confirmations[2]
        token = high_risk_confirmation['confirmation_token']

        final_confirmation = ExcelOperations.confirm_operation(
            file_path=temp_excel_file,
            range_expression=batch_operations[2]['range'],
            operation_type="update",
            impact_analysis=batch_operations[2]['impact'],
            confirmation_token=token
        )

        assert final_confirmation['success'] is True
        assert final_confirmation['confirmed'] is True


if __name__ == "__main__":
    # 运行测试
    pytest.main([__file__, "-v"])