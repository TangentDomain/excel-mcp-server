"""测试错误分类统计功能 — REQ-013

验证ToolCallTracker能自动检测返回值中的success=False并分类错误类型。
"""
import pytest
import time
from collections import defaultdict
from unittest.mock import patch, MagicMock


class TestErrorClassification:
    """错误分类逻辑测试"""

    def setup_method(self):
        """每个测试重置tracker"""
        # 导入最新的tracker
        from excel_mcp_server_fastmcp.server import _tracker
        _tracker.reset()

    def _get_tracker(self):
        from excel_mcp_server_fastmcp.server import _tracker
        return _tracker

    def test_classify_security_error(self):
        """🔒前缀的错误分类为security"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('🔒 安全验证失败: 路径穿越') == 'security'

    def test_classify_file_not_found(self):
        """文件不存在错误分类为file_not_found"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('文件不存在: /tmp/test.xlsx') == 'file_not_found'

    def test_classify_file_not_found_from_query_info(self):
        """SQL引擎error_type=file_not_found正确传递"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('error_type: file_not_found') == 'file_not_found'

    def test_classify_validation_error(self):
        """无效参数错误分类为validation"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('无效的列名: xxx') == 'validation'
        assert ToolCallTracker.classify_error('文件路径不能为空') == 'validation'

    def test_classify_sheet_not_found(self):
        """工作表不存在错误分类为sheet_not_found"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('工作表不存在: NotFound') == 'sheet_not_found'

    def test_classify_unsupported(self):
        """不支持操作分类为unsupported"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('不支持的功能') == 'unsupported'

    def test_classify_column_error(self):
        """列名相关错误分类为column"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('列名 skil_name 不存在') == 'column'

    def test_classify_sql_syntax(self):
        """SQL语法错误分类为sql_syntax"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('SQL语法错误: SELECT FROM') == 'sql_syntax'
        assert ToolCallTracker.classify_error('syntax_error: unexpected token') == 'sql_syntax'

    def test_classify_unknown(self):
        """无法识别的错误分类为unknown"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('') == 'unknown'
        assert ToolCallTracker.classify_error('某个未知错误') == 'unknown'

    def test_classify_file_load(self):
        """文件加载失败分类为file_load"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('无法加载文件: corrupted') == 'file_load'
        assert ToolCallTracker.classify_error('data_load_failed') == 'file_load'

    def test_classify_file_too_large(self):
        """文件过大分类为file_too_large"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('文件太大: 100MB') == 'file_too_large'

    def test_classify_execution_error(self):
        """执行错误分类为execution"""
        from excel_mcp_server_fastmcp.server import ToolCallTracker
        assert ToolCallTracker.classify_error('execution_error: timeout') == 'execution'


class TestTrackerErrorRecording:
    """Tracker错误记录测试"""

    def setup_method(self):
        from excel_mcp_server_fastmcp.server import _tracker
        _tracker.reset()

    def _get_tracker(self):
        from excel_mcp_server_fastmcp.server import _tracker
        return _tracker

    def test_record_with_error_type(self):
        """record方法支持error_type参数"""
        tracker = self._get_tracker()
        tracker.record('test_tool', 10.0, success=False, error_type='file_not_found')
        stats = tracker.get_stats()
        assert stats['total_errors'] == 1
        assert stats['error_types'] == {'file_not_found': 1}
        assert stats['tools']['test_tool']['error_types'] == {'file_not_found': 1}

    def test_record_without_error_type(self):
        """record方法不传error_type时默认为unknown"""
        tracker = self._get_tracker()
        tracker.record('test_tool', 10.0, success=False)
        stats = tracker.get_stats()
        assert stats['error_types'] == {'unknown': 1}

    def test_record_success_no_error(self):
        """成功调用不影响错误统计"""
        tracker = self._get_tracker()
        tracker.record('test_tool', 10.0, success=True)
        stats = tracker.get_stats()
        assert stats['total_errors'] == 0
        assert stats['error_types'] == {}

    def test_multiple_error_types(self):
        """多次不同错误类型正确累计"""
        tracker = self._get_tracker()
        tracker.record('tool_a', 10.0, success=False, error_type='file_not_found')
        tracker.record('tool_a', 5.0, success=False, error_type='file_not_found')
        tracker.record('tool_b', 3.0, success=False, error_type='security')
        stats = tracker.get_stats()
        assert stats['total_errors'] == 3
        assert stats['error_types'] == {'file_not_found': 2, 'security': 1}

    def test_global_error_types_aggregation(self):
        """全局error_types正确聚合所有工具的错误"""
        tracker = self._get_tracker()
        tracker.record('tool_a', 10.0, success=False, error_type='file_not_found')
        tracker.record('tool_b', 5.0, success=False, error_type='file_not_found')
        tracker.record('tool_b', 3.0, success=False, error_type='validation')
        stats = tracker.get_stats()
        assert stats['error_types'] == {'file_not_found': 2, 'validation': 1}

    def test_error_types_sorted(self):
        """error_types按类型名排序"""
        tracker = self._get_tracker()
        tracker.record('tool_a', 10.0, success=False, error_type='security')
        tracker.record('tool_a', 5.0, success=False, error_type='file_not_found')
        stats = tracker.get_stats()
        keys = list(stats['error_types'].keys())
        assert keys == ['file_not_found', 'security']

    def test_per_tool_error_types(self):
        """每个工具独立记录错误类型"""
        tracker = self._get_tracker()
        tracker.record('tool_a', 10.0, success=False, error_type='file_not_found')
        tracker.record('tool_b', 5.0, success=False, error_type='security')
        stats = tracker.get_stats()
        assert stats['tools']['tool_a']['error_types'] == {'file_not_found': 1}
        assert stats['tools']['tool_b']['error_types'] == {'security': 1}

    def test_reset_clears_error_types(self):
        """reset清空所有统计包括错误类型"""
        tracker = self._get_tracker()
        tracker.record('tool_a', 10.0, success=False, error_type='file_not_found')
        tracker.reset()
        stats = tracker.get_stats()
        assert stats['total_errors'] == 0
        assert stats['error_types'] == {}


class TestTrackCallDecorator:
    """_track_call装饰器错误检测测试"""

    def setup_method(self):
        from excel_mcp_server_fastmcp.server import _tracker
        _tracker.reset()

    def _get_tracker(self):
        from excel_mcp_server_fastmcp.server import _tracker
        return _tracker

    def test_success_false_detected_as_error(self):
        """装饰器检测返回值中的success=False"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool():
            return {'success': False, 'message': '文件不存在: test.xlsx'}

        result = fake_tool()
        assert result['success'] is False
        stats = self._get_tracker().get_stats()
        assert stats['total_errors'] == 1
        assert stats['error_types'] == {'file_not_found': 1}

    def test_success_true_not_counted_as_error(self):
        """装饰器不将success=True计为错误"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool():
            return {'success': True, 'data': 'ok'}

        fake_tool()
        stats = self._get_tracker().get_stats()
        assert stats['total_errors'] == 0
        assert stats['error_types'] == {}

    def test_non_dict_return_not_inspected(self):
        """非dict返回值不做错误检测"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool():
            return "not a dict"

        fake_tool()
        stats = self._get_tracker().get_stats()
        assert stats['total_errors'] == 0

    def test_exception_still_tracked(self):
        """异常仍然被追踪和分类"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool():
            raise ValueError("文件不存在: test.xlsx")

        with pytest.raises(ValueError):
            fake_tool()
        stats = self._get_tracker().get_stats()
        assert stats['total_errors'] == 1
        assert 'file_not_found' in stats['error_types']

    def test_sql_error_type_from_query_info(self):
        """SQL引擎的error_type优先于消息内容分类"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_sql_tool():
            return {
                'success': False,
                'message': 'SQL语法错误',
                'query_info': {'error_type': 'unsupported_sql', 'details': 'UNION'}
            }

        result = fake_sql_tool()
        stats = self._get_tracker().get_stats()
        assert stats['error_types'] == {'unsupported_sql': 1}

    def test_security_error_classification(self):
        """安全错误正确分类"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool():
            return {'success': False, 'message': '🔒 安全验证失败: 路径穿越'}

        fake_tool()
        stats = self._get_tracker().get_stats()
        assert stats['error_types'] == {'security': 1}

    def test_mixed_calls_tracked_correctly(self):
        """混合成功/失败调用正确统计"""
        from excel_mcp_server_fastmcp.server import _track_call

        @_track_call
        def fake_tool(success=True, msg=''):
            if success:
                return {'success': True, 'data': 'ok'}
            return {'success': False, 'message': msg}

        fake_tool(success=True)
        fake_tool(success=True)
        fake_tool(success=False, msg='文件不存在: a.xlsx')
        fake_tool(success=False, msg='🔒 安全验证失败')
        fake_tool(success=True)

        stats = self._get_tracker().get_stats()
        assert stats['total_calls'] == 5
        assert stats['total_errors'] == 2
        assert stats['error_types'] == {'file_not_found': 1, 'security': 1}
