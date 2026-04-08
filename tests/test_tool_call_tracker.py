"""ToolCallTracker 工具调用追踪器测试"""

import sys
import os
import time

# 确保能导入包
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from excel_mcp.server import ToolCallTracker, _track_call, _tracker


class TestToolCallTracker:
    """ToolCallTracker 单元测试"""

    def setup_method(self):
        """每个测试前重置追踪器"""
        _tracker.reset()

    def test_record_single_call(self):
        """单次调用记录"""
        _tracker.record('test_tool', 10.5, success=True)
        stats = _tracker.get_stats()
        assert stats['total_calls'] == 1
        assert stats['total_errors'] == 0
        assert 'test_tool' in stats['tools']
        assert stats['tools']['test_tool']['call_count'] == 1
        assert stats['tools']['test_tool']['avg_time_ms'] == 10.5
        assert stats['tools']['test_tool']['error_count'] == 0

    def test_record_multiple_calls(self):
        """多次调用统计"""
        _tracker.record('tool_a', 10.0, success=True)
        _tracker.record('tool_a', 20.0, success=True)
        _tracker.record('tool_a', 30.0, success=False)

        stats = _tracker.get_stats()
        t = stats['tools']['tool_a']
        assert t['call_count'] == 3
        assert t['avg_time_ms'] == 20.0  # (10+20+30)/3
        assert t['min_time_ms'] == 10.0
        assert t['max_time_ms'] == 30.0
        assert t['error_count'] == 1

    def test_record_error(self):
        """错误调用记录"""
        _tracker.record('fail_tool', 5.0, success=False)
        stats = _tracker.get_stats()
        assert stats['total_errors'] == 1
        assert stats['tools']['fail_tool']['error_count'] == 1

    def test_get_stats_sorted_by_call_count(self):
        """统计结果按调用次数降序"""
        for _ in range(5):
            _tracker.record('popular', 1.0, success=True)
        for _ in range(2):
            _tracker.record('normal', 1.0, success=True)
        for _ in range(1):
            _tracker.record('rare', 1.0, success=True)

        tools = list(_tracker.get_stats()['tools'].keys())
        assert tools == ['popular', 'normal', 'rare']

    def test_reset(self):
        """重置统计"""
        _tracker.record('tool', 10.0, success=True)
        _tracker.reset()
        stats = _tracker.get_stats()
        assert stats['total_calls'] == 0
        assert stats['total_errors'] == 0
        assert len(stats['tools']) == 0

    def test_uptime(self):
        """运行时间"""
        _tracker.record('tool', 1.0, success=True)
        stats = _tracker.get_stats()
        assert stats['uptime_seconds'] >= 0

    def test_last_called_timestamp(self):
        """最后调用时间"""
        _tracker.record('tool', 1.0, success=True)
        stats = _tracker.get_stats()
        assert stats['tools']['tool']['last_called'] is not None

    def test_min_max_time(self):
        """最小/最大耗时"""
        _tracker.record('tool', 100.0, success=True)
        _tracker.record('tool', 1.0, success=True)
        _tracker.record('tool', 50.0, success=True)

        t = stats = _tracker.get_stats()['tools']['tool']
        assert t['min_time_ms'] == 1.0
        assert t['max_time_ms'] == 100.0


class TestTrackCallDecorator:
    """_track_call 装饰器测试"""

    def setup_method(self):
        _tracker.reset()

    def test_tracks_success(self):
        """成功调用被追踪"""
        @_track_call
        def success_func():
            return {'result': 'ok'}

        result = success_func()
        assert result == {'result': 'ok'}
        stats = _tracker.get_stats()
        assert stats['total_calls'] == 1
        assert stats['total_errors'] == 0

    def test_tracks_error(self):
        """失败调用被追踪"""
        @_track_call
        def fail_func():
            raise ValueError("test error")

        try:
            fail_func()
            assert False, "Should have raised"
        except ValueError:
            pass

        stats = _tracker.get_stats()
        assert stats['total_calls'] == 1
        assert stats['total_errors'] == 1

    def test_preserves_function_name(self):
        """函数名保持不变"""
        @_track_call
        def my_function():
            pass

        assert my_function.__name__ == 'my_function'

    def test_preserves_arguments(self):
        """参数透传正确"""
        @_track_call
        def add(a, b, c=0):
            return a + b + c

        assert add(1, 2) == 3
        assert add(1, 2, c=3) == 6

    def test_duration_recorded(self):
        """耗时被记录"""
        @_track_call
        def slow_func():
            time.sleep(0.05)
            return 'done'

        slow_func()
        stats = _tracker.get_stats()
        t = stats['tools']['slow_func']
        assert t['avg_time_ms'] >= 40  # 至少40ms
        assert t['call_count'] == 1

    def test_multiple_tools_tracked(self):
        """多个工具分别追踪"""
        @_track_call
        def tool_a():
            return 'a'

        @_track_call
        def tool_b():
            return 'b'

        tool_a()
        tool_a()
        tool_b()

        stats = _tracker.get_stats()
        assert stats['total_calls'] == 3
        assert stats['tools']['tool_a']['call_count'] == 2
        assert stats['tools']['tool_b']['call_count'] == 1
