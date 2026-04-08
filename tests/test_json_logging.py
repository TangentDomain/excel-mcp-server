"""结构化JSON日志测试 - REQ-013"""

import json
import logging
import os
import io
import pytest
from datetime import datetime

# 导入server模块的JsonFormatter
from src.excel_mcp_server_fastmcp.server import JsonFormatter


class TestJsonFormatter:
    """测试JSON日志格式化器"""

    def _make_record(self, msg, **kwargs):
        """创建一个LogRecord"""
        logger = logging.getLogger('test')
        record = logger.makeRecord(
            name='test', level=logging.DEBUG, fn='test.py', lno=1,
            msg=msg, args=None, exc_info=None
        )
        for k, v in kwargs.items():
            setattr(record, k, v)
        return record

    def test_basic_fields(self):
        """基本字段: ts, level, module, message"""
        formatter = JsonFormatter()
        record = self._make_record("hello")
        output = formatter.format(record)
        data = json.loads(output)
        assert 'ts' in data
        assert data['level'] == 'DEBUG'
        assert data['module'] == 'test'
        assert data['message'] == 'hello'

    def test_extra_tool_field(self):
        """extra字段: tool"""
        formatter = JsonFormatter()
        record = self._make_record("test", tool='excel_query', duration_ms=12.3)
        output = formatter.format(record)
        data = json.loads(output)
        assert data['tool'] == 'excel_query'
        assert data['duration_ms'] == 12.3

    def test_extra_error_field(self):
        """extra字段: error"""
        formatter = JsonFormatter()
        record = self._make_record("failed", tool='excel_query', error='File not found')
        output = formatter.format(record)
        data = json.loads(output)
        assert data['error'] == 'File not found'

    def test_no_extra_fields(self):
        """无extra字段时不包含可选字段"""
        formatter = JsonFormatter()
        record = self._make_record("plain message")
        output = formatter.format(record)
        data = json.loads(output)
        assert 'tool' not in data
        assert 'duration_ms' not in data
        assert 'error' not in data

    def test_chinese_message(self):
        """中文消息不转义"""
        formatter = JsonFormatter()
        record = self._make_record("文件不存在: 配置表.xlsx")
        output = formatter.format(record)
        data = json.loads(output)
        assert data['message'] == "文件不存在: 配置表.xlsx"

    def test_output_is_single_line(self):
        """输出为单行JSON（不含换行符）"""
        formatter = JsonFormatter()
        record = self._make_record("test message")
        output = formatter.format(record)
        assert '\n' not in output

    def test_output_is_valid_json(self):
        """输出是合法JSON"""
        formatter = JsonFormatter()
        record = self._make_record("test", tool='excel_query', duration_ms=1.5, error='timeout')
        output = formatter.format(record)
        data = json.loads(output)
        assert isinstance(data, dict)

    def test_file_path_and_operation_fields(self):
        """file_path和operation可选字段"""
        formatter = JsonFormatter()
        record = self._make_record("operation", file_path='/tmp/test.xlsx', operation='update')
        output = formatter.format(record)
        data = json.loads(output)
        assert data['file_path'] == '/tmp/test.xlsx'
        assert data['operation'] == 'update'

    def test_level_names(self):
        """不同日志级别正确输出"""
        formatter = JsonFormatter()
        for level, name in [(logging.DEBUG, 'DEBUG'), (logging.INFO, 'INFO'),
                            (logging.WARNING, 'WARNING'), (logging.ERROR, 'ERROR')]:
            record = self._make_record("msg")
            record.levelno = level
            record.levelname = name
            data = json.loads(formatter.format(record))
            assert data['level'] == name
