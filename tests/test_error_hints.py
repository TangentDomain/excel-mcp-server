"""
集中式错误提示系统测试
验证所有非SQL错误都包含💡修复提示
"""

import pytest
from src.excel_mcp.server import _fail, _wrap, _ERROR_HINTS, _infer_error_code


class TestCentralizedErrorHints:
    """测试集中式错误提示映射"""

    def test_all_error_codes_have_hints(self):
        """所有已使用的error_code都应该有对应的提示"""
        used_codes = {
            'PATH_VALIDATION_FAILED', 'FILE_SIZE_EXCEEDED', 'VALIDATION_FAILED',
            'OPERATION_FAILED', 'PREVIEW_FAILED', 'ASSESSMENT_FAILED',
            'HISTORY_RETRIEVAL_FAILED', 'FILE_NOT_FOUND', 'BACKUP_FAILED',
            'BACKUP_NOT_FOUND', 'RESTORE_FAILED', 'LIST_BACKUPS_FAILED',
            'DELETE_SHEET_FAILED', 'DELETE_ROWS_FAILED', 'DELETE_COLUMNS_FAILED',
            'FORMULA_SECURITY_FAILED', 'MISSING_FILE_PATH', 'MISSING_QUERY',
            'INVALID_FORMAT', 'DEPENDENCY_MISSING', 'SQL_EXECUTION_FAILED',
            'UNSUPPORTED_SQL', 'UPDATE_EXECUTION_FAILED', 'FILE_OPEN_FAILED',
            'SHEET_NOT_FOUND', 'EMPTY_SHEET', 'DESCRIBE_FAILED',
        }
        for code in used_codes:
            assert code in _ERROR_HINTS, f"error_code '{code}' 缺少修复提示"

    def test_fail_without_meta_no_hint(self):
        """没有meta时不崩溃"""
        result = _fail('some error')
        assert result['success'] is False
        assert 'some error' in result['message']

    def test_fail_auto_appends_hint(self):
        """_fail自动附加error_code对应的提示"""
        result = _fail('文件不存在', meta={"error_code": "FILE_NOT_FOUND"})
        assert '💡' in result['message']
        assert 'excel_list_sheets' in result['message']

    def test_fail_no_duplicate_hint(self):
        """message中已有💡提示时不重复添加"""
        result = _fail('错误\n💡 已有提示', meta={"error_code": "FILE_NOT_FOUND"})
        # 只应出现一次💡
        assert result['message'].count('💡') == 1

    def test_fail_unknown_code_no_crash(self):
        """未知的error_code不崩溃，不附加提示"""
        result = _fail('未知错误', meta={"error_code": "UNKNOWN_CODE_XYZ"})
        assert result['success'] is False
        assert '💡' not in result['message']

    def test_hint_content_useful(self):
        """提示内容应包含具体的修复建议"""
        assert 'excel_list_sheets' in _ERROR_HINTS['FILE_NOT_FOUND']
        assert 'excel_list_sheets' in _ERROR_HINTS['SHEET_NOT_FOUND']
        assert 'SELECT' in _ERROR_HINTS['MISSING_QUERY']
        assert '范围' in _ERROR_HINTS['VALIDATION_FAILED']

    def test_fail_preserves_meta(self):
        """_fail保留所有meta字段"""
        result = _fail('错误', meta={"error_code": "X", "extra": "data"})
        assert result['meta']['error_code'] == 'X'
        assert result['meta']['extra'] == 'data'


class TestInferErrorCode:
    """测试消息内容推断error_code"""

    def test_infer_file_not_found(self):
        assert _infer_error_code('Excel文件不存在: test.xlsx') == 'FILE_NOT_FOUND'

    def test_infer_sheet_not_found(self):
        assert _infer_error_code('工作表 不存在的工作表 不存在') == 'SHEET_NOT_FOUND'

    def test_infer_file_open_failed(self):
        assert _infer_error_code('无法打开文件: permission denied') == 'FILE_OPEN_FAILED'

    def test_infer_empty_sheet(self):
        assert _infer_error_code('工作表为空') == 'EMPTY_SHEET'

    def test_infer_path_validation(self):
        assert _infer_error_code('路径不允许包含".."路径穿越') == 'PATH_VALIDATION_FAILED'

    def test_infer_unknown(self):
        assert _infer_error_code('some random error') == ''


class TestWrapHints:
    """测试_wrap也附加错误提示"""

    def test_wrap_operations_error_gets_hint(self):
        """Operations层错误（无error_code）也应获得提示"""
        result = _wrap({
            'success': False,
            'message': '获取工作表列表失败: Excel文件不存在: test.xlsx',
            'data': None
        })
        assert '💡' in result['message']
        assert result['meta']['error_code'] == 'FILE_NOT_FOUND'

    def test_wrap_success_no_hint(self):
        """成功结果不附加提示"""
        result = _wrap({'success': True, 'message': 'ok', 'data': []})
        assert '💡' not in result['message']

    def test_wrap_existing_error_code_no_override(self):
        """已有error_code时不覆盖"""
        result = _wrap({
            'success': False,
            'message': '自定义错误',
            'data': None,
            'meta': {'error_code': 'CUSTOM_CODE'}
        })
        assert result['meta']['error_code'] == 'CUSTOM_CODE'

    def test_wrap_existing_hint_no_duplicate(self):
        """已有💡时不重复"""
        result = _wrap({
            'success': False,
            'message': '错误\n💡 已有提示',
            'data': None
        })
        assert result['message'].count('💡') == 1

    def test_wrap_sql_query_info_preserved(self):
        """SQL查询的query_info不被破坏"""
        result = _wrap({
            'success': False,
            'message': 'SQL错误',
            'data': [],
            'query_info': {'error_type': 'sql_syntax', 'hint': 'fix syntax'}
        })
        assert result['query_info']['error_type'] == 'sql_syntax'
        # SQL错误由SQL引擎自己处理提示，_wrap不覆盖
        # 但Operations层错误会被附加提示
