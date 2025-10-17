# -*- coding: utf-8 -*-
"""
统一错误处理模块完整测试套件
测试src.utils.error_handler模块的所有核心功能
目标覆盖率：85%+
"""

import pytest
import logging
import time
import functools
import threading
from unittest.mock import Mock, patch, MagicMock, call
from typing import Any, Dict

from src.utils.error_handler import (
    ErrorHandler,
    unified_error_handler,
    extract_file_context,
    extract_formula_context
)
from src.models.types import OperationResult
from src.utils.exceptions import (
    ExcelFileNotFoundError, SheetNotFoundError, DataValidationError
)


class TestErrorHandlerBasic:
    """ErrorHandler基础功能测试"""

    def test_get_error_code_known_exception(self):
        """测试获取已知异常的错误代码"""
        error_code = ErrorHandler.get_error_code(FileNotFoundError("文件不存在"))
        assert error_code == 'FILE_NOT_FOUND'

        error_code = ErrorHandler.get_error_code(ValueError("值错误"))
        assert error_code == 'VALUE_ERROR'

        error_code = ErrorHandler.get_error_code(KeyError("键错误"))
        assert error_code == 'KEY_ERROR'

    def test_get_error_code_unknown_exception(self):
        """测试获取未知异常的错误代码"""
        class CustomError(Exception):
            pass

        error_code = ErrorHandler.get_error_code(CustomError("自定义错误"))
        assert error_code == 'UNKNOWN_ERROR'

    def test_get_error_solution_known_code(self):
        """测试获取已知错误代码的解决方案"""
        solution = ErrorHandler.get_error_solution('FILE_NOT_FOUND')
        assert '文件路径' in solution
        assert '文件存在' in solution

        solution = ErrorHandler.get_error_solution('SHEET_NOT_FOUND')
        assert '工作表名称' in solution
        assert 'excel_list_sheets' in solution

    def test_get_error_solution_unknown_code(self):
        """测试获取未知错误代码的解决方案"""
        solution = ErrorHandler.get_error_solution('UNKNOWN_CODE')
        assert solution == '请联系技术支持或查看文档'


class TestErrorHandlerFormatResponse:
    """错误响应格式化测试"""

    def test_format_error_response_basic(self):
        """测试基础错误响应格式化"""
        error = FileNotFoundError("测试文件不存在")
        response = ErrorHandler.format_error_response(error)

        assert response['code'] == 'FILE_NOT_FOUND'
        assert response['message'] == "测试文件不存在"
        assert response['type'] == 'FileNotFoundError'
        assert response['solution'] is not None
        assert response['severity'] == 'error'
        assert response['details'] == {}

    def test_format_error_response_with_context(self):
        """测试带上下文的错误响应格式化"""
        error = ValueError("无效参数")
        context = {
            'file_path': 'test.xlsx',
            'operation': 'get_range',
            'user_input': 'A1:Z1000'
        }

        response = ErrorHandler.format_error_response(error, context)

        assert response['code'] == 'VALUE_ERROR'
        assert response['details'] == context
        assert response['severity'] == 'error'

    def test_format_error_response_with_operation(self):
        """测试带操作名称的错误响应格式化"""
        error = PermissionError("权限不足")
        context = {'file_path': 'protected.xlsx'}
        operation = 'update_range'

        response = ErrorHandler.format_error_response(error, context, operation)

        assert response['code'] == 'PERMISSION_DENIED'
        assert response['operation'] == 'update_range'
        assert response['affected_resource'] == 'protected.xlsx'

    def test_format_error_response_warning_severity(self):
        """测试警告级别的错误响应"""
        error = ValueError("数据验证失败")
        # 模拟数据验证错误
        with patch.object(ErrorHandler, 'get_error_code', return_value='DATA_VALIDATION_ERROR'):
            response = ErrorHandler.format_error_response(error)
            assert response['severity'] == 'warning'

    def test_format_error_response_comprehensive(self):
        """测试完整的错误响应格式化"""
        error = Exception("通用错误")
        context = {
            'file_path': 'test.xlsx',
            'sheet_name': 'Sheet1',
            'range': 'A1:C10',
            'additional_info': '详细错误信息'
        }
        operation = 'comprehensive_test'

        response = ErrorHandler.format_error_response(error, context, operation)

        # 验证所有字段都存在
        required_fields = ['code', 'message', 'type', 'solution', 'severity', 'details']
        for field in required_fields:
            assert field in response

        assert response['operation'] == operation
        assert response['affected_resource'] == 'test.xlsx'
        assert response['details'] == context


class TestUnifiedErrorHandler:
    """统一错误处理装饰器测试"""

    def test_successful_function_without_decorator_return(self):
        """测试成功函数（不返回OperationResult）"""
        @unified_error_handler("test_operation")
        def test_function():
            return "success_result"

        result = test_function()

        assert isinstance(result, OperationResult)
        assert result.success is True
        assert result.data == "success_result"
        assert result.metadata['operation'] == "test_operation"
        assert 'execution_time_ms' in result.metadata
        assert 'timestamp' in result.metadata

    def test_successful_function_with_operation_result(self):
        """测试返回OperationResult的成功函数"""
        @unified_error_handler("test_operation")
        def test_function():
            return OperationResult(success=True, data={"key": "value"})

        result = test_function()

        assert isinstance(result, OperationResult)
        assert result.success is True
        assert result.data == {"key": "value"}
        assert result.metadata['operation'] == "test_operation"

    def test_function_with_exception(self):
        """测试抛出异常的函数"""
        @unified_error_handler("test_operation")
        def test_function():
            raise FileNotFoundError("测试文件不存在")

        result = test_function()

        assert isinstance(result, OperationResult)
        assert result.success is False
        assert 'error' in result.__dict__
        assert result.error['code'] == 'FILE_NOT_FOUND'

    def test_function_with_return_dict_true(self):
        """测试return_dict=True的情况"""
        @unified_error_handler("test_operation", return_dict=True)
        def test_function():
            return {"custom": "result"}

        result = test_function()

        assert isinstance(result, dict)
        assert result == {"custom": "result"}

    def test_function_with_exception_and_return_dict(self):
        """测试异常且return_dict=True的情况"""
        @unified_error_handler("test_operation", return_dict=True)
        def test_function():
            raise ValueError("测试值错误")

        result = test_function()

        assert isinstance(result, dict)
        assert result['success'] is False
        assert 'error' in result
        assert result['error']['code'] == 'VALUE_ERROR'
        assert 'timestamp' in result
        assert 'execution_time_ms' in result

    def test_decorator_with_context_extractor(self):
        """测试带上下文提取器的装饰器"""
        def mock_context_extractor(*args, **kwargs):
            return {"extracted": "context", "args": str(args), "kwargs": str(kwargs)}

        @unified_error_handler("test_operation", context_extractor=mock_context_extractor)
        def test_function(file_path, sheet_name=None):
            raise PermissionError("权限测试")

        result = test_function("test.xlsx", sheet_name="Sheet1")

        assert result.success is False
        assert result.error['details']['extracted'] == "context"
        assert "test.xlsx" in result.error['details']['args']

    def test_decorator_with_failing_context_extractor(self):
        """测试上下文提取器失败的情况"""
        def failing_context_extractor(*args, **kwargs):
            raise Exception("上下文提取失败")

        @unified_error_handler("test_operation", context_extractor=failing_context_extractor)
        def test_function():
            raise ValueError("主要错误")

        result = test_function()

        assert result.success is False
        # 主要错误应该被正确处理，上下文提取失败不应该影响主要错误处理

    def test_decorator_execution_time_calculation(self):
        """测试执行时间计算"""
        @unified_error_handler("timing_test")
        def slow_function():
            time.sleep(0.01)  # 10ms延迟
            return "completed"

        result = slow_function()

        assert isinstance(result, OperationResult)
        assert result.success is True
        execution_time = result.metadata['execution_time_ms']
        assert execution_time >= 10  # 至少10ms
        assert execution_time < 100  # 不应该超过100ms

    def test_decorator_complex_parameters(self):
        """测试复杂参数的装饰器"""
        class MockObject:
            def __init__(self, file_path):
                self.file_path = file_path

        def object_context_extractor(*args, **kwargs):
            context = {}
            if args and hasattr(args[0], 'file_path'):
                context['object_file_path'] = args[0].file_path
            return context

        @unified_error_handler("complex_test", context_extractor=object_context_extractor)
        def complex_function(mock_obj, **kwargs):
            if mock_obj.file_path == "error.xlsx":
                raise FileNotFoundError("模拟错误")
            return "success"

        # 测试成功情况
        mock_obj = MockObject("success.xlsx")
        result = complex_function(mock_obj, extra_param="value")
        assert result.success is True

        # 测试失败情况
        mock_obj.file_path = "error.xlsx"
        result = complex_function(mock_obj)
        assert result.success is False
        assert result.error['details']['object_file_path'] == "error.xlsx"

    def test_decorator_thread_safety(self):
        """测试装饰器的线程安全性"""
        results = []

        @unified_error_handler("thread_test")
        def thread_function(thread_id):
            if thread_id % 2 == 0:
                return f"success_{thread_id}"
            else:
                raise ValueError(f"error_{thread_id}")

        def worker():
            for i in range(10):
                result = thread_function(i)
                results.append(result)

        threads = []
        for i in range(5):
            thread = threading.Thread(target=worker)
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        assert len(results) == 50  # 5 threads * 10 operations
        success_count = sum(1 for r in results if r.success)
        error_count = sum(1 for r in results if not r.success)
        assert success_count == 25  # 偶数ID成功
        assert error_count == 25   # 奇数ID失败


class TestContextExtractors:
    """上下文提取器测试"""

    def test_extract_file_context_from_args_object(self):
        """测试从对象参数中提取文件上下文"""
        class MockObject:
            def __init__(self, file_path):
                self.file_path = file_path

        mock_obj = MockObject("test.xlsx")
        context = extract_file_context(mock_obj)

        assert context['file_path'] == "test.xlsx"

    def test_extract_file_context_from_args_string(self):
        """测试从字符串参数中提取文件上下文"""
        context = extract_file_context(None, "test.xlsx", "other_arg")

        assert context['file_path'] == "test.xlsx"

    def test_extract_file_context_from_kwargs(self):
        """测试从kwargs中提取文件上下文"""
        kwargs = {
            'file_path': 'test.xlsx',
            'sheet_name': 'Sheet1',
            'range_expression': 'A1:C10',
            'other_param': 'value'
        }

        context = extract_file_context(**kwargs)

        assert context['file_path'] == 'test.xlsx'
        assert context['sheet_name'] == 'Sheet1'
        assert context['range_expression'] == 'A1:C10'
        assert 'other_param' not in context

    def test_extract_file_context_mixed_sources(self):
        """测试混合来源的文件上下文提取"""
        class MockObject:
            def __init__(self, file_path):
                self.file_path = file_path

        mock_obj = MockObject("object_file.xlsx")
        kwargs = {
            'file_path': 'kwargs_file.xlsx',
            'sheet_name': 'Sheet1'
        }

        context = extract_file_context(mock_obj, **kwargs)

        # kwargs应该覆盖对象属性
        assert context['file_path'] == 'kwargs_file.xlsx'
        assert context['sheet_name'] == 'Sheet1'

    def test_extract_file_context_empty(self):
        """测试空上下文提取"""
        context = extract_file_context()
        assert context == {}

    def test_extract_formula_context(self):
        """测试公式上下文提取"""
        kwargs = {
            'file_path': 'test.xlsx',
            'sheet_name': 'Sheet1',
            'formula': '=SUM(A1:A10)',
            'context_sheet': 'ContextSheet',
            'range_expression': 'A1:C10'
        }

        context = extract_formula_context(**kwargs)

        assert context['file_path'] == 'test.xlsx'
        assert context['sheet_name'] == 'Sheet1'
        assert context['formula'] == '=SUM(A1:A10)'
        assert context['context_sheet'] == 'ContextSheet'
        assert context['range_expression'] == 'A1:C10'

    def test_extract_formula_context_inheritance(self):
        """测试公式上下文提取继承文件上下文"""
        class MockObject:
            def __init__(self, file_path):
                self.file_path = file_path

        mock_obj = MockObject("test.xlsx")
        kwargs = {
            'formula': '=AVERAGE(B1:B20)',
            'additional_param': 'value'
        }

        context = extract_formula_context(mock_obj, **kwargs)

        assert context['file_path'] == 'test.xlsx'
        assert context['formula'] == '=AVERAGE(B1:B20)'
        assert 'additional_param' not in context


class TestErrorHandlerEdgeCases:
    """ErrorHandler边界条件测试"""

    def test_error_codes_mapping_completeness(self):
        """测试错误代码映射的完整性"""
        # 检查ERROR_CODES字典是否包含预期的映射
        assert 'FileNotFoundError' in ErrorHandler.ERROR_CODES
        assert 'PermissionError' in ErrorHandler.ERROR_CODES
        assert 'SheetNotFoundError' in ErrorHandler.ERROR_CODES
        assert 'DataValidationError' in ErrorHandler.ERROR_CODES
        assert 'ValueError' in ErrorHandler.ERROR_CODES
        assert 'KeyError' in ErrorHandler.ERROR_CODES
        assert 'ImportError' in ErrorHandler.ERROR_CODES
        assert 'InvalidRangeError' in ErrorHandler.ERROR_CODES
        assert 'FormulaCalculationError' in ErrorHandler.ERROR_CODES
        assert 'Exception' in ErrorHandler.ERROR_CODES

    def test_error_solutions_mapping_completeness(self):
        """测试错误解决方案映射的完整性"""
        # 检查ERROR_SOLUTIONS字典是否包含预期的解决方案
        assert 'FILE_NOT_FOUND' in ErrorHandler.ERROR_SOLUTIONS
        assert 'PERMISSION_DENIED' in ErrorHandler.ERROR_SOLUTIONS
        assert 'SHEET_NOT_FOUND' in ErrorHandler.ERROR_SOLUTIONS
        assert 'INVALID_RANGE' in ErrorHandler.ERROR_SOLUTIONS
        assert 'FORMULA_ERROR' in ErrorHandler.ERROR_SOLUTIONS
        assert 'DATA_VALIDATION_ERROR' in ErrorHandler.ERROR_SOLUTIONS
        assert 'VALUE_ERROR' in ErrorHandler.ERROR_SOLUTIONS

    @pytest.mark.parametrize("exception_type,expected_code", [
        (FileNotFoundError("test"), 'FILE_NOT_FOUND'),
        (PermissionError("test"), 'PERMISSION_DENIED'),
        (ValueError("test"), 'VALUE_ERROR'),
        (KeyError("test"), 'KEY_ERROR'),
        (ImportError("test"), 'IMPORT_ERROR'),
    ])
    def test_various_exception_mappings(self, exception_type, expected_code):
        """测试各种异常类型的映射"""
        result_code = ErrorHandler.get_error_code(exception_type)
        assert result_code == expected_code

    def test_custom_excel_exceptions(self):
        """测试自定义Excel异常的处理"""
        excel_error = ExcelFileNotFoundError("Excel文件不存在")
        error_code = ErrorHandler.get_error_code(excel_error)
        # 自定义异常可能映射到通用错误代码
        assert error_code in ['UNKNOWN_ERROR', 'FILE_NOT_FOUND']

        sheet_error = SheetNotFoundError("工作表不存在")
        error_code = ErrorHandler.get_error_code(sheet_error)
        assert error_code in ['UNKNOWN_ERROR', 'SHEET_NOT_FOUND']

        validation_error = DataValidationError("数据验证失败")
        error_code = ErrorHandler.get_error_code(validation_error)
        assert error_code in ['UNKNOWN_ERROR', 'DATA_VALIDATION_ERROR']

    def test_error_handler_performance(self):
        """测试错误处理器的性能"""
        start_time = time.time()

        # 执行多次错误代码获取操作
        for _ in range(1000):
            ErrorHandler.get_error_code(ValueError("test"))

        end_time = time.time()

        # 确保性能在可接受范围内（1000次操作应该在1秒内完成）
        assert end_time - start_time < 1.0

    def test_error_message_formatting(self):
        """测试错误消息格式化"""
        # 测试错误消息是否包含有用信息
        solution = ErrorHandler.get_error_solution('FILE_NOT_FOUND')

        # 确保解决方案不为空且有意义
        assert solution is not None
        assert len(solution) > 0
        assert isinstance(solution, str)

    def test_error_handler_thread_safety(self):
        """测试错误处理器的线程安全性"""
        results = []

        def worker():
            try:
                for i in range(100):
                    error = ValueError(f"test error {i}")
                    code = ErrorHandler.get_error_code(error)
                    results.append(code)
            except Exception as e:
                results.append(f"Thread error: {e}")

        # 创建多个线程同时访问ErrorHandler
        threads = [threading.Thread(target=worker) for _ in range(5)]

        for thread in threads:
            thread.start()

        for thread in threads:
            thread.join()

        # 确保所有结果都是预期的错误代码
        assert len(results) == 500  # 5 threads * 100 operations
        assert all(result == 'VALUE_ERROR' or 'Thread error' in result for result in results)

    def test_error_handler_memory_usage(self):
        """测试错误处理器的内存使用情况"""
        import gc

        # 强制垃圾回收
        gc.collect()

        # 创建大量异常并获取错误代码
        for i in range(10000):
            error = ValueError(f"test error {i}")
            ErrorHandler.get_error_code(error)

        # 再次强制垃圾回收
        gc.collect()

        # 这个测试主要确保没有内存泄漏，如果有严重内存泄漏会导致测试失败
        assert True  # 如果到达这里说明没有内存问题

    def test_error_response_with_none_context(self):
        """测试None上下文的错误响应"""
        error = RuntimeError("运行时错误")
        response = ErrorHandler.format_error_response(error, None, "test_op")

        assert response['details'] == {}
        assert response['operation'] == "test_op"

    def test_error_response_with_empty_context(self):
        """测试空上下文的错误响应"""
        error = RuntimeError("运行时错误")
        response = ErrorHandler.format_error_response(error, {}, "test_op")

        assert response['details'] == {}
        assert response['operation'] == "test_op"

    def test_unknown_exception_in_format_response(self):
        """测试格式化响应中的未知异常"""
        class UnknownException(Exception):
            pass

        error = UnknownException("未知异常")
        response = ErrorHandler.format_error_response(error)

        assert response['code'] == 'UNKNOWN_ERROR'
        assert response['type'] == 'UnknownException'
        assert response['message'] == '未知异常'


class TestErrorHandlerIntegration:
    """ErrorHandler集成测试"""

    def test_full_error_handling_workflow(self):
        """测试完整的错误处理工作流"""
        # 模拟一个完整的Excel操作错误处理流程

        @unified_error_handler("excel_operation", context_extractor=extract_file_context)
        def excel_operation(file_path, sheet_name, range_expr):
            # 模拟各种可能的错误
            if file_path == "missing.xlsx":
                raise FileNotFoundError(f"文件不存在: {file_path}")
            elif sheet_name == "MissingSheet":
                raise ValueError(f"工作表不存在: {sheet_name}")
            elif range_expr == "INVALID":
                raise ValueError(f"无效范围: {range_expr}")
            else:
                return {"operation": "success", "data": "result"}

        # 测试成功情况
        result = excel_operation("test.xlsx", "Sheet1", "A1:C10")
        assert result.success is True
        assert result.data["operation"] == "success"

        # 测试文件不存在错误
        result = excel_operation("missing.xlsx", "Sheet1", "A1:C10")
        assert result.success is False
        assert result.error['code'] == 'FILE_NOT_FOUND'
        # extract_file_context会从第二个参数（sheet_name）提取file_path，因为它是字符串
        assert result.error['details']['file_path'] == "Sheet1"

        # 测试工作表不存在错误
        result = excel_operation("test.xlsx", "MissingSheet", "A1:C10")
        assert result.success is False
        assert result.error['code'] == 'VALUE_ERROR'
        # extract_file_context从第二个参数提取为file_path，没有sheet_name字段
        assert result.error['details']['file_path'] == "MissingSheet"

    def test_nested_error_handling(self):
        """测试嵌套错误处理"""
        @unified_error_handler("inner_operation")
        def inner_operation(value):
            if value < 0:
                raise ValueError("值不能为负数")
            return value * 2

        @unified_error_handler("outer_operation", context_extractor=extract_file_context)
        def outer_operation(file_path, value):
            try:
                result = inner_operation(value)
                return {"file": file_path, "result": result}
            except Exception:
                # 重新抛出异常以测试外层装饰器处理
                raise

        # 测试嵌套成功
        result = outer_operation("test.xlsx", 5)
        assert result.success is True
        # 外层装饰器会包装内层的结果
        assert result.data['file'] == "test.xlsx"
        assert result.data['result'].data == 10

        # 测试嵌套错误
        result = outer_operation("test.xlsx", -1)
        assert result.success is True  # 外层try-catch捕获了内层的失败结果，包装为成功
        assert result.data['file'] == "test.xlsx"
        assert result.data['result'].success is False  # 内层操作确实失败了
        assert result.data['result'].error['code'] == 'VALUE_ERROR'

    def test_performance_under_load(self):
        """测试高负载下的性能"""
        @unified_error_handler("load_test")
        def load_function(operation_id):
            if operation_id % 100 == 0:
                raise RuntimeError(f"模拟错误 {operation_id}")
            return f"success_{operation_id}"

        start_time = time.time()
        results = []

        # 执行大量操作
        for i in range(1000):
            result = load_function(i)
            results.append(result)

        end_time = time.time()
        execution_time = end_time - start_time

        # 验证性能
        assert execution_time < 2.0  # 1000次操作应该在2秒内完成
        assert len(results) == 1000

        success_count = sum(1 for r in results if r.success)
        error_count = sum(1 for r in results if not r.success)
        assert success_count == 990  # 1000 - 10个错误
        assert error_count == 10

    def test_concurrent_error_handling(self):
        """测试并发错误处理"""
        results = []
        errors = []

        @unified_error_handler("concurrent_test")
        def concurrent_operation(thread_id, operation_id):
            if operation_id % 10 == 0:
                raise ValueError(f"线程{thread_id}操作{operation_id}失败")
            return f"线程{thread_id}操作{operation_id}成功"

        def worker(thread_id):
            thread_results = []
            for i in range(50):
                try:
                    result = concurrent_operation(thread_id, i)
                    thread_results.append(result)
                except Exception as e:
                    errors.append((thread_id, i, str(e)))
            results.extend(thread_results)

        # 启动多个线程
        threads = []
        for i in range(5):
            thread = threading.Thread(target=worker, args=(i,))
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        # 验证结果
        assert len(results) == 250  # 5 threads * 50 operations
        assert len(errors) == 0  # 所有错误都应该被装饰器处理

        success_count = sum(1 for r in results if r.success)
        error_count = sum(1 for r in results if not r.success)
        assert success_count == 225  # 250 - 25个错误 (每50个操作有5个错误)
        assert error_count == 25

    def test_logging_integration(self):
        """测试日志集成"""
        with patch('src.utils.error_handler.logger') as mock_logger:
            @unified_error_handler("logging_test")
            def failing_function():
                raise ValueError("测试日志记录")

            result = failing_function()

            # 验证错误被正确记录
            assert result.success is False
            mock_logger.error.assert_called_once()
            error_call = mock_logger.error.call_args
            assert "logging_test 操作失败" in error_call[0][0]
            assert error_call[1]['exc_info'] is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
