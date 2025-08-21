"""
Excel MCP Server - 统一错误处理装饰器

提供统一的错误处理和响应格式标准化
"""

import logging
import time
import functools
from typing import Any, Dict, Callable, Optional
from ..models.types import OperationResult

logger = logging.getLogger(__name__)


class ErrorHandler:
    """统一错误处理器"""
    
    # 错误代码映射
    ERROR_CODES = {
        'FileNotFoundError': 'FILE_NOT_FOUND',
        'PermissionError': 'PERMISSION_DENIED', 
        'SheetNotFoundError': 'SHEET_NOT_FOUND',
        'DataValidationError': 'DATA_VALIDATION_ERROR',
        'ValueError': 'VALUE_ERROR',
        'KeyError': 'KEY_ERROR',
        'ImportError': 'IMPORT_ERROR',
        'Exception': 'GENERAL_ERROR'
    }
    
    @staticmethod
    def get_error_code(exception: Exception) -> str:
        """获取错误代码"""
        exception_name = type(exception).__name__
        return ErrorHandler.ERROR_CODES.get(exception_name, 'UNKNOWN_ERROR')
    
    @staticmethod
    def format_error_response(
        error: Exception,
        context: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """格式化错误响应"""
        error_code = ErrorHandler.get_error_code(error)
        
        return {
            'code': error_code,
            'message': str(error),
            'type': type(error).__name__,
            'details': context or {}
        }


def unified_error_handler(
    operation_name: str,
    context_extractor: Optional[Callable] = None
):
    """
    统一错误处理装饰器
    
    Args:
        operation_name: 操作名称，用于日志记录
        context_extractor: 上下文提取函数，用于提供错误详情
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> OperationResult:
            start_time = time.time()
            
            try:
                # 执行原函数
                result = func(*args, **kwargs)
                
                # 如果返回的不是OperationResult，包装它
                if not isinstance(result, OperationResult):
                    result = OperationResult(
                        success=True,
                        data=result
                    )
                
                # 添加元数据
                execution_time = (time.time() - start_time) * 1000  # 毫秒
                if not result.metadata:
                    result.metadata = {}
                
                result.metadata.update({
                    'operation': operation_name,
                    'execution_time_ms': round(execution_time, 2),
                    'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
                })
                
                return result
                
            except Exception as e:
                # 记录错误日志
                logger.error(f"{operation_name} 操作失败: {e}", exc_info=True)
                
                # 提取上下文信息
                context = {}
                if context_extractor:
                    try:
                        context = context_extractor(*args, **kwargs)
                    except Exception as ctx_error:
                        logger.warning(f"提取错误上下文失败: {ctx_error}")
                
                # 添加执行信息到上下文
                context.update({
                    'operation': operation_name,
                    'execution_time_ms': round((time.time() - start_time) * 1000, 2),
                    'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
                })
                
                # 格式化错误信息
                error_info = ErrorHandler.format_error_response(e, context)
                
                return OperationResult(
                    success=False,
                    error=error_info
                )
        
        return wrapper
    return decorator


def extract_file_context(*args, **kwargs) -> Dict[str, Any]:
    """提取文件操作相关的上下文信息"""
    context = {}
    
    # 尝试从参数中提取文件路径
    if args:
        if hasattr(args[0], 'file_path'):
            context['file_path'] = args[0].file_path
        elif len(args) > 1 and isinstance(args[1], str):
            context['file_path'] = args[1]
    
    # 从kwargs中提取
    if 'file_path' in kwargs:
        context['file_path'] = kwargs['file_path']
    if 'sheet_name' in kwargs:
        context['sheet_name'] = kwargs['sheet_name']
    if 'range_expression' in kwargs:
        context['range_expression'] = kwargs['range_expression']
    
    return context


def extract_formula_context(*args, **kwargs) -> Dict[str, Any]:
    """提取公式操作相关的上下文信息"""
    context = extract_file_context(*args, **kwargs)
    
    if 'formula' in kwargs:
        context['formula'] = kwargs['formula']
    if 'context_sheet' in kwargs:
        context['context_sheet'] = kwargs['context_sheet']
    
    return context
