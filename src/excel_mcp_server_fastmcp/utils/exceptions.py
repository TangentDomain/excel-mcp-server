"""
Excel MCP Server - 自定义异常类

定义了项目中使用的所有自定义异常，包含详细的错误信息和修复建议
"""
from typing import List


class ExcelException(Exception):
    """Excel操作基础异常"""
    
    def __init__(self, message: str, hint: str = None, suggested_fix: str = None):
        self.message = message
        self.hint = hint
        self.suggested_fix = suggested_fix
        super().__init__(self.get_formatted_message())
    
    def get_formatted_message(self) -> str:
        """格式化错误消息，包含提示和修复建议"""
        parts = [f"Excel操作错误: {self.message}"]
        if self.hint:
            parts.append(f"💡 提示: {self.hint}")
        if self.suggested_fix:
            parts.append(f"🔧 建议: {self.suggested_fix}")
        return "\n".join(parts)


class ExcelFileNotFoundError(FileNotFoundError):
    """文件不存在异常 - 继承自Python内置FileNotFoundError"""
    
    def __init__(self, file_path: str, hint: str = None):
        message = f"Excel文件不存在: {file_path}"
        suggested_fix = f"请检查文件路径是否正确，或使用绝对路径。确保文件确实存在于指定位置。"
        super().__init__(message, hint, suggested_fix)


class InvalidFormatError(ExcelException):
    """无效文件格式异常"""
    
    def __init__(self, file_path: str, expected_formats: List[str] = None):
        message = f"无效的Excel文件格式: {file_path}"
        hint = "文件必须是Excel格式（.xlsx, .xls）"
        expected = f"支持的格式: {', '.join(expected_formats or ['.xlsx', '.xls'])}" if expected_formats else ""
        suggested_fix = f"请确保文件是有效的Excel格式。{expected}"
        super().__init__(message, hint, suggested_fix)


class InvalidRangeError(ExcelException):
    """无效范围异常"""
    
    def __init__(self, range_expression: str, reason: str = None):
        message = f"无效的Excel范围表达式: {range_expression}"
        hint = reason or "范围表达式格式不正确"
        suggested_fix = "请使用标准Excel范围格式，例如: Sheet1!A1:C10, A1:B5, Sheet1!A:D"
        super().__init__(message, hint, suggested_fix)


class SheetNotFoundError(ExcelException):
    """工作表不存在异常"""
    
    def __init__(self, sheet_name: str, available_sheets: List[str] = None):
        message = f"工作表不存在: {sheet_name}"
        hint = "工作表名称区分大小写"
        available = f"可用工作表: {', '.join(available_sheets)}" if available_sheets else ""
        suggested_fix = f"请检查工作表名称是否正确。{available}"
        super().__init__(message, hint, suggested_fix)


class DataValidationError(ExcelException):
    """数据验证异常"""
    
    def __init__(self, validation_type: str, details: str = None):
        message = f"数据验证失败 ({validation_type})"
        hint = details or "输入的数据不符合要求"
        suggested_fix = "请检查数据的格式、类型和范围是否符合要求"
        super().__init__(message, hint, suggested_fix)


class OperationLimitError(ExcelException):
    """操作限制异常"""
    
    def __init__(self, operation: str, limit: str, reason: str = None):
        message = f"操作超出限制: {operation}"
        hint = f"限制: {limit}"
        detail = f"原因: {reason}" if reason else ""
        suggested_fix = f"请减小操作规模或分批执行。{detail}"
        super().__init__(message, hint + f"\n{detail}", suggested_fix)


class ExcelMCPError(ExcelException):
    """Excel MCP 通用操作异常"""
    
    def __init__(self, operation: str, error_code: str = None, **kwargs):
        message = f"Excel MCP操作失败: {operation}"
        if error_code:
            message += f" (错误代码: {error_code})"
        super().__init__(message, **kwargs)
