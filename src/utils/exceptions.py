"""
Excel MCP Server - 自定义异常类

定义了项目中使用的所有自定义异常
"""


class ExcelException(Exception):
    """Excel操作基础异常"""
    pass


class FileNotFoundError(ExcelException):
    """文件不存在异常"""
    pass


class InvalidFormatError(ExcelException):
    """无效文件格式异常"""
    pass


class InvalidRangeError(ExcelException):
    """无效范围异常"""
    pass


class SheetNotFoundError(ExcelException):
    """工作表不存在异常"""
    pass


class DataValidationError(ExcelException):
    """数据验证异常"""
    pass


class OperationLimitError(ExcelException):
    """操作限制异常"""
    pass
