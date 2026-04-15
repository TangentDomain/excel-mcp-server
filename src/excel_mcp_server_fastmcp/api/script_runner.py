"""Python脚本执行引擎 - 直接运行Python代码操作Excel文件。

暴露现有API函数给用户代码，复用已测试的业务逻辑。
"""

import io
import logging
import traceback
from contextlib import redirect_stdout
from functools import partial
from typing import Any

logger = logging.getLogger(__name__)

from .advanced_sql_query import (
    execute_advanced_delete_query,
    execute_advanced_insert_query,
    execute_advanced_sql_query,
    execute_advanced_update_query,
)
from .excel_operations import ExcelOperations


def _query_wrapper(file_path: str, sql: str) -> list:
    """query快捷函数：自动提取data字段，节省token。

    Fix BUG-004: 失败时抛出RuntimeError而非静默返回错误字典，
    避免用户代码把 {success: False, ...} 当成空数据处理。
    正常空结果(有表头无数据行) 返回 [[col1, col2, ...]]。
    """
    result = execute_advanced_sql_query(file_path, sql)
    if result.get("success"):
        data = result["data"]  # [[headers], [row1], ...]
        # 空结果诊断：正常空结果至少包含表头行
        # 如果 data 为空列表，说明返回格式异常（可能是 include_headers=False 的边界情况）
        if not data:
            logger.warning(
                f"query() 返回成功但 data 为空列表 sql={sql[:80]!r} "
                f"keys={list(result.keys())}"
            )
            # 尝试从 query_info 中恢复列名
            cols = result.get("query_info", {}).get("returned_columns", [])
            if cols:
                return [cols]
        return data
    # 失败时抛异常，让用户代码能 catch 或直接看到错误
    error_msg = result.get("message", "未知错误")
    raise RuntimeError(f"SQL查询失败: {error_msg}")


def _safe_repr(value: Any, max_length: int = 2000) -> Any:
    """安全repr，截断过长输出。"""
    if value is None:
        return None
    try:
        s = repr(value)
        if len(s) > max_length:
            return s[:max_length] + f"... (truncated, total {len(s)} chars)"
        return s
    except Exception:
        return f"<repr failed: {type(value).__name__}>"


def execute_python_script(
    file_path: str,
    code: str,
    sheet_name: str | None = None,
    timeout: int = 30,
) -> dict[str, Any]:
    """执行Python代码操作Excel文件。

    用户代码可直接调用现有API函数，file_path已预绑定。
    也可通过 ExcelOperations 调用所有Excel操作。

    Args:
        file_path: Excel文件路径
        code: Python代码
        sheet_name: 工作表名称（传递给需要sheet_name的API）
        timeout: 超时秒数（1-120）
    """
    timeout = max(1, min(timeout, 120))
    stdout_buf = io.StringIO()
    result_value = None

    try:
        # 构建执行环境：预绑定file_path的便捷函数
        user_globals = {
            "file_path": file_path,
            "sheet_name": sheet_name,
            # SQL快捷函数（file_path已预绑定）
            "query": partial(_query_wrapper, file_path),
            "update": partial(execute_advanced_update_query, file_path),
            "insert": partial(execute_advanced_insert_query, file_path),
            "delete": partial(execute_advanced_delete_query, file_path),
            # 完整API类（所有操作）
            "ExcelOperations": ExcelOperations,
        }

        # 执行代码
        with redirect_stdout(stdout_buf):
            try:
                compiled = compile(code, "<user_script>", "eval")
                result_value = eval(compiled, user_globals)
            except SyntaxError:
                compiled = compile(code, "<user_script>", "exec")
                exec(compiled, user_globals)
                result_value = user_globals.get("result", None)

        return {
            "success": True,
            "message": "脚本执行成功",
            "data": {
                "result": _safe_repr(result_value),
                "stdout": stdout_buf.getvalue(),
            },
            "meta": {"file_path": file_path, "sheet_name": sheet_name},
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"脚本执行失败: {type(e).__name__}: {e}",
            "data": {
                "stdout": stdout_buf.getvalue(),
                "traceback": traceback.format_exc(),
            },
            "meta": {"file_path": file_path},
        }
