"""
Excel MCP Server - 文本处理工具

提供从 openpyxl RichText 对象提取纯文本等文本处理功能
"""

from typing import Any


def extract_rich_text(title_obj: Any) -> str:
    """从 openpyxl RichText 对象提取纯文本

    Args:
        title_obj: openpyxl 的 RichText 对象或文本对象

    Returns:
        提取的纯文本字符串，最多返回前100个字符
    """
    if not title_obj:
        return ""
    try:
        t = title_obj.text
        if isinstance(t, str):
            return t
        if hasattr(t, "p"):
            parts = []
            for p in t.p:
                for r in p.r or []:
                    if hasattr(r, "t"):
                        parts.append(r.t)
            return "".join(parts)
    except Exception:
        pass
    return str(title_obj)[:100]
