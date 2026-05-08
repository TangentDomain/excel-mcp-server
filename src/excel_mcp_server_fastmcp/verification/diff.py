"""结构化差异比较与归一化工具。"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any


UNSTABLE_QUERY_INFO_KEYS = {
    "execution_time_ms",
    "markdown_table",
    "formatted_output",
    "sql_query",
    "record_count",
}


def normalize_value(value: Any) -> Any:
    """把结果归一化成稳定 JSON 结构。"""

    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(key): normalize_value(val) for key, val in value.items()}
    if isinstance(value, tuple):
        return [normalize_value(item) for item in value]
    if isinstance(value, list):
        return [normalize_value(item) for item in value]
    if isinstance(value, set):
        normalized = [normalize_value(item) for item in value]
        return sorted(normalized, key=lambda item: repr(item))
    return value


def normalize_select_result(result: dict[str, Any]) -> dict[str, Any]:
    """提取查询结果中稳定、适合做 baseline 的部分。"""

    query_info = result.get("query_info", {}) or {}
    stable_query_info = {
        key: normalize_value(value)
        for key, value in query_info.items()
        if key not in UNSTABLE_QUERY_INFO_KEYS
    }
    return {
        "success": bool(result.get("success")),
        "data": normalize_value(result.get("data", [])),
        "query_info": stable_query_info,
    }


def normalize_update_result(result: dict[str, Any]) -> dict[str, Any]:
    """提取更新结果中稳定、适合做 baseline 的部分。"""

    return {
        "success": bool(result.get("success")),
        "affected_rows": result.get("affected_rows", 0),
        "changes": normalize_value(result.get("changes", [])),
    }


def compare_structured(expected: Any, actual: Any, path: str = "$") -> list[dict[str, Any]]:
    """递归比较两个结构化对象，返回可定位的差异列表。"""

    diffs: list[dict[str, Any]] = []

    if type(expected) is not type(actual):
        diffs.append({
            "path": path,
            "type": "type_mismatch",
            "expected_type": type(expected).__name__,
            "actual_type": type(actual).__name__,
            "expected": expected,
            "actual": actual,
        })
        return diffs

    if isinstance(expected, Mapping):
        expected_keys = set(expected.keys())
        actual_keys = set(actual.keys())
        for key in sorted(expected_keys - actual_keys, key=str):
            diffs.append({
                "path": f"{path}.{key}",
                "type": "missing_key",
                "expected": expected[key],
                "actual": None,
            })
        for key in sorted(actual_keys - expected_keys, key=str):
            diffs.append({
                "path": f"{path}.{key}",
                "type": "extra_key",
                "expected": None,
                "actual": actual[key],
            })
        for key in sorted(expected_keys & actual_keys, key=str):
            diffs.extend(compare_structured(expected[key], actual[key], f"{path}.{key}"))
        return diffs

    if isinstance(expected, Sequence) and not isinstance(expected, (str, bytes, bytearray)):
        min_len = min(len(expected), len(actual))
        for index in range(min_len):
            diffs.extend(compare_structured(expected[index], actual[index], f"{path}[{index}]"))
        for index in range(min_len, len(expected)):
            diffs.append({
                "path": f"{path}[{index}]",
                "type": "missing_item",
                "expected": expected[index],
                "actual": None,
            })
        for index in range(min_len, len(actual)):
            diffs.append({
                "path": f"{path}[{index}]",
                "type": "extra_item",
                "expected": None,
                "actual": actual[index],
            })
        return diffs

    if expected != actual:
        diffs.append({
            "path": path,
            "type": "value_mismatch",
            "expected": expected,
            "actual": actual,
        })

    return diffs


__all__ = [
    "compare_structured",
    "normalize_select_result",
    "normalize_update_result",
    "normalize_value",
]
