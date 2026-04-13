#!/usr/bin/env python3
"""
Excel MCP Server - 结果格式化工具

提供MCP响应结果的格式化功能，包括JSON序列化、紧凑数组格式转换、null值清理等。

主要功能：
1. 结果格式化：将OperationResult对象转换为MCP响应格式
2. JSON序列化：处理dataclass、枚举等复杂类型的序列化
3. 紧凑数组转换：将结构化比较数据转换为紧凑的数组格式
4. null值清理：递归清理响应中的null/None值

技术特性：
- 支持自动回退机制，当JSON序列化失败时使用原始转换方案
- 防重复转换，避免对已转换的紧凑数组格式重复处理
- 深度null值清理，确保响应数据的简洁性
"""

import json
from enum import Enum
from typing import Any

# ==================== 主干 ====================


def _strip_defaults(obj: Any) -> Any:
    """
    递归移除数据中的默认值：None、空字符串、空容器

    @intention 进一步减少响应数据体积，过滤掉无意义的默认值

    Args:
        obj: 待清理的任意类型对象

    Returns:
        Any: 清理后的对象，移除了None/空字符串/空容器

    Rules:
        - None值 → 移除
        - 空字符串 "" → 移除
        - 空列表 [] → 移除
        - 空字典 {} → 移除
        - 假值但有意义的数据（如False、0）→ 保留
        - 递归处理嵌套结构
        - 特例：data、matches等重要字段即使为空也保留
    """
    if obj is None:
        return None
    elif isinstance(obj, str):
        return obj if obj.strip() != "" else None
    elif isinstance(obj, dict):
        cleaned = {}
        # CellInfo语义检测：coordinate+value结构中，value=None/''代表空单元格
        _is_cell_info = "coordinate" in obj
        for key, value in obj.items():
            if key in [
                "data",
                "matches",
                "differences",
                "results",
                "row_differences",
                "headers",
            ]:
                # 重要字段即使为空也要保留
                cleaned[key] = _strip_defaults(value)
            elif key == "value" and _is_cell_info:
                # 单元格value：保留None和空字符串（空单元格有语义意义）
                cleaned[key] = value
            else:
                # 普通字段过滤掉空值
                stripped_value = _strip_defaults(value)
                if stripped_value is not None and stripped_value != {} and stripped_value != []:
                    cleaned[key] = stripped_value
        return cleaned
    elif isinstance(obj, list):
        cleaned = []
        for item in obj:
            stripped_item = _strip_defaults(item)
            if stripped_item is not None and stripped_item != {} and stripped_item != []:
                cleaned.append(stripped_item)
        return cleaned
    else:
        return obj


def _optimize_excel_data(obj: Any) -> Any:
    """
    Excel数据专项优化：移除冗余字段，简化数据结构

    @intention 专门针对Excel操作的响应数据进行进一步优化，移除冗余信息

    Args:
        obj: Excel操作响应数据

    Returns:
        Any: 优化后的数据

    Optimizations:
        - 移除get_headers中的冗余字段（field_names/headers重复）
        - 移除单元格数据中的coordinate字段（与位置信息冗余）
        - 简化元数据结构
    """
    if isinstance(obj, dict):
        cleaned = {}

        # 专门处理headers数据中的冗余
        if obj.get("success") and obj.get("data") and isinstance(obj.get("data"), dict):
            data = obj["data"]

            # 如果是headers响应，移除field_names与headers的重复
            if "headers" in data and "field_names" in data:
                if data["headers"] == data["field_names"]:
                    # 保留headers，移除field_names
                    cleaned_data = data.copy()
                    cleaned_data.pop("field_names", None)
                    cleaned["data"] = cleaned_data
                else:
                    cleaned["data"] = data
            else:
                cleaned["data"] = data
        else:
            cleaned["data"] = obj.get("data")

        # 处理其他字段
        for key, value in obj.items():
            if key != "data":
                cleaned[key] = value

        return cleaned
    elif isinstance(obj, list):
        # 如果是单元格数据列表，移除coordinate字段
        if len(obj) > 0 and isinstance(obj[0], dict) and "coordinate" in obj[0]:
            optimized_list = []
            for item in obj:
                if isinstance(item, dict):
                    # 移除coordinate字段，只保留value
                    optimized_item = {"value": item.get("value")}
                    optimized_list.append(optimized_item)
                else:
                    optimized_list.append(item)
            return optimized_list
        else:
            return obj
    else:
        return obj


def _normalize_error_key(result_dict: dict[str, Any]) -> None:
    """
    统一错误信息键名：将error字段映射到message，确保MCP客户端只需检查message键。

    核心层OperationResult使用error字段，server层和advanced_sql_query使用message字段。
    此函数在输出边界统一，避免MCP客户端需要检查两个不同的键。
    仅处理字符串类型的error值（结构化错误对象如error_handler的error dict不处理）。
    """
    if not isinstance(result_dict, dict):
        return
    error_val = result_dict.get("error")
    if isinstance(error_val, str) and not result_dict.get("message"):
        result_dict["message"] = result_dict.pop("error")


def format_operation_result(result) -> dict[str, Any]:
    """
    格式化操作结果为MCP响应格式

    @intention 统一处理OperationResult对象的格式化，确保MCP响应的一致性和简洁性

    Args:
        result: OperationResult对象，包含操作的成功状态、数据、错误信息等

    Returns:
        Dict[str, Any]: 格式化后的字典，已清理null值并转换为紧凑格式

    Features:
        - JSON序列化：自动处理dataclass和枚举类型
        - 紧凑转换：结构化比较数据转换为数组格式，减少60-80%体积
        - null清理：递归移除空值、空字典、空列表
        - 默认值清理：移除None、空字符串、空容器
        - 错误回退：序列化失败时自动切换到原始转换方案
    """
    # 步骤1：尝试JSON序列化方案（推荐）
    try:
        serialized_dict = _serialize_to_json_dict(result)

        # 步骤2：转换为紧凑数组格式（仅用于结构化比较结果）
        if serialized_dict.get("data"):
            serialized_dict["data"] = _convert_to_compact_array_format(serialized_dict["data"])

        # 步骤3：应用深度null清理
        cleaned_dict = _deep_clean_nulls(serialized_dict)

        # 步骤4：应用默认值清理（REQ-027新功能）
        stripped_dict = _strip_defaults(cleaned_dict)

        # 步骤5：Excel数据专项优化（REQ-027优化功能）
        optimized_dict = _optimize_excel_data(stripped_dict)

        # 步骤6：统一错误键 — 将error映射到message，确保MCP客户端只需检查message
        _normalize_error_key(optimized_dict)
        return optimized_dict

    except Exception as e:
        # 步骤6：JSON方案失败，使用回退方案
        fallback = _fallback_format_result(result, e)
        _strip_defaults(fallback)  # 回退方案也应用默认值清理
        _normalize_error_key(fallback)
        return fallback


# ==================== 分支 ====================


# --- JSON序列化处理 ---
def _serialize_to_json_dict(result) -> dict[str, Any]:
    """
    将OperationResult对象序列化为JSON字典

    Args:
        result: OperationResult对象

    Returns:
        Dict[str, Any]: JSON序列化后的字典

    Raises:
        Exception: 当对象包含不可序列化内容时抛出异常
    """

    def _json_serializer(obj):
        """自定义JSON序列化器，专门处理dataclass和枚举"""
        if isinstance(obj, Enum):
            return obj.value
        elif hasattr(obj, "__dict__"):
            return obj.__dict__
        else:
            return str(obj)

    # 转换为JSON字符串再解析回字典，自动处理复杂类型
    # separators=(',', ':') 移除空格，确保紧凑输出
    json_str = json.dumps(result, default=_json_serializer, ensure_ascii=False, separators=(",", ":"))
    return json.loads(json_str)


def _fallback_format_result(result, original_exception: Exception) -> dict[str, Any]:
    """
    JSON序列化失败时的回退格式化方案

    Args:
        result: OperationResult对象
        original_exception: 原始异常信息

    Returns:
        Dict[str, Any]: 使用原始转换的格式化结果
    """
    response = {"success": result.success}

    if result.success:
        # 统一数据处理，避免重复
        if result.data is not None:
            try:
                if hasattr(result.data, "__dict__"):
                    # 数据类转换为字典并放在data字段中
                    response["data"] = result.data.__dict__
                elif isinstance(result.data, list):
                    # 列表元素逐个处理
                    response["data"] = [item.__dict__ if hasattr(item, "__dict__") else item for item in result.data]
                else:
                    response["data"] = result.data
            except Exception:
                # 数据处理也失败，使用字符串表示
                response["data"] = str(result.data)

        # 分离处理metadata，避免键冲突
        if result.metadata:
            try:
                response["metadata"] = result.metadata
            except Exception:
                response["metadata"] = str(result.metadata)

        if result.message:
            response["message"] = result.message
    else:
        response["error"] = result.error

    # 应用默认值清理（REQ-027新功能）
    return _strip_defaults(response)


# --- 紧凑数组格式转换 ---
def _convert_to_compact_array_format(data: dict[str, Any]) -> dict[str, Any]:
    """
    将结构化比较结果转换为紧凑的数组格式

    @intention 减少JSON传输体积，将重复的键名转换为数组索引，特别适合大量行差异数据

    Args:
        data: 包含row_differences的结构化比较数据

    Returns:
        Dict[str, Any]: 转换后的紧凑格式数据，如果不符合转换条件则返回原数据

    Format:
        转换前: [{"row_id": "1001", "difference_type": "row_added", ...}, ...]
        转换后: [["row_id", "difference_type", ...], ["1001", "row_added", ...], ...]
    """
    # 检查数据是否符合转换条件
    if not isinstance(data, dict) or "row_differences" not in data:
        return data

    row_differences = data.get("row_differences", [])
    if not row_differences:
        return data

    # 防止重复转换：检查是否已经是数组格式
    if _is_already_compact_format(row_differences):
        return data

    # 执行转换为紧凑数组格式
    compact_differences = _build_compact_array(row_differences)

    # 创建新的数据副本，替换row_differences
    new_data = data.copy()
    new_data["row_differences"] = compact_differences
    return new_data


def _is_already_compact_format(row_differences: list[Any]) -> bool:
    """
    检查数据是否已经是紧凑数组格式

    Args:
        row_differences: 行差异数据列表

    Returns:
        bool: True表示已经是紧凑格式，False表示需要转换
    """
    return isinstance(row_differences, list) and len(row_differences) > 0 and isinstance(row_differences[0], list)


def _build_compact_array(row_differences: list[dict[str, Any]]) -> list[list[Any]]:
    """
    构建紧凑数组格式的行差异数据

    Args:
        row_differences: 原始的行差异字典列表

    Returns:
        List[List[Any]]: 紧凑数组格式，第一行为字段定义，后续行为数据
    """
    compact_differences = []

    # 第一行：字段定义（作为数据索引的说明）
    field_definitions = [
        "row_id",
        "difference_type",
        "row_index1",
        "row_index2",
        "sheet_name",
        "field_differences",
    ]
    compact_differences.append(field_definitions)

    # 后续行：实际数据按字段定义顺序排列
    for diff in row_differences:
        if isinstance(diff, dict):
            # 转换字段级差异为数组格式
            field_diffs = diff.get("detailed_field_differences", [])
            compact_field_diffs = _convert_field_differences_to_array(field_diffs)

            # 主要差异数据数组：严格按字段定义顺序
            compact_row = [
                diff.get("row_id", ""),
                diff.get("difference_type", ""),
                diff.get("row_index1", 0),
                diff.get("row_index2", 0),
                diff.get("sheet_name", ""),
                compact_field_diffs,
            ]
            compact_differences.append(compact_row)

    return compact_differences


def _convert_field_differences_to_array(
    field_diffs: list[dict[str, Any]],
) -> list[list[Any]] | None:
    """
    将字段级差异转换为数组格式

    Args:
        field_diffs: 字段差异字典列表

    Returns:
        Optional[List[List[Any]]]: 转换后的字段差异数组，如果为空则返回None

    Format:
        每个字段差异: [field_name, old_value, new_value, change_type]
    """
    if not field_diffs:
        return None

    compact_field_diffs = []
    for field_diff in field_diffs:
        if isinstance(field_diff, dict):
            # 数组格式：[field_name, old_value, new_value, change_type]
            compact_field_diffs.append(
                [
                    field_diff.get("field_name", ""),
                    field_diff.get("old_value", ""),
                    field_diff.get("new_value", ""),
                    field_diff.get("change_type", ""),
                ]
            )

    return compact_field_diffs


# --- null值清理 ---
def _deep_clean_nulls(obj: Any) -> Any:
    """
    递归深度清理对象中的null/None值

    @intention 减少响应数据体积，移除所有无效的null值、空字典、空列表

    Args:
        obj: 任意类型的对象

    Returns:
        Any: 清理后的对象，移除了所有null值和空容器

    Rules:
        - None值被移除
        - 空字典{}被移除
        - 空列表[]被移除（除了重要的字段）
        - 递归处理嵌套结构
        - 特例：紧凑数组格式保持结构完整性
        - 特例：data、matches等重要字段即使为空也保留
    """
    if isinstance(obj, dict):
        cleaned = {}
        # CellInfo语义检测：coordinate+value结构中，value=None代表空单元格，必须保留
        _is_cell_info = "coordinate" in obj
        for key, value in obj.items():
            # 单元格value字段：保留None和空字符串（空单元格有语义意义，calamine返回''而openpyxl返回None）
            if key == "value" and _is_cell_info and (value is None or value == ""):
                cleaned[key] = value  # 保留原值（包括None和''）
                continue
            if value is not None:
                cleaned_value = _deep_clean_nulls(value)
                # 保留重要字段，即使为空
                if key in ["data", "matches", "differences", "results"]:
                    cleaned[key] = cleaned_value
                # 只保留非空的值
                elif cleaned_value is not None and cleaned_value != {} and cleaned_value != []:
                    cleaned[key] = cleaned_value
        return cleaned
    elif isinstance(obj, list):
        # 检查是否是紧凑数组格式（第一个元素是字段定义列表）
        if len(obj) > 0 and isinstance(obj[0], list) and len(obj[0]) > 0 and isinstance(obj[0][0], str) and obj[0][0] in ["row_id", "field_name"]:
            # 对于紧凑数组格式，保持结构不变，只递归清理每个元素
            return [_deep_clean_nulls_preserve_structure(item) for item in obj]
        else:
            # 普通列表的null清理
            cleaned = []
            for item in obj:
                if item is not None:
                    cleaned_item = _deep_clean_nulls(item)
                    # 只保留非空的项
                    if cleaned_item is not None and cleaned_item != {} and cleaned_item != []:
                        cleaned.append(cleaned_item)
            return cleaned
    else:
        return obj


def _deep_clean_nulls_preserve_structure(obj: Any) -> Any:
    """
    在保持数组结构的前提下递归清理null值
    用于紧凑数组格式，不会改变数组长度和位置对应关系
    """
    if isinstance(obj, dict):
        cleaned = {}
        for key, value in obj.items():
            if value is not None:
                cleaned_value = _deep_clean_nulls_preserve_structure(value)
                if cleaned_value is not None and cleaned_value != {} and cleaned_value != []:
                    cleaned[key] = cleaned_value
        return cleaned
    elif isinstance(obj, list):
        # 对每个元素递归清理，但保持位置不变
        return [_deep_clean_nulls_preserve_structure(item) for item in obj]
    else:
        return obj


# ==================== 公共接口 ====================

# 向后兼容的别名
format_result = format_operation_result

# 导出的公共函数
__all__ = [
    "format_operation_result",
    "format_result",  # 向后兼容别名
]
