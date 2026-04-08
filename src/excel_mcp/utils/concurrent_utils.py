"""
并发工具模块 - REQ-032

提供线程池执行器，用于并行化 I/O 密集型批量操作。
主要优化场景：多文件读取、多工作表并行加载、批量数据准备、批量数据验证。
"""

import logging
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# 默认线程池大小：基于 CPU 核心数，但不超过 4（避免 I/O 争用）
_DEFAULT_MAX_WORKERS = min(4, (os.cpu_count() or 1))

# 批量操作分块大小：超过此阈值启用并行验证
_BATCH_CHUNK_SIZE = 500


def parallel_read_files(
    file_paths: List[str],
    read_fn: Callable[[str], Any],
    max_workers: Optional[int] = None,
) -> List[Tuple[str, Any]]:
    """并行读取多个文件。

    Args:
        file_paths: 待读取的文件路径列表
        read_fn: 单文件读取函数，接收文件路径，返回读取结果
        max_workers: 最大线程数（默认 min(4, cpu_count)）

    Returns:
        与 file_paths 顺序对应的结果列表，每项为 (file_path, result) 元组。
        读取失败的文件 result 为 None。
    """
    if not file_paths:
        return []

    workers = max_workers or _DEFAULT_MAX_WORKERS
    # 单文件无需线程池开销
    if len(file_paths) == 1:
        try:
            result = read_fn(file_paths[0])
            return [(file_paths[0], result)]
        except Exception as e:
            logger.error(f"读取文件失败 {file_paths[0]}: {e}")
            return [(file_paths[0], None)]

    results_map: Dict[str, Any] = {}
    errors: List[str] = []

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_path = {
            executor.submit(read_fn, fp): fp for fp in file_paths
        }
        for future in as_completed(future_to_path):
            fp = future_to_path[future]
            try:
                results_map[fp] = future.result()
            except Exception as e:
                logger.error(f"并行读取失败 {fp}: {e}")
                errors.append(f"{fp}: {e}")
                results_map[fp] = None

    if errors:
        logger.warning(f"并行读取完成，{len(errors)}/{len(file_paths)} 个文件失败")

    # 保持原始顺序
    return [(fp, results_map.get(fp)) for fp in file_paths]


def parallel_map(
    items: List[Any],
    process_fn: Callable[[Any], Any],
    max_workers: Optional[int] = None,
) -> List[Any]:
    """并行处理列表中的每个元素。

    Args:
        items: 待处理的项目列表
        process_fn: 处理函数，接收单个项目，返回处理结果
        max_workers: 最大线程数（默认 min(4, cpu_count)）

    Returns:
        与 items 顺序对应的结果列表。
    """
    if not items:
        return []

    workers = max_workers or _DEFAULT_MAX_WORKERS
    # 少量项目无需线程池
    if len(items) <= 2:
        return [process_fn(item) for item in items]

    results_map: Dict[int, Any] = {}

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(process_fn, item): idx
            for idx, item in enumerate(items)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results_map[idx] = future.result()
            except Exception as e:
                logger.error(f"并行处理失败（索引 {idx}）: {e}")
                results_map[idx] = None

    return [results_map[i] for i in range(len(items))]


def parallel_validate_batch_data(
    rows: List[Dict[str, Any]],
    validate_fn: Callable[[Dict[str, Any]], Optional[str]],
    max_workers: Optional[int] = None,
) -> List[Optional[str]]:
    """并行验证批量数据行，返回每行的错误信息（None表示通过）。

    Args:
        rows: 待验证的行数据列表，每行为字典
        validate_fn: 单行验证函数，接收字典，返回错误信息字符串（None表示通过）
        max_workers: 最大线程数（默认 min(4, cpu_count)）

    Returns:
        与 rows 顺序对应的错误信息列表，每项为错误字符串或 None。
    """
    if not rows:
        return []

    workers = max_workers or _DEFAULT_MAX_WORKERS
    if len(rows) <= _BATCH_CHUNK_SIZE:
        return [validate_fn(row) for row in rows]

    results_map: Dict[int, Optional[str]] = {}

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(validate_fn, row): idx
            for idx, row in enumerate(rows)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results_map[idx] = future.result()
            except Exception as e:
                logger.error(f"并行验证失败（索引 {idx}）: {e}")
                results_map[idx] = f"验证异常: {e}"

    return [results_map[i] for i in range(len(rows))]


def parallel_group_execute(
    groups: Dict[str, List[Any]],
    process_fn: Callable[[str, List[Any]], Any],
    max_workers: Optional[int] = None,
) -> Dict[str, Any]:
    """按分组并行执行操作，每个分组一个线程。

    Args:
        groups: 分组数据字典，key 为分组名，value 为该组数据列表
        process_fn: 分组处理函数，接收 (group_key, group_data)，返回处理结果
        max_workers: 最大线程数（默认 min(4, cpu_count)）

    Returns:
        与 groups key 对应的结果字典。
    """
    if not groups:
        return {}

    workers = max_workers or _DEFAULT_MAX_WORKERS
    if len(groups) <= 2:
        return {key: process_fn(key, data) for key, data in groups.items()}

    results: Dict[str, Any] = {}

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_key = {
            executor.submit(process_fn, key, data): key
            for key, data in groups.items()
        }
        for future in as_completed(future_to_key):
            key = future_to_key[future]
            try:
                results[key] = future.result()
            except Exception as e:
                logger.error(f"分组 '{key}' 并行执行失败: {e}")
                results[key] = {"error": str(e)}

    return results
