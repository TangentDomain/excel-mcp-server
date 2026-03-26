#!/usr/bin/env python3
"""
Excel MCP Server - 基于 FastMCP 和 openpyxl 实现

重构后的服务器文件，只包含MCP接口定义，具体实现委托给核心模块

主要功能：
1. 正则搜索：在Excel文件中搜索符合正则表达式的单元格
2. 范围获取：读取指定范围的Excel数据
3. 范围修改：修改指定范围的Excel数据
4. 工作表管理：创建、删除、重命名工作表
5. 行列操作：插入、删除行列

技术栈：
- FastMCP: 用于MCP服务器框架
- openpyxl: 用于Excel文件操作
"""

import functools
import glob
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
from collections import defaultdict
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional, List, Dict, Any, Union

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    print(f"Error: 缺少必要的依赖包: {e}")
    print("请运行: pip install mcp openpyxl")
    exit(1)

# 导入API模块
from .api.excel_operations import ExcelOperations
from .utils.validators import ExcelValidator, DataValidationError

# ==================== 操作日志系统 ====================
class OperationLogger:
    """操作日志记录器，用于跟踪所有Excel操作"""

    def __init__(self):
        self.log_file = None
        self.current_session = []

    def start_session(self, file_path: str):
        """开始新的操作会话"""
        self.log_file = os.path.join(
            os.path.dirname(file_path),
            ".excel_mcp_logs",
            f"operations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )

        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)

        self.current_session = [{
            'session_id': datetime.now().isoformat(),
            'file_path': file_path,
            'operations': []
        }]

        self._save_log()

    def log_operation(self, operation: str, details: Dict[str, Any]):
        """记录操作"""
        if not self.current_session:
            return

        operation_record = {
            'timestamp': datetime.now().isoformat(),
            'operation': operation,
            'details': details
        }

        self.current_session[0]['operations'].append(operation_record)
        self._save_log()

    def _save_log(self):
        """保存日志到文件"""
        if not self.log_file:
            return

        try:
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump(self.current_session, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"保存操作日志失败: {e}")

    def get_recent_operations(self, limit: int = 10) -> List[Dict[str, Any]]:
        """获取最近的操作记录"""
        if not self.current_session:
            return []

        operations = self.current_session[0]['operations']
        return operations[-limit:] if len(operations) > limit else operations

# ==================== 工具调用追踪器 ====================
class ToolCallTracker:
    """工具调用追踪器，记录每个工具的调用次数、耗时和错误率"""

    def __init__(self):
        self._stats: Dict[str, Dict[str, Any]] = defaultdict(lambda: {
            'call_count': 0,
            'total_time_ms': 0.0,
            'error_count': 0,
            'last_called': None,
            'min_time_ms': float('inf'),
            'max_time_ms': 0.0,
        })
        self._lock = threading.Lock()
        self._start_time = datetime.now()

    def record(self, tool_name: str, duration_ms: float, success: bool = True):
        """记录一次工具调用"""
        with self._lock:
            s = self._stats[tool_name]
            s['call_count'] += 1
            s['total_time_ms'] += duration_ms
            s['last_called'] = datetime.now().isoformat()
            if duration_ms < s['min_time_ms']:
                s['min_time_ms'] = duration_ms
            if duration_ms > s['max_time_ms']:
                s['max_time_ms'] = duration_ms
            if not success:
                s['error_count'] += 1

    def get_stats(self) -> Dict[str, Any]:
        """获取所有工具的调用统计，按调用次数降序排列"""
        with self._lock:
            tools = {}
            for name, s in sorted(self._stats.items(), key=lambda x: -x[1]['call_count']):
                tools[name] = {
                    'call_count': s['call_count'],
                    'avg_time_ms': round(s['total_time_ms'] / s['call_count'], 1) if s['call_count'] else 0,
                    'min_time_ms': round(s['min_time_ms'], 1) if s['min_time_ms'] != float('inf') else 0,
                    'max_time_ms': round(s['max_time_ms'], 1),
                    'error_count': s['error_count'],
                    'last_called': s['last_called'],
                }
            return {
                'uptime_seconds': round((datetime.now() - self._start_time).total_seconds(), 0),
                'total_calls': sum(s['call_count'] for s in self._stats.values()),
                'total_errors': sum(s['error_count'] for s in self._stats.values()),
                'tools': tools,
            }

    def reset(self):
        """重置所有统计"""
        with self._lock:
            self._stats.clear()
            self._start_time = datetime.now()


_tracker = ToolCallTracker()


def _track_call(func):
    """工具调用追踪装饰器，记录每次调用的耗时和结果"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            duration_ms = (time.perf_counter() - start) * 1000
            _tracker.record(func.__name__, duration_ms, success=True)
            logger.debug(f"[TOOL] {func.__name__}: {duration_ms:.1f}ms",
                         extra={'tool': func.__name__, 'duration_ms': round(duration_ms, 1)})
            return result
        except Exception as e:
            duration_ms = (time.perf_counter() - start) * 1000
            _tracker.record(func.__name__, duration_ms, success=False)
            logger.debug(f"[TOOL] {func.__name__}: {duration_ms:.1f}ms ERROR: {e}",
                         extra={'tool': func.__name__, 'duration_ms': round(duration_ms, 1), 'error': str(e)})
            raise
    return wrapper


# ==================== 安全验证模块 ====================
class SecurityValidator:
    """文件路径安全验证，防止路径穿越和资源耗尽"""

    # 允许的文件扩展名
    ALLOWED_EXTENSIONS = {'.xlsx', '.xlsm', '.xls', '.csv', '.json', '.bak'}
    # 文件大小上限（50MB）
    MAX_FILE_SIZE = 50 * 1024 * 1024
    # 危险公式模式（DDE攻击、命令执行等）
    DANGEROUS_FORMULA_PATTERNS = [
        re.compile(r'\bDDE\b', re.IGNORECASE),
        re.compile(r'\bCMD\b', re.IGNORECASE),
        re.compile(r'\bSHELL\b', re.IGNORECASE),
        re.compile(r'\bREGISTER\b', re.IGNORECASE),
        re.compile(r'\|.*\|', re.IGNORECASE),  # DDE链接格式
        re.compile(r'@SUM', re.IGNORECASE),  # Lotus 1-2-3 兼容性攻击
    ]

    @classmethod
    def validate_file_path(cls, file_path: str) -> Dict[str, Any]:
        """验证文件路径安全性，返回 {'valid': bool, 'error': str|None}"""
        if not file_path:
            return {'valid': False, 'error': '文件路径不能为空'}

        # 解析为绝对路径并规范化
        try:
            resolved = Path(file_path).resolve()
        except (OSError, ValueError) as e:
            return {'valid': False, 'error': f'无效的文件路径: {e}'}

        # 路径穿越检测：规范化后路径必须以 / 开头（绝对路径）且不含 ..
        if '..' in Path(file_path).parts:
            return {'valid': False, 'error': '文件路径不允许包含 ".." 路径穿越'}

        # 拒绝符号链接（用原始路径检查，resolve会解析符号链接）
        try:
            if Path(file_path).is_symlink():
                return {'valid': False, 'error': '不允许通过符号链接访问文件'}
        except OSError:
            pass

        # 拒绝隐藏文件（以 . 开头）
        basename = os.path.basename(file_path)
        if basename.startswith('.') and not basename.endswith('.bak'):
            return {'valid': False, 'error': f'不允许访问隐藏文件: {basename}'}

        # 扩展名检查（对有扩展名的文件检查）
        _, ext = os.path.splitext(file_path)
        if ext and ext.lower() not in cls.ALLOWED_EXTENSIONS:
            return {'valid': False, 'error': f'不支持的文件格式: {ext}（允许: {", ".join(sorted(cls.ALLOWED_EXTENSIONS))}）'}

        return {'valid': True, 'error': None}

    @classmethod
    def validate_file_size(cls, file_path: str) -> Dict[str, Any]:
        """验证文件大小是否在允许范围内"""
        try:
            size = os.path.getsize(file_path)
            if size > cls.MAX_FILE_SIZE:
                return {'valid': False, 'error': f'文件过大: {size / 1024 / 1024:.1f}MB（上限{cls.MAX_FILE_SIZE / 1024 / 1024:.0f}MB）'}
            return {'valid': True, 'error': None}
        except OSError as e:
            return {'valid': False, 'error': f'无法获取文件大小: {e}'}

    @classmethod
    def validate_formula(cls, formula: str) -> Dict[str, Any]:
        """验证公式安全性，检测危险模式"""
        for pattern in cls.DANGEROUS_FORMULA_PATTERNS:
            if pattern.search(formula):
                return {'valid': False, 'error': f'检测到危险公式模式，公式中不允许包含 {pattern.pattern}'}
        return {'valid': True, 'error': None}

    @classmethod
    def cleanup_orphan_temp_files(cls, temp_dir: str = None) -> int:
        """清理孤儿临时文件（.xlsx.bak），返回清理数量"""
        target = temp_dir or tempfile.gettempdir()
        pattern = os.path.join(target, '*.xlsx.bak')
        cleaned = 0
        for f in glob.glob(pattern):
            try:
                # 超过1小时的临时备份文件自动清理
                age = time.time() - os.path.getmtime(f)
                if age > 3600:
                    os.remove(f)
                    cleaned += 1
            except OSError:
                pass
        return cleaned


def _validate_path(file_path: str) -> Optional[Dict[str, Any]]:
    """统一的文件路径安全验证入口，失败返回错误dict，通过返回None"""
    result = SecurityValidator.validate_file_path(file_path)
    if not result['valid']:
        return {'success': False, 'message': f'🔒 安全验证失败: {result["error"]}'}
    # 文件存在时额外检查大小
    if os.path.exists(file_path):
        size_result = SecurityValidator.validate_file_size(file_path)
        if not size_result['valid']:
            return {'success': False, 'message': f'🔒 安全验证失败: {size_result["error"]}'}
    return None


# 全局操作日志器
operation_logger = OperationLogger()

# ==================== 结构化JSON日志 ====================
class JsonFormatter(logging.Formatter):
    """结构化JSON日志格式化器，每条日志输出为单行JSON对象
    
    输出字段: ts(ISO时间戳), level, module, message
    可选字段: tool(工具名), duration_ms(耗时), error(错误信息), file_path, operation
    激活方式: EXCEL_MCP_JSON_LOG=1
    """
    def format(self, record):
        log_entry = {
            'ts': datetime.now().isoformat(),
            'level': record.levelname,
            'module': record.module,
            'message': record.getMessage(),
        }
        for field in ('tool', 'duration_ms', 'error', 'file_path', 'operation'):
            value = getattr(record, field, None)
            if value is not None:
                log_entry[field] = value
        return json.dumps(log_entry, ensure_ascii=False)


# ==================== 配置和初始化 ====================
# 日志级别: 默认WARNING，设置EXCEL_MCP_DEBUG=1开启DEBUG
_log_level = logging.DEBUG if os.environ.get('EXCEL_MCP_DEBUG') else logging.WARNING
_json_log = os.environ.get('EXCEL_MCP_JSON_LOG')
if _json_log:
    _handler = logging.StreamHandler()
    _handler.setFormatter(JsonFormatter())
    logging.basicConfig(level=_log_level, handlers=[_handler], force=True)
else:
    logging.basicConfig(
        level=_log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
        ],
        force=True,
    )
logger = logging.getLogger(__name__)

# 启动时清理孤儿临时文件
_cleaned = SecurityValidator.cleanup_orphan_temp_files()
if _cleaned:
    logger.info(f"启动清理: 删除 {_cleaned} 个孤儿临时文件")

# 创建FastMCP服务器实例，开启调试模式和详细日志
mcp = FastMCP(
    name="excel-mcp",
    instructions=r"""🎮 游戏开发Excel配置表管理专家

## 🔥 核心原则：SQL优先

**优先使用 `excel_query`** - 所有数据查询分析任务
- 复杂条件筛选 ✅ WHERE, LIKE, IN, BETWEEN, 子查询
- 聚合统计分析 ✅ COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 高级查询 ✅ CASE WHEN, CTE(WITH), EXISTS, JOIN
- 排序限制 ✅ ORDER BY, LIMIT, OFFSET
- 字符串函数 ✅ UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING

## 📊 工具选择决策树
```
需要数据分析/查询？   → excel_query (SQL引擎，支持WHERE/GROUP BY/JOIN/子查询/CTE)
需要快速了解表结构？ → excel_describe_table (列名+类型+样本)
需要定位单元格？     → excel_search (返回row/column)
需要批量修改数据？   → excel_update_query (SQL UPDATE语法)
需要修改指定单元格？ → excel_update_range (范围写入)
需要格式调整？       → excel_format_cells
```

## ✅ SQL已支持功能 (35项)
基础查询: SELECT, DISTINCT, 别名(AS), 数学表达式(+-*/%)
条件筛选: WHERE, =/>/</<=/>=/!=, LIKE, IN, NOT IN, BETWEEN, AND/OR, IS NULL, NOT
高级查询: 子查询(WHERE col IN (SELECT...)), CASE WHEN, COALESCE, EXISTS, CTE(WITH)
聚合统计: COUNT(*), COUNT(col), SUM, AVG, MAX, MIN, GROUP BY, HAVING, TOTAL行
排序限制: ORDER BY DESC/ASC, LIMIT, OFFSET
表关联: INNER JOIN, LEFT JOIN（同文件内工作表）
字符串函数: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT

## ❌ SQL不支持功能
UNION, 窗口函数(ROW_NUMBER/RANK), INSERT/DELETE, RIGHT JOIN, CROSS JOIN
替代方案: 子查询用 WHERE col IN (SELECT...)，不支持FROM子查询

## ⚠️ 重要原则
- 双行表头自动识别: 第1行中文描述+第2行英文字段名，可用中英文列名查询
- 1-based索引: 第1行=1, 第1列=1
- 范围格式: 必须包含工作表名 "技能表!A1:Z100"
- 默认覆盖: update_range默认覆盖，需保留数据用insert_mode=True

## 🎮 游戏配置表示例
技能统计: SELECT 技能类型, AVG(伤害), COUNT(*) FROM 技能表 GROUP BY 技能类型
高级筛选: SELECT * FROM 技能表 WHERE 伤害 > (SELECT AVG(伤害) FROM 技能表)
条件表达式: SELECT 技能名, CASE WHEN 伤害>100 THEN '高' ELSE '低' END AS 等级 FROM 技能表
CTE: WITH 高伤 AS (SELECT * FROM 技能表 WHERE 伤害>100) SELECT COUNT(*) FROM 高伤

## ⚡ 常用流程
1. excel_list_sheets - 列出工作表
2. excel_describe_table - 快速了解表结构（列名+类型+样本）
3. excel_query - SQL查询分析
4. excel_update_query / excel_update_range - 数据更新
5. excel_compare_sheets - 版本对比
""",
    debug=bool(os.environ.get('EXCEL_MCP_DEBUG')),
    log_level="DEBUG" if os.environ.get('EXCEL_MCP_DEBUG') else "WARNING"
)


# ==================== MCP 工具定义 ====================

@mcp.tool()
@_track_call
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """
列出Excel文件中所有工作表名称。查询前先用此工具确认有哪些工作表。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.list_sheets(file_path)


@mcp.tool()
@_track_call
def excel_get_sheet_headers(file_path: str) -> Dict[str, Any]:
    """
批量获取所有工作表的双行表头（游戏开发专用）。
返回每个表的字段描述（中文）和字段名（英文）。专为游戏配置表双行表头设计。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.get_sheet_headers(file_path)


@mcp.tool()
@_track_call
def excel_search(
    file_path: str,
    pattern: str,
    sheet_name: Optional[str] = None,
    case_sensitive: bool = False,
    whole_word: bool = False,
    use_regex: bool = False,
    include_values: bool = True,
    include_formulas: bool = False,
    range: Optional[str] = None
) -> Dict[str, Any]:
    """
在单个Excel文件中搜索文本。返回匹配的单元格位置(row/column)和值。支持正则。
搜索整个目录的多个文件请用excel_search_directory。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.search(file_path, pattern, sheet_name, case_sensitive, whole_word, use_regex, include_values, include_formulas, range)


@mcp.tool()
@_track_call
def excel_search_directory(
    directory_path: str,
    pattern: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    use_regex: bool = False,
    include_values: bool = True,
    include_formulas: bool = False,
    recursive: bool = True,
    file_extensions: Optional[List[str]] = None,
    file_pattern: Optional[str] = None,
    max_files: int = 100
) -> Dict[str, Any]:
    """
在目录下所有Excel文件中批量搜索。返回文件名+单元格位置+值。支持正则、大小写、全词匹配、文件名过滤。
搜索单个文件请用excel_search。
    """
    _path_err = _validate_path(directory_path)
    if _path_err:
        return _path_err
    return ExcelOperations.search_directory(directory_path, pattern, case_sensitive, whole_word, use_regex, include_values, include_formulas, recursive, file_extensions, file_pattern, max_files)


@mcp.tool()
@_track_call
def excel_get_range(
    file_path: str,
    range: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
读取指定范围的原始单元格数据。适合精确读取已知区域（如"Sheet1!A1:C10"）。
如需查询/筛选/聚合，优先使用excel_query。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 增强参数验证

    try:
        # 验证范围表达式格式
        range_validation = ExcelValidator.validate_range_expression(range)

        # 验证操作规模
        scale_validation = ExcelValidator.validate_operation_scale(range_validation['range_info'])

        # 记录验证成功到调试日志
        logger.debug(f"范围验证成功: {range_validation['normalized_range']}")

    except DataValidationError as e:
        # 记录验证失败
        logger.error(f"范围验证失败: {str(e)}")

        return {
            'success': False,
            'error': 'VALIDATION_FAILED',
            'message': f"范围表达式验证失败: {str(e)}"
        }

    # 调用原始函数
    result = ExcelOperations.get_range(file_path, range, include_formatting)

    # 如果成功，添加验证信息到结果中
    if result.get('success'):
        result['validation_info'] = {
            'normalized_range': range_validation['normalized_range'],
            'sheet_name': range_validation['sheet_name'],
            'range_type': range_validation['range_info']['type'],
            'scale_assessment': scale_validation
        }

    return result


@mcp.tool()
@_track_call
def excel_get_headers(
    file_path: str,
    sheet_name: str,
    header_row: int = 1,
    max_columns: Optional[int] = None
) -> Dict[str, Any]:
    """
获取单个工作表的表头行。返回指定行的列名列表。
查看所有工作表的双行表头（中英映射）请用excel_get_sheet_headers。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.get_headers(file_path, sheet_name, header_row, max_columns)


@mcp.tool()
@_track_call
def excel_update_range(
    file_path: str,
    range: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = True
) -> Dict[str, Any]:
    """
写入数据到指定范围。适合批量写入已知数据（二维数组）。
如需条件修改（如"火系伤害+10%"），优先使用excel_update_query。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err

    try:
        # 验证范围表达式格式
        range_validation = ExcelValidator.validate_range_expression(range)

        # 验证操作规模
        scale_validation = ExcelValidator.validate_operation_scale(range_validation['range_info'])

        # 如果有警告信息，记录到操作日志
        if scale_validation.get('warning'):
            logger.warning(f"操作规模警告: {scale_validation['warning']}")

    except DataValidationError as e:
        # 记录验证失败
        operation_logger.start_session(file_path)
        operation_logger.log_operation("validation_failed", {
            "operation": "update_range",
            "range": range,
            "error": str(e)
        })

        return {
            'success': False,
            'error': 'VALIDATION_FAILED',
            'message': f"参数验证失败: {str(e)}"
        }

    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录操作日志
    operation_logger.log_operation("update_range", {
        "range": range,
        "validated_range": range_validation['normalized_range'],
        "data_rows": len(data),
        "insert_mode": insert_mode,
        "preserve_formulas": preserve_formulas,
        "scale_info": scale_validation
    })

    try:
        result = ExcelOperations.update_range(file_path, range, data, preserve_formulas, insert_mode)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "updated_cells": result.get('updated_cells', 0),
            "message": result.get('message', '')
        })

        return result

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"更新操作失败: {str(e)}"
        })

        return {
            'success': False,
            'error': 'OPERATION_FAILED',
            'message': f"更新操作失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_preview_operation(
    file_path: str,
    range: str,
    operation_type: str = "update",
    data: Optional[List[List[Any]]] = None
) -> Dict[str, Any]:
    """
预览操作影响范围和当前数据，不实际执行。修改前确认用。
全面评估请用excel_assess_data_impact。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 读取当前数据
    current_data = ExcelOperations.get_range(file_path, range)

    if not current_data.get('success'):
        return {
            'success': False,
            'error': 'PREVIEW_FAILED',
            'message': f"无法预览操作: {current_data.get('message', '未知错误')}"
        }

    # 分析影响
    data_rows = len(current_data.get('data', []))
    data_cols = len(current_data.get('data', [])) if data_rows > 0 else 0
    total_cells = data_rows * data_cols

    # 检查是否包含非空数据
    has_data = any(
        any(cell is not None and str(cell).strip() for cell in row)
        for row in current_data.get('data', [])
    )

    # 安全评估
    risk_level = "LOW"
    if has_data:
        if total_cells > 100:
            risk_level = "HIGH"
        elif total_cells > 20:
            risk_level = "MEDIUM"
        else:
            risk_level = "LOW"

    return {
        'success': True,
        'operation_type': operation_type,
        'range': range,
        'current_data': current_data.get('data', []),
        'impact_assessment': {
            'rows_affected': data_rows,
            'columns_affected': data_cols,
            'total_cells': total_cells,
            'has_existing_data': has_data,
            'risk_level': risk_level
        },
        'recommendations': _get_safety_recommendations(operation_type, has_data, risk_level),
        'safety_warning': _generate_safety_warning(operation_type, has_data, risk_level)
    }


def _get_safety_recommendations(operation_type: str, has_data: bool, risk_level: str) -> List[str]:
    """获取安全操作建议"""
    recommendations = []

    if operation_type == "update":
        if has_data:
            recommendations.append("⚠️ 范围内已有数据，建议使用 insert_mode=True")
            if risk_level == "HIGH":
                recommendations.append("🔴 大范围数据操作，强烈建议先备份")
            recommendations.append("📊 建议先预览完整数据再操作")
        else:
            recommendations.append("✅ 范围为空，可以安全操作")

    elif operation_type == "delete":
        recommendations.append("🗑️ 删除操作不可逆，请确认")
        if has_data:
            recommendations.append("⚠️ 将删除现有数据，请仔细检查")

    return recommendations


def _generate_safety_warning(operation_type: str, has_data: bool, risk_level: str) -> str:
    """生成安全警告"""
    if risk_level == "HIGH":
        return f"🔴 高风险警告: {operation_type}操作将影响大量数据，请谨慎操作"
    elif risk_level == "MEDIUM":
        return f"🟡 中等风险: {operation_type}操作将影响部分数据，建议先备份"
    else:
        return f"✅ 低风险: {operation_type}操作影响较小，可以安全执行"


@mcp.tool()
@_track_call
def excel_assess_data_impact(
    file_path: str,
    range: str,
    operation_type: str = "update",
    data: Optional[List[List[Any]]] = None
) -> Dict[str, Any]:
    """
全面评估操作对数据的潜在影响（风险等级、数据类型分析、结果预测）。
简单预览请用excel_preview_operation。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err

    try:
        # 验证范围表达式
        range_validation = ExcelValidator.validate_range_expression(range)
        range_info = range_validation['range_info']

        # 获取当前数据
        current_data_result = ExcelOperations.get_range(file_path, range)

        if not current_data_result.get('success'):
            return {
                'success': False,
                'error': 'DATA_RETRIEVAL_FAILED',
                'message': f"无法获取当前数据: {current_data_result.get('message', '未知错误')}"
            }

        current_data = current_data_result.get('data', [])

        # 分析当前数据内容
        data_analysis = _analyze_current_data(current_data)

        # 计算操作规模
        scale_info = ExcelValidator.validate_operation_scale(range_info)

        # 评估操作风险
        risk_assessment = _assess_operation_risk(
            operation_type,
            data_analysis,
            scale_info,
            data
        )

        # 生成建议
        recommendations = _generate_safety_recommendations(
            operation_type,
            data_analysis,
            risk_assessment,
            scale_info
        )

        # 预测结果
        prediction = _predict_operation_result(
            operation_type,
            current_data,
            data,
            scale_info
        )

        return {
            'success': True,
            'operation_type': operation_type,
            'range': range,
            'validation_info': range_validation,
            'current_data_analysis': data_analysis,
            'scale_assessment': scale_info,
            'risk_assessment': risk_assessment,
            'safety_recommendations': recommendations,
            'result_prediction': prediction,
            'impact_summary': {
                'total_cells': scale_info['total_cells'],
                'non_empty_cells': data_analysis['non_empty_cell_count'],
                'data_type_distribution': data_analysis['data_types'],
                'potential_data_loss': data_analysis['non_empty_cell_count'] if operation_type in ['delete', 'update'] else 0,
                'overall_risk_level': risk_assessment['overall_risk']
            }
        }

    except DataValidationError as e:
        return {
            'success': False,
            'error': 'VALIDATION_FAILED',
            'message': f"参数验证失败: {str(e)}"
        }

    except Exception as e:
        return {
            'success': False,
            'error': 'ASSESSMENT_FAILED',
            'message': f"数据影响评估失败: {str(e)}"
        }


def _analyze_current_data(data: List[List[Any]]) -> Dict[str, Any]:
    """分析当前数据内容"""
    if not data:
        return {
            'row_count': 0,
            'column_count': 0,
            'total_cells': 0,
            'non_empty_cell_count': 0,
            'empty_cell_count': 0,
            'data_types': {},
            'has_formulas': False,
            'has_numeric_data': False,
            'has_text_data': False,
            'has_dates': False,
            'completeness_rate': 0.0
        }

    total_cells = len(data) * max(len(row) for row in data) if data else 0
    non_empty_cells = 0
    data_types = {}
    has_formulas = False
    has_numeric_data = False
    has_text_data = False
    has_dates = False

    for row in data:
        for cell in row:
            if cell is not None and str(cell).strip():
                non_empty_cells += 1

                # 分析数据类型
                if isinstance(cell, str):
                    if cell.startswith('='):
                        has_formulas = True
                        data_types['formulas'] = data_types.get('formulas', 0) + 1
                    else:
                        has_text_data = True
                        data_types['text'] = data_types.get('text', 0) + 1
                elif isinstance(cell, (int, float)):
                    has_numeric_data = True
                    data_types['numeric'] = data_types.get('numeric', 0) + 1
                elif isinstance(cell, datetime):
                    has_dates = True
                    data_types['dates'] = data_types.get('dates', 0) + 1
                else:
                    data_types['other'] = data_types.get('other', 0) + 1

    return {
        'row_count': len(data),
        'column_count': max(len(row) for row in data) if data else 0,
        'total_cells': total_cells,
        'non_empty_cell_count': non_empty_cells,
        'empty_cell_count': total_cells - non_empty_cells,
        'data_types': data_types,
        'has_formulas': has_formulas,
        'has_numeric_data': has_numeric_data,
        'has_text_data': has_text_data,
        'has_dates': has_dates,
        'completeness_rate': (non_empty_cells / total_cells * 100) if total_cells > 0 else 0.0
    }


def _assess_operation_risk(
    operation_type: str,
    data_analysis: Dict[str, Any],
    scale_info: Dict[str, Any],
    new_data: Optional[List[List[Any]]] = None
) -> Dict[str, Any]:
    """评估操作风险"""
    risk_factors = []
    risk_score = 0

    # 基于操作类型的风险
    if operation_type == "delete":
        risk_factors.append("删除操作不可逆")
        risk_score += 30
    elif operation_type == "update":
        if data_analysis['non_empty_cell_count'] > 0:
            risk_factors.append("将覆盖现有数据")
            risk_score += 20
    elif operation_type == "format":
        risk_factors.append("格式化操作")
        risk_score += 10

    # 基于数据量的风险
    if scale_info['total_cells'] > 10000:
        risk_factors.append("大范围操作")
        risk_score += 25
    elif scale_info['total_cells'] > 1000:
        risk_factors.append("中等范围操作")
        risk_score += 15

    # 基于数据内容的风险
    if data_analysis['has_formulas']:
        risk_factors.append("包含公式数据")
        risk_score += 15

    if data_analysis['completeness_rate'] > 80:
        risk_factors.append("高密度数据区域")
        risk_score += 10

    # 确定整体风险等级
    if risk_score >= 60:
        overall_risk = "HIGH"
    elif risk_score >= 30:
        overall_risk = "MEDIUM"
    else:
        overall_risk = "LOW"

    return {
        'risk_score': risk_score,
        'overall_risk': overall_risk,
        'risk_factors': risk_factors,
        'requires_backup': overall_risk in ["HIGH", "MEDIUM"],
        'requires_confirmation': overall_risk == "HIGH"
    }


def _generate_safety_recommendations(
    operation_type: str,
    data_analysis: Dict[str, Any],
    risk_assessment: Dict[str, Any],
    scale_info: Dict[str, Any]
) -> List[str]:
    """生成安全建议"""
    recommendations = []

    # 基础建议
    if risk_assessment['requires_backup']:
        recommendations.append("🔴 强烈建议在操作前创建备份")

    if risk_assessment['requires_confirmation']:
        recommendations.append("⚠️ 高风险操作，请仔细确认后再执行")

    # 基于数据内容的建议
    if data_analysis['has_formulas']:
        recommendations.append("📊 检测到公式数据，建议验证公式的正确性")

    if data_analysis['completeness_rate'] > 50:
        recommendations.append("💾 数据密度较高，建议先导出重要数据")

    # 基于操作类型的建议
    if operation_type == "delete":
        recommendations.append("🗑️ 删除操作不可逆，请确认数据不再需要")
    elif operation_type == "update":
        if data_analysis['non_empty_cell_count'] > 0:
            recommendations.append("✏️ 将覆盖现有数据，建议使用insert_mode=True")

    # 性能建议
    if scale_info['total_cells'] > 5000:
        recommendations.append("⏱️ 大范围操作可能需要较长时间，请耐心等待")

    return recommendations


def _predict_operation_result(
    operation_type: str,
    current_data: List[List[Any]],
    new_data: Optional[List[List[Any]]],
    scale_info: Dict[str, Any]
) -> Dict[str, Any]:
    """预测操作结果"""
    prediction = {
        'affected_cells': scale_info['total_cells'],
        'data_overwrite_count': 0,
        'data_insert_count': 0,
        'estimated_time': "minimal"
    }

    if operation_type == "update" and new_data:
        prediction['data_overwrite_count'] = len([cell for row in current_data for cell in row if cell is not None])
        prediction['data_insert_count'] = len([cell for row in new_data for cell in row if cell is not None])
    elif operation_type == "delete":
        prediction['data_overwrite_count'] = len([cell for row in current_data for cell in row if cell is not None])

    # 估算执行时间
    if scale_info['total_cells'] > 10000:
        prediction['estimated_time'] = "long"
    elif scale_info['total_cells'] > 1000:
        prediction['estimated_time'] = "medium"

    return prediction


@mcp.tool()
@_track_call
def excel_get_operation_history(
    file_path: Optional[str] = None,
    limit: int = 20
) -> Dict[str, Any]:
    """
获取操作历史记录。可按文件过滤。file_path为空时返回所有记录。
    """
    if file_path is not None:
        _path_err = _validate_path(file_path)
        if _path_err:
            return _path_err
    try:
        recent_operations = operation_logger.get_recent_operations(limit)

        # 如果指定了文件路径，过滤操作
        if file_path:
            recent_operations = [
                op for op in recent_operations
                if op.get('details', {}).get('file_path') == file_path
            ]

        # 统计信息
        total_operations = len(recent_operations)
        operation_types = {}
        for op in recent_operations:
            op_type = op.get('operation', 'unknown')
            operation_types[op_type] = operation_types.get(op_type, 0) + 1

        # 统计成功/失败
        success_count = sum(1 for op in recent_operations
                          if op.get('operation') == 'operation_result' and
                          op.get('details', {}).get('success', False))

        error_count = sum(1 for op in recent_operations
                        if op.get('operation') == 'operation_error')

        return {
            'success': True,
            'file_path': file_path,
            'operations': recent_operations,
            'statistics': {
                'total_operations': total_operations,
                'operation_types': operation_types,
                'success_count': success_count,
                'error_count': error_count,
                'success_rate': f"{(success_count / (success_count + error_count) * 100):.1f}%" if (success_count + error_count) > 0 else "0%"
            },
            'message': f"找到 {total_operations} 条操作记录"
        }

    except Exception as e:
        return {
            'success': False,
            'error': 'HISTORY_RETRIEVAL_FAILED',
            'message': f"获取操作历史失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_create_backup(
    file_path: str,
    backup_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
创建Excel文件备份。默认存入同目录.backup/文件夹。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    if not os.path.exists(file_path):
        return {
            'success': False,
            'error': 'FILE_NOT_FOUND',
            'message': f"源文件不存在: {file_path}"
        }

    try:
        # 创建备份目录
        if backup_dir is None:
            base_dir = os.path.dirname(file_path)
            backup_dir = os.path.join(base_dir, ".excel_mcp_backups")

        os.makedirs(backup_dir, exist_ok=True)

        # 生成备份文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        backup_filename = f"{name}_backup_{timestamp}{ext}"
        backup_path = os.path.join(backup_dir, backup_filename)

        # 创建备份
        shutil.copy2(file_path, backup_path)

        # 检查备份大小
        original_size = os.path.getsize(file_path)
        backup_size = os.path.getsize(backup_path)

        return {
            'success': True,
            'original_file': file_path,
            'backup_file': backup_path,
            'backup_directory': backup_dir,
            'file_size': {
                'original': original_size,
                'backup': backup_size
            },
            'timestamp': timestamp,
            'message': f"备份创建成功: {backup_filename}"
        }

    except Exception as e:
        return {
            'success': False,
            'error': 'BACKUP_FAILED',
            'message': f"备份创建失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_restore_backup(
    backup_path: str,
    target_path: Optional[str] = None
) -> Dict[str, Any]:
    """
从备份恢复Excel文件。可恢复到原位置或指定位置。
    """
    _path_err = _validate_path(backup_path)
    if _path_err:
        return _path_err

    if not os.path.exists(backup_path):
        return {
            'success': False,
            'error': 'BACKUP_NOT_FOUND',
            'message': f"备份文件不存在: {backup_path}"
        }

    try:
        # 确定目标路径
        if target_path is None:
            # 尝试从备份文件名推断原始文件名
            filename = os.path.basename(backup_path)
            if "_backup_" in filename:
                # 移除备份时间戳
                parts = filename.split("_backup_")
                target_path = parts[0] + os.path.splitext(backup_path)[1]
            else:
                target_path = filename.replace("_backup_", ".")

        # 创建目标目录
        target_dir = os.path.dirname(target_path)
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)

        # 检查目标文件是否存在
        target_exists = os.path.exists(target_path)

        # 执行恢复
        shutil.copy2(backup_path, target_path)

        return {
            'success': True,
            'backup_file': backup_path,
            'target_file': target_path,
            'target_existed': target_exists,
            'message': f"文件恢复成功: {os.path.basename(target_path)}"
        }

    except Exception as e:
        return {
            'success': False,
            'error': 'RESTORE_FAILED',
            'message': f"恢复失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_list_backups(
    file_path: str,
    backup_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
列出指定文件的所有备份版本。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    try:
        # 确定备份目录
        if backup_dir is None:
            base_dir = os.path.dirname(file_path)
            backup_dir = os.path.join(base_dir, ".excel_mcp_backups")

        if not os.path.exists(backup_dir):
            return {
                'success': True,
                'backups': [],
                'message': "备份目录不存在"
            }

        # 获取文件名
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        backup_pattern = f"{name}_backup_*{ext}"

        # 查找备份文件
        backup_files = []
        for file in os.listdir(backup_dir):
            if file.startswith(f"{name}_backup_") and file.endswith(ext):
                full_path = os.path.join(backup_dir, file)
                stat = os.stat(full_path)
                backup_files.append({
                    'filename': file,
                    'path': full_path,
                    'size': stat.st_size,
                    'created_time': datetime.fromtimestamp(stat.st_ctime),
                    'modified_time': datetime.fromtimestamp(stat.st_mtime)
                })

        # 按时间排序
        backup_files.sort(key=lambda x: x['created_time'], reverse=True)

        return {
            'success': True,
            'original_file': file_path,
            'backup_directory': backup_dir,
            'backups': backup_files,
            'total_backups': len(backup_files)
        }

    except Exception as e:
        return {
            'success': False,
            'error': 'LIST_BACKUPS_FAILED',
            'message': f"列出备份失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_insert_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
在指定位置插入空行。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.insert_rows(file_path, sheet_name, row_index, count)


@mcp.tool()
@_track_call
def excel_insert_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
在指定位置插入空列（1-based索引，1=A列）。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.insert_columns(file_path, sheet_name, column_index, count)


@mcp.tool()
@_track_call
def excel_find_last_row(
    file_path: str,
    sheet_name: str,
    column: Optional[Union[str, int]] = None
) -> Dict[str, Any]:
    """
查找工作表最后一行（有数据的最大行号）。追加数据前用此工具确定插入位置。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.find_last_row(file_path, sheet_name, column)


@mcp.tool()
@_track_call
def excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
创建新的空Excel文件（默认含1个空工作表）。创建后用excel_update_range写入数据。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.create_file(file_path, sheet_names)


@mcp.tool()
@_track_call
def excel_export_to_csv(
    file_path: str,
    output_path: str,
    sheet_name: Optional[str] = None,
    encoding: str = "utf-8"
) -> Dict[str, Any]:
    """
将Excel工作表导出为CSV文件。支持指定编码(utf-8/gbk)。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.export_to_csv(file_path, output_path, sheet_name, encoding)


@mcp.tool()
@_track_call
def excel_import_from_csv(
    csv_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    encoding: str = "utf-8",
    has_header: bool = True
) -> Dict[str, Any]:
    """
从CSV文件导入数据创建Excel文件。
    """
    for _p in [csv_path, output_path]:
        _err = _validate_path(_p)
        if _err:
            return _err

    return ExcelOperations.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header)


@mcp.tool()
@_track_call
def excel_convert_format(
    input_path: str,
    output_path: str,
    target_format: str = "xlsx"
) -> Dict[str, Any]:
    """
转换Excel文件格式。支持xlsx/xlsm/csv/json互转。
    """
    for _p in [input_path, output_path]:
        _err = _validate_path(_p)
        if _err:
            return _err

    return ExcelOperations.convert_format(input_path, output_path, target_format)


@mcp.tool()
@_track_call
def excel_merge_files(
    input_files: List[str],
    output_path: str,
    merge_mode: str = "sheets"
) -> Dict[str, Any]:
    """
合并多个Excel文件。支持三种模式: sheets(各文件独立工作表)/append(追加行)/horizontal(按列拼接)。
    """
    for _f in input_files:
        _err = _validate_path(_f)
        if _err:
            return _err

    return ExcelOperations.merge_files(input_files, output_path, merge_mode)


@mcp.tool()
@_track_call
def excel_get_file_info(file_path: str) -> Dict[str, Any]:
    """
获取Excel文件信息：大小、工作表数量、格式、创建/修改时间。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.get_file_info(file_path)


@mcp.tool()
@_track_call
def excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
创建新工作表，可指定位置。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.create_sheet(file_path, sheet_name, index)


@mcp.tool()
@_track_call
def excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
删除指定工作表。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录删除工作表操作日志
    operation_logger.log_operation("delete_sheet", {
        "sheet_name": sheet_name
    })

    try:
        result = ExcelOperations.delete_sheet(file_path, sheet_name)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "deleted_sheet": result.get('deleted_sheet', ''),
            "remaining_sheets": result.get('remaining_sheets', 0),
            "message": result.get('message', '')
        })

        return result

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除工作表操作失败: {str(e)}"
        })

        return {
            'success': False,
            'error': 'DELETE_SHEET_FAILED',
            'message': f"删除工作表操作失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]:
    """
重命名工作表。新名称不能与已有工作表重复。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.rename_sheet(file_path, old_name, new_name)


@mcp.tool()
@_track_call
def excel_delete_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
删除指定位置的行。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录删除操作日志
    operation_logger.log_operation("delete_rows", {
        "sheet_name": sheet_name,
        "row_index": row_index,
        "count": count
    })

    try:
        result = ExcelOperations.delete_rows(file_path, sheet_name, row_index, count)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "deleted_rows": result.get('deleted_rows', 0),
            "message": result.get('message', '')
        })

        return result

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除行操作失败: {str(e)}"
        })

        return {
            'success': False,
            'error': 'DELETE_ROWS_FAILED',
            'message': f"删除行操作失败: {str(e)}"
        }


@mcp.tool()
@_track_call
def excel_delete_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1
) -> Dict[str, Any]:
    """
删除指定位置的列（1-based索引）。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录删除列操作日志
    operation_logger.log_operation("delete_columns", {
        "sheet_name": sheet_name,
        "column_index": column_index,
        "count": count
    })

    try:
        result = ExcelOperations.delete_columns(file_path, sheet_name, column_index, count)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "deleted_columns": result.get('deleted_columns', 0),
            "message": result.get('message', '')
        })

        return result

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除列操作失败: {str(e)}"
        })

        return {
            'success': False,
            'error': 'DELETE_COLUMNS_FAILED',
            'message': f"删除列操作失败: {str(e)}"
        }

@mcp.tool()
@_track_call
def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """
设置单元格公式（不含等号）。如"SUM(A1:A10)"。返回计算结果。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return {'success': False, 'message': f'🔒 安全验证失败: {_formula_err["error"]}'}
    return ExcelOperations.set_formula(file_path, sheet_name, cell_address, formula)

@mcp.tool()
@_track_call
def excel_evaluate_formula(
    formula: str,
    context_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
临时执行公式并返回结果，不修改文件。可做快速计算器。
    """
    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return {'success': False, 'message': f'🔒 安全验证失败: {_formula_err["error"]}'}
    return ExcelOperations.evaluate_formula(formula, context_sheet)


@mcp.tool()
@_track_call
def excel_query(
    file_path: str,
    query_expression: str,
    include_headers: bool = True,
    output_format: str = "table"
) -> Dict[str, Any]:
    """
SQL查询Excel数据（只读）。优先使用此工具而非excel_get_range进行数据查询和分析。
参数query_expression只接受SELECT语句。批量修改请用excel_update_query。
支持中文列名、双行表头自动识别、数学表达式。
基础: SELECT/DISTINCT/别名(AS)/数学表达式(+-*/%)
条件: WHERE/AND/OR/LIKE/IN/NOT IN/BETWEEN/IS NULL/NOT
高级: 子查询(WHERE col IN(SELECT...))/CASE WHEN/COALESCE/EXISTS/CTE(WITH)
聚合: COUNT/SUM/AVG/MAX/MIN/COUNT(DISTINCT)/GROUP BY/HAVING/TOTAL行
排序: ORDER BY DESC/ASC/LIMIT/OFFSET
关联: INNER JOIN/LEFT JOIN（同文件内工作表）
字符串: UPPER/LOWER/TRIM/LENGTH/CONCAT/REPLACE/SUBSTRING/LEFT/RIGHT
不支持: UNION/窗口函数/INSERT/DELETE/RIGHT JOIN/CROSS JOIN/FROM子查询
输出格式: table(默认Markdown)/json/csv
示例: excel_query("技能表.xlsx", "SELECT 类型, AVG(伤害) FROM 技能配置 GROUP BY 类型 HAVING COUNT(*)>2 ORDER BY AVG(伤害) DESC")
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    # 参数验证
    if not file_path or not file_path.strip():
        return {
            'success': False,
            'message': '文件路径不能为空',
            'data': [],
            'query_info': {'error_type': 'parameter_validation'}
        }

    if not query_expression or not query_expression.strip():
        return {
            'success': False,
            'message': 'SQL查询语句不能为空',
            'data': [],
            'query_info': {'error_type': 'parameter_validation'}
        }

    # 验证 output_format
    valid_formats = ('table', 'json', 'csv')
    if output_format and output_format not in valid_formats:
        return {
            'success': False,
            'message': f'不支持的输出格式: {output_format}。可选: {", ".join(valid_formats)}',
            'data': [],
            'query_info': {'error_type': 'parameter_validation'}
        }

    # 使用高级SQL查询引擎
    try:
        from .api.advanced_sql_query import execute_advanced_sql_query
        return execute_advanced_sql_query(
            file_path=file_path,
            sql=query_expression,
            sheet_name=None,  # 统一使用SQL FROM子句中的表名
            limit=None,  # 统一使用SQL中的LIMIT
            include_headers=include_headers,
            output_format=output_format or 'table'
        )

    except ImportError:
        return {
            'success': False,
            'message': 'SQLGlot未安装，无法使用高级SQL功能。请运行: pip install sqlglot\n\n💡 智能降级建议：\n• 对于简单数据读取：尝试使用 excel_get_range("文件路径", "工作表名!A1:Z100")\n• 对于文本搜索：尝试使用 excel_search("文件路径", "关键词", "工作表名")\n• 对于表头信息：尝试使用 excel_get_headers("文件路径", "工作表名")',
            'data': [],
            'query_info': {
                'error_type': 'missing_dependency',
                'alternatives': ['excel_get_range', 'excel_search', 'excel_get_headers'],
                'suggestion': '使用基础Excel操作API作为保底方案'
            }
        }
    except Exception as e:
        # SQL引擎已处理大部分错误并返回结构化响应，此处仅捕获未预期的异常
        return {
            'success': False,
            'message': f'SQL查询失败: {str(e)}',
            'data': [],
            'query_info': {'error_type': 'unexpected_error', 'details': str(e)}
        }


@mcp.tool()
@_track_call
def excel_update_query(
    file_path: str,
    update_expression: str,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
SQL UPDATE批量修改Excel数据。优先使用此工具而非excel_update_range进行条件修改。
参数update_expression只接受UPDATE语句（如"UPDATE 表 SET 列=值 WHERE 条件"），查询请用excel_query。
SET支持: 列=常量/列引用/算术表达式(如 伤害*1.1, 攻击力+10)
WHERE支持: 全部excel_query条件语法(=/>/</LIKE/IN/BETWEEN/IS NULL/AND/OR/NOT)，含中文列名。
事务保护：写入失败自动回滚，不会损坏文件。
dry_run=True 可预览影响范围不实际修改。
示例: excel_update_query("技能表.xlsx", "UPDATE 技能配置 SET 伤害 = 伤害 * 1.1 WHERE 元素 = '火'", dry_run=True)
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    if not file_path or not file_path.strip():
        return {'success': False, 'message': '文件路径不能为空',
                'affected_rows': 0, 'changes': []}

    if not update_expression or not update_expression.strip():
        return {'success': False, 'message': 'UPDATE语句不能为空',
                'affected_rows': 0, 'changes': []}

    # 安全检查：只允许UPDATE语句
    stripped = update_expression.strip().upper()
    if not stripped.startswith('UPDATE'):
        return {'success': False,
                'message': '只支持UPDATE语句。查询请使用 excel_query',
                'affected_rows': 0, 'changes': []}

    try:
        from .api.advanced_sql_query import execute_advanced_update_query
        return execute_advanced_update_query(
            file_path=file_path,
            sql=update_expression,
            dry_run=dry_run
        )
    except ImportError:
        return {'success': False, 'message': 'SQLGlot未安装，无法使用UPDATE功能',
                'affected_rows': 0, 'changes': []}
    except Exception as e:
        return {'success': False, 'message': f'UPDATE执行失败: {str(e)}',
                'affected_rows': 0, 'changes': []}


@mcp.tool()
@_track_call
def excel_describe_table(
    file_path: str,
    sheet_name: str = None
) -> Dict[str, Any]:
    """
快速查看Excel工作表结构。查询前先用此工具了解有哪些列、什么类型。
自动识别双行表头（第1行中文描述+第2行英文字段名），输出中英映射。
返回: 列名、类型、非空数量、示例值、行数。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    if not file_path or not file_path.strip():
        return {'success': False, 'message': '文件路径不能为空', 'data': []}

    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    except Exception as e:
        return {'success': False, 'message': f'无法打开文件: {e}', 'data': []}

    try:
        # 选择工作表
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                return {
                    'success': False,
                    'message': f'工作表 "{sheet_name}" 不存在。可用工作表: {wb.sheetnames}',
                    'data': []
                }
            ws = wb[sheet_name]
        else:
            ws = wb.worksheets[0]
            sheet_name = ws.title

        # 读取前几行来判断表头类型
        rows = list(ws.iter_rows(max_row=4, values_only=True))
        if not rows:
            return {'success': False, 'message': '工作表为空', 'data': []}

        # 检测双行表头：第2行全是合法英文字段名
        is_dual_header = False
        header_row_idx = 0
        if len(rows) >= 2:
            second_row = rows[1]
            if second_row and len(second_row) >= 3:
                all_valid = all(
                    isinstance(c, str) and c.strip().startswith(tuple('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_'))
                    for c in second_row if c is not None
                )
                first_row = rows[0]
                any_chinese = any(
                    isinstance(c, str) and any('\u4e00' <= ch <= '\u9fff' for ch in c)
                    for c in first_row if c is not None
                )
                if all_valid and any_chinese:
                    is_dual_header = True
                    header_row_idx = 1

        headers = rows[header_row_idx]
        descriptions = rows[0] if is_dual_header else None

        # 读取所有数据行统计信息 — 单次遍历所有列（优化：旧方案逐列遍历→N列×M行，新方案一次遍历→M行）
        data_start = header_row_idx + 1
        num_cols = len(headers)
        col_stats = {}
        # 初始化统计结构
        col_name_list = []
        col_stats = {}
        for col_idx in range(num_cols):
            col_name = headers[col_idx]
            if col_name is None:
                col_name = f"column_{col_idx + 1}"
            col_name = str(col_name).strip()
            if not col_name:
                col_name = f"column_{col_idx + 1}"
            col_name_list.append(col_name)
            col_stats[col_name] = {'non_null': 0, 'samples': [], 'type_values': []}

        # 单次遍历所有行和列
        for row in ws.iter_rows(min_row=data_start + 1, values_only=True):
            for col_idx in range(min(len(row), num_cols)):
                val = row[col_idx]
                if val is not None:
                    s = col_stats[col_name_list[col_idx]]
                    s['non_null'] += 1
                    if len(s['samples']) < 3:
                        s['samples'].append(val)
                    if len(s['type_values']) < 100:
                        s['type_values'].append(val)

        # 推断类型并构建最终结果
        for col_idx, col_name in enumerate(col_name_list):
            s = col_stats[col_name]
            tv = s['type_values']
            if not tv:
                col_type = "empty"
            elif all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in tv):
                col_type = "number"
            elif all(isinstance(v, str) for v in tv):
                col_type = "text"
            else:
                col_type = "mixed"

            description = str(descriptions[col_idx]).strip() if is_dual_header and descriptions and col_idx < len(descriptions) and descriptions[col_idx] else None
            col_stats[col_name] = {
                'name': col_name, 'type': col_type, 'description': description,
                'non_null': s['non_null'], 'sample_values': s['samples']
            }

        # 统计行数
        row_count = ws.max_row - data_start if ws.max_row > data_start else 0

        wb.close()

        columns = list(col_stats.values())
        return {
            'success': True,
            'sheet_name': sheet_name,
            'header_type': 'dual' if is_dual_header else 'single',
            'row_count': row_count,
            'column_count': len(columns),
            'columns': columns,
            'message': f"表 '{sheet_name}': {len(columns)}列, {row_count}行数据, {'双行表头' if is_dual_header else '单行表头'}"
        }
    except Exception as e:
        return {'success': False, 'message': f'查看表结构失败: {e}', 'data': []}
    finally:
        wb.close()


@mcp.tool()
@_track_call
def excel_format_cells(
    file_path: str,
    sheet_name: str,
    range: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None
) -> Dict[str, Any]:
    """
格式化单元格样式。支持preset(highlight/warning/success)和自定义格式。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.format_cells(file_path, sheet_name, range, formatting, preset)


@mcp.tool()
@_track_call
def excel_merge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
合并指定范围的单元格（如"A1:C3"）。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.merge_cells(file_path, sheet_name, range)


@mcp.tool()
@_track_call
def excel_unmerge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
取消合并指定范围的单元格。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.unmerge_cells(file_path, sheet_name, range)


@mcp.tool()
@_track_call
def excel_set_borders(
    file_path: str,
    sheet_name: str,
    range: str,
    border_style: str = "thin"
) -> Dict[str, Any]:
    """
为范围设置边框。样式: thin/thick/medium/double/dotted/dashed。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.set_borders(file_path, sheet_name, range, border_style)


@mcp.tool()
@_track_call
def excel_set_row_height(
    file_path: str,
    sheet_name: str,
    row_index: int,
    height: float,
    count: int = 1
) -> Dict[str, Any]:
    """
调整行高（磅值，如25.0）。可同时调整连续多行。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.set_row_height(file_path, sheet_name, row_index, height, count)


@mcp.tool()
@_track_call
def excel_set_column_width(
    file_path: str,
    sheet_name: str,
    column_index: int,
    width: float,
    count: int = 1
) -> Dict[str, Any]:
    """
调整列宽（字符单位，如15.0）。可同时调整连续多列。
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.set_column_width(file_path, sheet_name, column_index, width, count)


# ==================== Excel比较功能 ====================

@mcp.tool()
@_track_call
def excel_compare_files(
    file1_path: str,
    file2_path: str
) -> Dict[str, Any]:
    """
对比两个Excel文件的所有工作表差异（结构差异+单元格值变化）。输出逐单元格对比。
按ID列做对象级对比（新增/删除/修改记录）请用excel_compare_sheets。
    """
    for _p in [file1_path, file2_path]:
        _err = _validate_path(_p)
        if _err:
            return _err
    return ExcelOperations.compare_files(file1_path, file2_path)


@mcp.tool()
@_track_call
def excel_check_duplicate_ids(
    file_path: str,
    sheet_name: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """
检查配置表ID列是否有重复值。返回重复ID及其所在行号。
也可用excel_query: SELECT ID, COUNT(*) as c FROM 表 GROUP BY ID HAVING c>1
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return ExcelOperations.check_duplicate_ids(file_path, sheet_name, id_column, header_row)


@mcp.tool()
@_track_call
def excel_compare_sheets(
    file1_path: str,
    sheet1_name: str,
    file2_path: str,
    sheet2_name: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """
按ID列精确对比两个工作表，输出新增/删除/修改的记录（对象级差异）。
逐单元格对比请用excel_compare_files。
    """
    for _p in [file1_path, file2_path]:
        _err = _validate_path(_p)
        if _err:
            return _err
    return ExcelOperations.compare_sheets(file1_path, sheet1_name, file2_path, sheet2_name, id_column, header_row)


@mcp.tool()
@_track_call
def excel_server_stats() -> Dict[str, Any]:
    """
获取MCP服务器运行统计：每个工具的调用次数、平均耗时、错误率。用于监控和调试。
    """
    return _tracker.get_stats()


# ==================== 主程序 ====================
def main():
    """Entry point for excel-mcp-server-fastmcp.

    支持三种传输模式：
    - stdio（默认）：本地使用，uvx/claude/cursor
    - sse：Server-Sent Events远程模式
    - streamable-http：Streamable HTTP远程模式，推荐用于团队共享
    """
    import sys
    if len(sys.argv) > 1 and sys.argv[1] in ('--version', '-v'):
        from excel_mcp_server_fastmcp import __version__
        print(f"excel-mcp-server-fastmcp {__version__}", flush=True)
        sys.exit(0)

    transport = 'stdio'
    mount_path = None
    for arg in sys.argv[1:]:
        if arg in ('--stdio', '--sse', '--streamable-http'):
            transport = arg[2:]  # remove '--'
        elif arg.startswith('--mount-path='):
            mount_path = arg.split('=', 1)[1]

    mcp.run(transport=transport, mount_path=mount_path)

if __name__ == "__main__":
    main()
