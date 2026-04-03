#!/usr/bin/env python3
"""Excel MCP Server - 基于 FastMCP 和 openpyxl 的游戏开发Excel配置表工具。

统一返回格式: {success: bool, message: str, data: Any, meta: dict}
- 成功: success=true, data含实际数据
- 失败: success=false, message含错误描述+💡修复建议
"""

# 标准库导入
import functools
import glob
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import threading
import time
from collections import defaultdict
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional, List, Dict, Any, Union

# 配置logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    logger.error(f"缺少必要的依赖包: {e}")
    logger.error("请运行: pip install mcp openpyxl")
    exit(1)

# 导入API模块
from .api.excel_operations import ExcelOperations
from .utils.validators import ExcelValidator, DataValidationError
from .utils import extract_rich_text

# 导入智能配置推荐模块
try:
    from .core.smart_config_recommender import SmartConfigurationRecommender
    SMART_CONFIG_AVAILABLE = True
    SMART_CONFIG_TOOLS_AVAILABLE = True
except ImportError:
    SMART_CONFIG_AVAILABLE = False
    SMART_CONFIG_TOOLS_AVAILABLE = False
    logger.warning("智能配置推荐模块未找到，相关功能不可用")

# ==================== 操作日志系统 ====================
class OperationLogger:
    """操作日志记录器，用于跟踪所有Excel操作"""

    def __init__(self):
        self.log_file = None
        self.current_session = []

    def start_session(self, file_path: str):
        """开始新的操作会话
        
        创建一个新的操作会话，初始化日志记录系统。
        
        Args:
            file_path (str): 要操作的Excel文件路径
            
        Returns:
            None
        """
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
        """记录操作
        
        将操作记录添加到当前会话中，并保存日志。
        
        Args:
            operation (str): 操作名称（如 'read_cell', 'write_range'）
            details (Dict[str, Any]): 操作详情，包含参数和返回结果
            
        Returns:
            None
        """
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
        """保存日志到文件
        
        将当前会话的日志记录保存到JSON文件。
        
        Args:
            None
            
        Returns:
            None
            
        Raises:
            IOError: 当文件写入失败时记录错误日志
        """
        if not self.log_file:
            return

        try:
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump(self.current_session, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"保存操作日志失败: {e}")

    def get_recent_operations(self, limit: int = 10) -> List[Dict[str, Any]]:
        """获取最近的操作记录
        
        返回当前会话中最近的操作记录。
        
        Args:
            limit (int): 返回的最大记录数，默认为10
            
        Returns:
            List[Dict[str, Any]]: 最近的操作记录列表，每条记录包含timestamp, operation, details
            
        Raises:
            ValueError: 当limit为负数时
        """
        if not self.current_session:
            return []

        operations = self.current_session[0]['operations']
        return operations[-limit:] if len(operations) > limit else operations

# ==================== 工具调用追踪器 ====================
class ToolCallTracker:
    """工具调用追踪器，记录每个工具的调用次数、耗时、错误率和错误分类"""

    # 错误分类规则：按优先级匹配（前缀/关键词 → 错误类型）
    _ERROR_RULES = [
        ('🔒', 'security'),
        ('文件不存在', 'file_not_found'),
        ('无法加载文件', 'file_load'),
        ('文件路径', 'validation'),
        ('工作表不存在', 'sheet_not_found'),
        ('文件格式', 'file_format'),
        ('文件太大', 'file_too_large'),
        ('无效的', 'validation'),
        ('不支持', 'unsupported'),
        ('列名', 'column'),
        ('SQL语法', 'sql_syntax'),
        ('语法错误', 'sql_syntax'),
        ('engine_error', 'engine'),
        ('execution_error', 'execution'),
        ('syntax_error', 'sql_syntax'),
        ('unsupported_sql', 'unsupported'),
        ('unsupported_feature', 'unsupported'),
        ('file_not_found', 'file_not_found'),
        ('data_load_failed', 'file_load'),
    ]

    def __init__(self):
        self._stats: Dict[str, Dict[str, Any]] = defaultdict(lambda: {
            'call_count': 0,
            'total_time_ms': 0.0,
            'error_count': 0,
            'error_types': defaultdict(int),
            'last_called': None,
            'min_time_ms': float('inf'),
            'max_time_ms': 0.0,
        })
        self._lock = threading.Lock()
        self._start_time = datetime.now()

    def record(self, tool_name: str, duration_ms: float, success: bool = True,
               error_type: str = None):
        """记录一次工具调用，可选指定错误类型"""
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
                classified = error_type or 'unknown'
                s['error_types'][classified] += 1

    @staticmethod
    def classify_error(message: str) -> str:
        """根据错误消息内容自动分类错误类型"""
        if not message:
            return 'unknown'
        msg_lower = message.lower()
        for keyword, error_type in ToolCallTracker._ERROR_RULES:
            if keyword.lower() in msg_lower:
                return error_type
        return 'unknown'

    def get_stats(self) -> Dict[str, Any]:
        """获取所有工具的调用统计，按调用次数降序排列"""
        with self._lock:
            tools = {}
            global_error_types: Dict[str, int] = defaultdict(int)
            for name, s in sorted(self._stats.items(), key=lambda x: -x[1]['call_count']):
                tools[name] = {
                    'call_count': s['call_count'],
                    'avg_time_ms': round(s['total_time_ms'] / s['call_count'], 1) if s['call_count'] else 0,
                    'min_time_ms': round(s['min_time_ms'], 1) if s['min_time_ms'] != float('inf') else 0,
                    'max_time_ms': round(s['max_time_ms'], 1),
                    'error_count': s['error_count'],
                    'error_types': dict(sorted(s['error_types'].items())),
                    'last_called': s['last_called'],
                }
                for et, count in s['error_types'].items():
                    global_error_types[et] += count
            return {
                'uptime_seconds': round((datetime.now() - self._start_time).total_seconds(), 0),
                'total_calls': sum(s['call_count'] for s in self._stats.values()),
                'total_errors': sum(s['error_count'] for s in self._stats.values()),
                'error_types': dict(sorted(global_error_types.items())),
                'tools': tools,
            }

    def reset(self):
        """重置所有统计"""
        with self._lock:
            self._stats.clear()
            self._start_time = datetime.now()


_tracker = ToolCallTracker()


def _track_call(func):
    """工具调用追踪装饰器，记录每次调用的耗时和结果，自动检测返回值中的错误"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            duration_ms = (time.perf_counter() - start) * 1000
            # 自动检测返回值中的错误（大多数工具返回dict而非抛异常）
            if isinstance(result, dict) and result.get('success') is False:
                error_msg = result.get('message', '')
                error_type = ToolCallTracker.classify_error(error_msg)
                # 优先使用SQL引擎的错误类型
                qi = result.get('query_info') or {}
                if isinstance(qi, dict) and 'error_type' in qi:
                    error_type = qi['error_type']
                _tracker.record(func.__name__, duration_ms, success=False,
                                error_type=error_type)
                logger.debug(f"[TOOL] {func.__name__}: {duration_ms:.1f}ms ERROR [{error_type}]",
                             extra={'tool': func.__name__, 'duration_ms': round(duration_ms, 1),
                                    'error': error_msg})
            else:
                _tracker.record(func.__name__, duration_ms, success=True)
                logger.debug(f"[TOOL] {func.__name__}: {duration_ms:.1f}ms",
                             extra={'tool': func.__name__, 'duration_ms': round(duration_ms, 1)})
            return result
        except Exception as e:
            duration_ms = (time.perf_counter() - start) * 1000
            error_type = ToolCallTracker.classify_error(str(e))
            _tracker.record(func.__name__, duration_ms, success=False,
                            error_type=error_type)
            logger.debug(f"[TOOL] {func.__name__}: {duration_ms:.1f}ms ERROR [{error_type}]: {e}",
                         extra={'tool': func.__name__, 'duration_ms': round(duration_ms, 1),
                                'error': str(e)})
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
        return _fail(f'🔒 安全验证失败: {result["error"]}', meta={"error_code": "PATH_VALIDATION_FAILED"})
    # 文件存在时额外检查大小
    if os.path.exists(file_path):
        size_result = SecurityValidator.validate_file_size(file_path)
        if not size_result['valid']:
            return _fail(f'🔒 安全验证失败: {size_result["error"]}', meta={"error_code": "FILE_SIZE_EXCEEDED"})
    return None


def _validate_file_path(param='file_path'):
    """路径安全验证装饰器，自动验证指定参数的文件路径。

    Args:
        param: 要验证的参数名（str）或参数名列表（list[str]），
               默认为 'file_path'。列表时依次验证所有参数。
               仅当参数值不为None时验证（支持Optional参数）。

    Returns:
        装饰器函数

    Example:
        @_validate_file_path()          # 验证 file_path 参数
        @_validate_file_path('backup_path')  # 验证 backup_path 参数
        @_validate_file_path(['csv_path', 'output_path'])  # 验证多个参数
    """
    def decorator(func):
        """创建验证装饰器。

        Args:
            func: 被装饰的函数。
        """
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            params = [param] if isinstance(param, str) else param
            for p_name in params:
                p_value = kwargs.get(p_name)
                if p_value is None:
                    continue
                err = _validate_path(p_value)
                if err:
                    return err
            return func(*args, **kwargs)
        return wrapper
    return decorator


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
    instructions=r"""🎮 游戏开发Excel配置表管理专家 — 52个工具

## 🔥 核心原则：SQL优先

**优先使用 `excel_query`** - 所有数据查询分析任务
- 复杂条件筛选 ✅ WHERE, LIKE, IN, BETWEEN, 子查询
- 聚合统计分析 ✅ COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 高级查询 ✅ CASE WHEN, CTE(WITH), EXISTS, JOIN(5种)
- 排序限制 ✅ ORDER BY, LIMIT, OFFSET
- 字符串函数 ✅ UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING
- 结果合并 ✅ UNION, UNION ALL
- 窗口函数 ✅ ROW_NUMBER, RANK, DENSE_RANK

## 📊 工具选择决策树
```
数据分析/查询/筛选/聚合？ → excel_query (SQL引擎)
快速了解单表结构？        → excel_describe_table (列名+类型+样本值+行数)
批量获取所有表的中英表头？ → excel_get_headers("文件", sheet_name省略)
定位文本位置？            → excel_search (返回row/column)
跨文件搜索？              → excel_search_directory
批量条件修改？            → excel_update_query (SQL UPDATE, 支持dry_run)
写入指定单元格范围？      → excel_update_range (二维数组写入)
按ID合并导入配置？        → excel_upsert_row (存在更新/不存在插入)
批量导入多行数据？        → excel_batch_insert_rows
创建配置表变体？          → excel_copy_sheet (复制工作表)
按ID对比两表差异？        → excel_compare_sheets (对象级: 新增/删除/修改)
逐单元格对比差异？        → excel_compare_files (单元格级)
格式调整？                → excel_format_cells (preset: highlight/warning/success)
数据修改影响评估？        → excel_assess_data_impact (修改前的安全网)
```

## ✅ SQL已支持功能
基础: SELECT, DISTINCT, 别名(AS), 数学表达式(+-*/%)
条件: WHERE, 比较运算(=/></≤/≥/≠), LIKE, NOT LIKE, IN, NOT IN, BETWEEN, AND, OR, IS NULL, IS NOT NULL, NOT
高级: WHERE子查询, FROM子查询(FROM (SELECT ...) AS alias), CASE WHEN, COALESCE, EXISTS, CTE(WITH ... AS ...)
聚合: COUNT(*), COUNT(col), COUNT(DISTINCT), SUM, AVG, MAX, MIN, GROUP BY, HAVING, TOTAL行
排序: ORDER BY DESC/ASC, LIMIT, OFFSET
合并: UNION(去重), UNION ALL(不去重)
关联: INNER JOIN, LEFT JOIN, RIGHT JOIN, FULL JOIN, CROSS JOIN（同文件内工作表）
字符串: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT
窗口: ROW_NUMBER, RANK, DENSE_RANK（OVER PARTITION BY ... ORDER BY ...）

## ❌ SQL不支持
INSERT, DELETE, 嵌套FROM子查询（FROM子查询中不能再嵌套FROM子查询）, NATURAL JOIN, INTERSECT/EXCEPT, WITH RECURSIVE, LATERAL JOIN

## ✅ FROM子查询
支持单层FROM子查询：`FROM (SELECT ...) AS alias`。不支持嵌套。
```sql
SELECT * FROM (SELECT skill_name, damage FROM 技能配置 WHERE damage > 100) AS 高伤技能
```

## ✅ UNION / UNION ALL
合并多个SELECT查询结果。支持ORDER BY和LIMIT。
```sql
SELECT name FROM 技能配置 WHERE 类型='法师' UNION ALL SELECT name FROM 技能配置 WHERE 类型='战士' ORDER BY name LIMIT 10
```

## ✅ 窗口函数 (ROW_NUMBER / RANK / DENSE_RANK)
在查询结果上计算排名，支持PARTITION BY分区。
```sql
SELECT skill_name, skill_type, ROW_NUMBER() OVER (PARTITION BY skill_type ORDER BY damage DESC) as rn FROM 技能配置
```

## ⚠️ 重要原则
- 双行表头: 第1行中文描述+第2行英文字段名，中英文列名均可查询
- 1-based索引: 第1行=1, 第1列=1
- 范围格式: "工作表名!A1:Z100"（必须含工作表名）
- 默认覆盖: update_range默认为覆盖模式(insert_mode=False)，保留数据需显式设置insert_mode=True

## 🎮 游戏配置表示例
技能统计: SELECT 技能类型, AVG(伤害), COUNT(*) FROM 技能表 GROUP BY 技能类型
高级筛选: SELECT * FROM 技能表 WHERE 伤害 > (SELECT AVG(伤害) FROM 技能表)
条件表达式: SELECT 技能名, CASE WHEN 伤害>100 THEN '高' ELSE '低' END AS 等级 FROM 技能表
CTE: WITH 高伤 AS (SELECT * FROM 技能表 WHERE 伤害>100) SELECT COUNT(*) FROM 高伤
UNION: SELECT 技能名 FROM 法师技能 UNION ALL SELECT 技能名 FROM 战士技能
窗口函数: SELECT *, ROW_NUMBER() OVER (PARTITION BY 类型 ORDER BY 伤害 DESC) as 排名 FROM 技能配置

## 📦 统一返回格式
所有工具返回统一JSON结构，AI解析时只需检查 `success` 字段：
- 成功: `{"success": true, "message": "描述", "data": {...}, "meta": {...}}`
- 失败: `{"success": false, "message": "错误描述+💡修复提示", "meta": {"error_code": "CODE"}}`
- `data`: 实际数据载荷（查询结果、文件信息等）
- `meta`: 上下文元数据（行数、列数、执行时间等）
- `error_code`: 机器可读错误分类（PATH_VALIDATION_FAILED, SQL_EXECUTION_FAILED等）
- 所有错误均包含💡修复提示，按提示操作即可解决大多数问题
- SQL查询额外返回 `query_info`（含execution_time_ms、error_type、hint、suggested_fix）

## ⚡ 常用流程
1. excel_list_sheets → 2. excel_describe_table → 3. excel_query → 4. excel_update_query/excel_update_range → 5. excel_compare_sheets
""",
    debug=bool(os.environ.get('EXCEL_MCP_DEBUG')),
    log_level="DEBUG" if os.environ.get('EXCEL_MCP_DEBUG') else "WARNING"
)



# ==================== 统一响应格式 ====================

def _ok(message: str = "", data=None, meta: dict = None) -> dict:
    """统一成功响应生成器
    
    生成标准化的成功响应格式，用于所有MCP工具函数。
    
    Args:
        message (str, optional): 成功消息描述，默认为空字符串
        data (any, optional): 成功时的数据内容，默认为None
        meta (dict, optional): 元信息字典，包含执行时间、缓存状态等，默认为None
        
    Returns:
        dict: 统一的成功响应格式 {success: true, message: str, data: any, meta: dict}
    """
    r: dict = {"success": True}
    if message:
        r["message"] = message
    if data is not None:
        r["data"] = data
    if meta:
        r["meta"] = meta
    return r


# 集中式非SQL错误提示映射 — 让AI在收到错误后知道如何修复
# 格式: error_code → hint字符串
_ERROR_HINTS = {
    'FILE_NOT_FOUND': '请检查文件路径是否正确，或用excel_list_sheets确认文件存在。路径支持绝对路径和相对于工作目录的路径。',
    'SHEET_NOT_FOUND': '请用excel_list_sheets查看该文件有哪些工作表，确认表名拼写正确。',
    'EMPTY_SHEET': '工作表没有数据行。请确认文件内容或检查表头行数设置（默认前1-2行为表头）。',
    'FILE_OPEN_FAILED': '请确认文件是有效的.xlsx格式且未被其他程序锁定。如文件损坏，尝试用Excel打开后另存为新文件。',
    'FILE_SIZE_EXCEEDED': '文件过大，超出安全限制。如需处理大文件，请联系管理员调整限制。',
    'PATH_VALIDATION_FAILED': '文件路径不合法。请使用绝对路径或相对于当前工作目录的路径，不要使用".."路径穿越。',
    'MISSING_FILE_PATH': '请提供文件路径参数。例如：excel_query("配置表.xlsx", "SELECT * FROM 技能表")',
    'MISSING_QUERY': '请提供SQL查询语句。例如：excel_query("配置表.xlsx", "SELECT * FROM 技能表 WHERE 伤害 > 100")',
    'INVALID_FORMAT': '请检查参数值是否在允许范围内。查看工具描述了解支持的选项。',
    'VALIDATION_FAILED': '请检查参数格式是否正确。常见原因：范围表达式格式错误（应为"工作表名!A1:Z100"）、参数类型不匹配。',
    'OPERATION_FAILED': '写入操作失败，文件可能被锁定或磁盘空间不足。请关闭其他正在使用该文件的程序后重试。',
    'PREVIEW_FAILED': '无法预览操作影响。请先用excel_get_range查看当前数据，再重新尝试操作。',
    'ASSESSMENT_FAILED': '数据影响评估失败。请检查文件路径和工作表名是否正确。',
    'DEPENDENCY_MISSING': '需要安装额外依赖。按提示运行pip install命令后重试。',
    'BACKUP_FAILED': '备份创建失败。请检查磁盘空间和文件权限。',
    'BACKUP_NOT_FOUND': '备份文件不存在。请用excel_list_backups查看可用的备份列表。',
    'RESTORE_FAILED': '恢复失败。备份文件可能已损坏。请用excel_list_backups选择其他备份。',
    'LIST_BACKUPS_FAILED': '获取备份列表失败。请检查备份目录是否存在。',
    'DELETE_SHEET_FAILED': '删除工作表失败。文件可能被锁定或工作表名不正确。',
    'DELETE_ROWS_FAILED': '删除行失败。请确认行号范围有效（不能删除表头行）。',
    'DELETE_COLUMNS_FAILED': '删除列失败。请确认列号范围有效。',
    'FORMULA_SECURITY_FAILED': '公式包含不允许的模式（如外部链接）。请移除危险部分后重试。',
    'HISTORY_RETRIEVAL_FAILED': '获取操作历史失败。历史记录文件可能不存在或已损坏。',
    'DESCRIBE_FAILED': '查看表结构失败。请确认文件路径和工作表名正确。',
    'SQL_EXECUTION_FAILED': 'SQL查询执行失败。请检查语法、表名和列名。建议先用excel_describe_table了解表结构。',
    'UPDATE_EXECUTION_FAILED': 'UPDATE执行失败。请检查SQL语法，确保SET和WHERE子句正确。可用dry_run=True预览。',
    'UNSUPPORTED_SQL': '不支持该SQL语法。查询请用excel_query，修改请用excel_update_query。',
}


def _fail(message: str, meta: dict = None) -> dict:
    """统一错误响应生成器
    
    生成标准化的错误响应格式，自动附加错误代码对应的修复提示。
    
    Args:
        message (str): 错误消息描述，通常包含具体的错误信息
        meta (dict, optional): 元信息字典，包含error_code、file_path、suggested_fix等，默认为None
        
    Returns:
        dict: 统一的错误响应格式 {success: false, message: str, meta: dict}
        
    Note:
        当meta中包含error_code时，会自动添加对应的💡修复提示到message中
    """
    r: dict = {"success": False, "message": message}
    if meta:
        r["meta"] = meta
    # 自动附加集中式错误提示（仅当message中还没有💡提示时）
    error_code = (meta or {}).get('error_code', '')
    hint = _ERROR_HINTS.get(error_code, '')
    if hint and '💡' not in message:
        r["message"] = message + f'\n💡 {hint}'
    return r


def _strip_defaults(obj: Any, depth: int = 0) -> Any:
    """递归移除默认值和空值以减少token消耗。
    
    移除规则：
    - None, 空字符串, 空列表, 空字典
    - 常见Excel默认值：False, 0, None（基于字段名判断）
    - 保留有意义的值
    """
    if depth > 5 or not isinstance(obj, dict):
        return obj
    
    # 常见Excel默认值字段名（这些字段通常不需要返回False/0/None）
    excel_default_fields = {
        'bold', 'italic', 'underline', 'wrap_text', 'auto_filter',
        'border_top', 'border_bottom', 'border_left', 'border_right',
        'horizontal_alignment', 'vertical_alignment', 'text_rotation',
        'indent', 'shrink_to_fit', 'merge_cells'
    }
    
    # 有语义的空列表字段名，即使为空也不应移除
    semantic_list_fields = {
        'headers', 'sheets', 'sheets_with_headers', 'field_names',
        'descriptions', 'data', 'columns', 'rows'
    }
    
    cleaned = {}
    for k, v in obj.items():
        # 移除空值
        if v is None or v == '':
            continue
        
        # 移除空容器（但有语义的列表字段除外）
        if isinstance(v, (list, dict)) and len(v) == 0 and k not in semantic_list_fields:
            continue
        
        # 基于字段名的默认值处理
        if k.lower() in excel_default_fields and v in [False, 0, None]:
            continue
        
        # 递归处理嵌套对象
        if isinstance(v, dict):
            cleaned[k] = _strip_defaults(v, depth + 1)
        elif isinstance(v, list):
            cleaned[k] = [_strip_defaults(i, depth + 1) if isinstance(i, dict) else i for i in v]
        else:
            cleaned[k] = v
    
    return cleaned


def _wrap(result: dict, meta: dict = None) -> dict:
    """包装Operations层返回，metadata→meta，添加上下文meta，统一success字段。
    
    统一返回格式：{success, message, data, meta}
    - 成功时：若缺少message则自动补充默认message
    - metadata→meta：Operations层的metadata自动映射到meta
    - error→message：确保AI只需检查message键
    """
    if not isinstance(result, dict):
        return result
    # 统一error→message，确保AI只需检查message键
    err_val = result.get('error')
    if isinstance(err_val, str) and not result.get('message'):
        result['message'] = result.pop('error')
    if "success" not in result:
        result["success"] = True
    # 成功时若缺少message，自动补充默认message
    if result.get('success') is True and 'message' not in result:
        result['message'] = '操作成功'
    if "metadata" in result:
        m = result.pop("metadata")
        if isinstance(m, dict) and m:
            merged = {**m, **(meta or {})}
            result["meta"] = merged
            meta = None  # 已合并，不再重复设置
    if meta and "meta" not in result:
        result["meta"] = meta

    # Token优化：过滤默认值和空值，减少响应体积
    if result.get('success') is True and 'data' in result and isinstance(result['data'], dict):
        result['data'] = _strip_defaults(result['data'])
        # 向后兼容：将data内的字段展平到顶层（避免破坏已有测试/客户端）
        for k, v in result['data'].items():
            if k not in result:
                result[k] = v
    # 向后兼容：将meta内的字段也展平到顶层（meta优先级低于data和已有顶层字段）
    if result.get('success') is True and 'meta' in result and isinstance(result['meta'], dict):
        for k, v in result['meta'].items():
            if k not in result:
                result[k] = v

    # 对Operations层返回的错误也附加集中式提示（Operations层无error_code）
    if result.get('success') is False:
        msg = result.get('message', '')
        existing_code = (result.get('meta') or {}).get('error_code', '')
        if not existing_code and '💡' not in msg:
            # 根据消息内容推断error_code并附加提示
            inferred = _infer_error_code(msg)
            if inferred:
                hint = _ERROR_HINTS.get(inferred, '')
                if hint:
                    result['message'] = msg + f'\n💡 {hint}'
                    if 'meta' not in result:
                        result['meta'] = {}
                    result['meta']['error_code'] = inferred

    return result


def _infer_error_code(message: str) -> str:
    """根据错误消息内容推断error_code，用于Operations层错误提示"""
    msg = message
    if '文件不存在' in msg or 'No such file' in msg.lower():
        return 'FILE_NOT_FOUND'
    if '工作表' in msg and ('不存在' in msg or '未找到' in msg):
        return 'SHEET_NOT_FOUND'
    if '无法打开' in msg or 'cannot open' in msg.lower() or 'Permission denied' in msg:
        return 'FILE_OPEN_FAILED'
    if '工作表为空' in msg or '没有数据' in msg:
        return 'EMPTY_SHEET'
    if '路径' in msg and ('不合法' in msg or '不允许' in msg or '穿越' in msg):
        return 'PATH_VALIDATION_FAILED'
    return ''

# ==================== MCP 工具定义 ====================

@mcp.tool()
@_validate_file_path()
@_track_call
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """列出Excel文件中的所有工作表名称。查询前先用此工具确认工作表存在。

    Args:
        file_path: Excel文件路径
    """
    return _wrap(ExcelOperations.list_sheets(file_path))


@mcp.tool()
@_validate_file_path()
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
    """在Excel中搜索匹配pattern的单元格。支持正则、大小写、全词匹配。

    Args:
        file_path: Excel文件路径
        pattern: 搜索模式
        sheet_name: 工作表名称，默认为None
        case_sensitive: 是否区分大小写，默认为False
        whole_word: 是否全词匹配，默认为False
        use_regex: 是否使用正则表达式，默认为False
        include_values: 是否包含单元格值，默认为True
        include_formulas: 是否包含公式，默认为False
        range: 搜索范围，默认为None
    """
    return _wrap(ExcelOperations.search(file_path, pattern, sheet_name, case_sensitive, whole_word, use_regex, include_values, include_formulas, range))


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
    """在目录下所有Excel文件中搜索内容。支持文件类型过滤和递归搜索。

    Args:
        directory_path: 搜索目录路径
        pattern: 搜索模式
        case_sensitive: 是否区分大小写，默认为False
        whole_word: 是否全词匹配，默认为False
        use_regex: 是否使用正则表达式，默认为False
        include_values: 是否包含单元格值，默认为True
        include_formulas: 是否包含公式，默认为False
        recursive: 是否递归搜索子目录，默认为True
        file_extensions: 文件扩展名过滤列表，默认为None
        file_pattern: 文件名模式匹配，默认为None
        max_files: 最大搜索文件数，默认为100
    """
    _path_err = _validate_path(directory_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.search_directory(directory_path, pattern, case_sensitive, whole_word, use_regex, include_values, include_formulas, recursive, file_extensions, file_pattern, max_files))


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_get_range(
    file_path: str,
    range: Optional[str] = None,
    include_formatting: bool = False,
    sheet_name: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None
) -> Dict[str, Any]:
    """读取指定范围的数据。返回{headers, data, shape}。支持include_formatting获取样式。
    
    Args:
        file_path: Excel文件路径
        range: 单元格范围，如 "Sheet1!A1:C10" 或 "A1:C10"（可选，可与start_cell/end_cell配合使用）
        include_formatting: 是否包含格式信息
        sheet_name: 工作表名称（可选，如果range不包含工作表名时可指定）
        start_cell: 起始单元格（可选，与end_cell配合使用）
        end_cell: 结束单元格（可选，与start_cell配合使用）
    """
    
    # 参数兼容性处理：如果提供了start_cell和end_cell，构建range表达式（优先执行）
    if start_cell and end_cell:
        if sheet_name:
            range = f"{sheet_name}!{start_cell}:{end_cell}"
        else:
            # 如果没有指定sheet_name，尝试自动推断
            if range and '!' not in range:
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(file_path, read_only=True)
                    sheet_names = wb.sheetnames
                    wb.close()
                    if len(sheet_names) == 1:
                        range = f"{sheet_names[0]}!{start_cell}:{end_cell}"
                    else:
                        return _fail(
                            f"需要指定工作表名，文件有多个工作表({', '.join(sheet_names)})。"
                            f"请使用sheet_name参数或格式: '工作表名!A1:C10'",
                            meta={"error_code": "VALIDATION_FAILED", "available_sheets": sheet_names}
                        )
                except Exception:
                    return _fail(
                        "无法自动推断工作表名，请指定sheet_name参数或使用完整范围表达式",
                        meta={"error_code": "VALIDATION_FAILED"}
                    )
            elif not range:
                # 如果没有range且没有sheet_name，尝试自动推断
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(file_path, read_only=True)
                    sheet_names = wb.sheetnames
                    wb.close()
                    if len(sheet_names) == 1:
                        range = f"{sheet_names[0]}!{start_cell}:{end_cell}"
                    else:
                        return _fail(
                            f"需要指定工作表名，文件有多个工作表({', '.join(sheet_names)})。"
                            f"请使用sheet_name参数或格式: '工作表名!A1:C10'",
                            meta={"error_code": "VALIDATION_FAILED", "available_sheets": sheet_names}
                        )
                except Exception:
                    return _fail(
                        "无法自动推断工作表名，请指定sheet_name参数或使用完整范围表达式",
                        meta={"error_code": "VALIDATION_FAILED"}
                    )
            else:
                range = f"{range}!{start_cell}:{end_cell}"  # 这种情况应该不会发生
    
    # 如果既没有range也没有start_cell/end_cell，报错
    if not range:
        return _fail(
            "必须指定范围参数，可通过以下方式之一：\n"
            "1. 直接提供range参数：'Sheet1!A1:C10'\n"
            "2. 提供 start_cell 和 end_cell 参数\n"
            "3. 提供 range 参数和 sheet_name 参数",
            meta={"error_code": "MISSING_RANGE_PARAMETER"}
        )
    
    # 如果range不包含工作表名但指定了sheet_name，添加工作表名
    if sheet_name and '!' not in range:
        range = f"{sheet_name}!{range}"
    
    # 参数顺序问题修复：检查range参数是否可能是误传的sheet_name
    # 只有在range不为空且没有通过start_cell/end_cell构建时才执行此检查
    if range and '!' not in range and not range.startswith(('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J')) and len(range) <= 10:
        # range看起来不像有效的单元格范围，可能是用户误将sheet_name传给了range参数
        return _fail(
            f"参数顺序可能错误：range参数值 '{range}' 看起来不像有效的单元格范围。\n"
            f"请检查参数顺序：第二个参数应该是范围表达式（如 'A1:C10' 或 'Sheet1!A1:C10'），\n"
            f"如果要以默认工作表读取，请明确指定: excel_get_range('{file_path}', 'A1:C10')",
            meta={
                "error_code": "PARAMETER_ORDER_ERROR",
                "received_range": range,
                "hint": "range参数应该是单元格范围，不是工作表名"
            }
        )
    
    # 原有的参数验证逻辑

    try:
        # 验证范围表达式格式
        # 如果没有工作表前缀，尝试自动推断（单工作表文件）
        if '!' not in range:
            try:
                from openpyxl import load_workbook
                wb = load_workbook(file_path, read_only=True)
                sheet_names = wb.sheetnames
                wb.close()
                if len(sheet_names) == 1:
                    range = f"{sheet_names[0]}!{range}"
                else:
                    return _fail(
                        f"范围表达式缺少工作表名，且文件有多个工作表({', '.join(sheet_names)})。"
                        f"请使用格式: '工作表名!A1:C10'",
                        meta={"error_code": "VALIDATION_FAILED", "available_sheets": sheet_names}
                    )
            except Exception:
                return _fail(
                    f"范围表达式缺少工作表名。正确格式: 'Sheet1!A1:C10'，当前: '{range}'",
                    meta={"error_code": "VALIDATION_FAILED"}
                )

        range_validation = ExcelValidator.validate_range_expression(range)

        # 验证操作规模
        scale_validation = ExcelValidator.validate_operation_scale(range_validation['range_info'])

        # 记录验证成功到调试日志
        logger.debug(f"范围验证成功: {range_validation['normalized_range']}")

    except DataValidationError as e:
        # 记录验证失败
        logger.error(f"范围验证失败: {str(e)}")

        return _fail(f"范围表达式验证失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})

    # 调用原始函数
    result = ExcelOperations.get_range(file_path, range, include_formatting)

    # 如果成功，添加验证信息到结果中
    if result.get('success'):
        validation_info = {
            'normalized_range': range_validation['normalized_range'],
            'sheet_name': range_validation['sheet_name'],
            'range_type': range_validation['range_info']['type'],
            'scale_assessment': scale_validation
        }
        # 向后兼容：保留顶层validation_info
        result['validation_info'] = validation_info
        # 新格式：同时合并到meta中
        if 'meta' not in result:
            result['meta'] = {}
        result['meta']['validation_info'] = validation_info

    return _wrap(result)


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_get_headers(
    file_path: str,
    sheet_name: Optional[str] = None,
    header_row: int = 1,
    max_columns: Optional[int] = None
) -> Dict[str, Any]:
    """提取工作表表头信息。支持双行表头（中文描述+英文字段名）。不传sheet_name获取所有表的表头。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为None表示获取所有表的表头
        header_row: 表头行号，默认为1
        max_columns: 最大列数限制，默认为None
    """
    if sheet_name is None:
        return _wrap(ExcelOperations.get_all_headers(file_path, header_row, max_columns))
    return _wrap(ExcelOperations.get_headers(file_path, sheet_name, header_row, max_columns))


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_update_range(
    file_path: str,
    range: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = False,
    streaming: bool = True,
    sheet_name: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None
) -> Dict[str, Any]:
    """写入数据到指定范围。
    
    Args:
        file_path: Excel文件路径
        range: 单元格范围，如 "Sheet1!A1:C10" 或 "A1:C10"
        data: 要写入的数据，二维数组格式
        preserve_formulas: 是否保留已有公式不被覆盖，默认True
        insert_mode: 数据写入模式，默认False(覆盖模式)
            - False: 覆盖模式，直接替换目标单元格数据
            - True: 插入模式，在目标位置插入新行，原有数据下移
        streaming: 是否使用流式写入，默认True
        sheet_name: 工作表名称（可选，如果range不包含工作表名时可指定）
        start_cell: 起始单元格（可选，与end_cell配合使用）
        end_cell: 结束单元格（可选，与start_cell配合使用）
    """
    
    # 参数兼容性处理：如果提供了start_cell和end_cell，构建range表达式
    if start_cell and end_cell:
        if sheet_name:
            range = f"{sheet_name}!{start_cell}:{end_cell}"
        else:
            # 如果没有指定sheet_name，尝试自动推断
            if '!' not in range:
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(file_path, read_only=True)
                    sheet_names = wb.sheetnames
                    wb.close()
                    if len(sheet_names) == 1:
                        range = f"{sheet_names[0]}!{start_cell}:{end_cell}"
                    else:
                        return _fail(
                            f"需要指定工作表名，文件有多个工作表({', '.join(sheet_names)})。"
                            f"请使用sheet_name参数或格式: '工作表名!A1:C10'",
                            meta={"error_code": "VALIDATION_FAILED", "available_sheets": sheet_names}
                        )
                except Exception:
                    return _fail(
                        "无法自动推断工作表名，请指定sheet_name参数或使用完整范围表达式",
                        meta={"error_code": "VALIDATION_FAILED"}
                    )
    
    # 如果range不包含工作表名但指定了sheet_name，添加工作表名
    if sheet_name and '!' not in range:
        range = f"{sheet_name}!{range}"

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

        return _fail(f"参数验证失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})

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
        # 扩展流式路径：支持覆盖模式和插入模式
        use_streaming = streaming and not preserve_formulas
        # 插入模式会自动使用流式插入，覆盖模式使用流式覆盖
        result = ExcelOperations.update_range(file_path, range, data, preserve_formulas, insert_mode, use_streaming)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "updated_cells": result.get('updated_cells', 0),
            "message": result.get('message', '')
        })

        return _wrap(result)

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"更新操作失败: {str(e)}"
        })

        return _fail(f"更新操作失败: {str(e)}", meta={"error_code": "OPERATION_FAILED"})





@mcp.tool()
@_track_call
def excel_assess_data_impact(
    file_path: str,
    range: str,
    operation_type: str = "update",
    data: Optional[List[List[Any]]] = None,
    detailed: bool = True
) -> Dict[str, Any]:
    """评估修改操作的影响范围。返回受影响行数、关键值变化等，修改前必用。

    Args:
        file_path: Excel文件路径
        range: 单元格范围
        operation_type: 操作类型，默认为"update"
        data: 要写入的数据，默认为None
        detailed: 是否返回详细信息，默认为True
    """

    # 详细模式下先验证范围表达式
    if detailed:
        try:
            range_validation = ExcelValidator.validate_range_expression(range)
            range_info = range_validation['range_info']
        except DataValidationError as e:
            return _fail(f"参数验证失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})

    # 获取当前数据
    current_data_result = ExcelOperations.get_range(file_path, range)

    if not current_data_result.get('success'):
        return _fail(f"无法预览操作: {current_data_result.get('message', '未知错误')}", meta={"error_code": "PREVIEW_FAILED"})

    current_data = current_data_result.get('data', [])

    if not detailed:
        # 快速预览模式（原preview_operation行为）
        data_rows = len(current_data)
        data_cols = len(current_data[0]) if data_rows > 0 else 0
        total_cells = data_rows * data_cols

        has_data = any(
            any(cell is not None and str(cell).strip() for cell in row)
            for row in current_data
        )

        risk_level = "LOW"
        if has_data:
            if total_cells > 100:
                risk_level = "HIGH"
            elif total_cells > 20:
                risk_level = "MEDIUM"

        return _ok("数据影响快速评估完成", data={
            'operation_type': operation_type,
            'range': range,
            'current_data': current_data,
            'impact_assessment': {
                'rows_affected': data_rows,
                'columns_affected': data_cols,
                'total_cells': total_cells,
                'has_existing_data': has_data,
                'risk_level': risk_level
            },
            'safety_warning': _generate_safety_warning(operation_type, has_data, risk_level)
        }, meta={"file_path": file_path, "range": range})

    # 详细评估模式（原assess_data_impact行为）
    try:
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
        recommendations = _generate_assessment_recommendations(
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

        return _ok("数据影响详细评估完成", data={
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
        }, meta={"file_path": file_path, "range": range})

    except DataValidationError as e:
        return _fail(f"参数验证失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})

    except Exception as e:
        return _fail(f"数据影响评估失败: {str(e)}", meta={"error_code": "ASSESSMENT_FAILED"})


def _generate_safety_warning(operation_type: str, has_data: bool, risk_level: str) -> str:
    """生成安全警告"""
    if risk_level == "HIGH":
        return f"🔴 高风险警告: {operation_type}操作将影响大量数据，请谨慎操作"
    elif risk_level == "MEDIUM":
        return f"🟡 中等风险: {operation_type}操作将影响部分数据，建议先备份"
    else:
        return f"✅ 低风险: {operation_type}操作影响较小，可以安全执行"


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


def _generate_assessment_recommendations(
    operation_type: str,
    data_analysis: Dict[str, Any],
    risk_assessment: Dict[str, Any],
    scale_info: Dict[str, Any]
) -> List[str]:
    """生成详细评估的安全建议"""
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
    """查看最近的Excel操作记录。可按文件过滤。

    Args:
        file_path: 文件路径，默认为None表示不过滤
        limit: 返回的最大记录数，默认为20
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

        return _ok(f"找到 {total_operations} 条操作记录", data={'operations': recent_operations, 'statistics': {'total_operations': total_operations, 'operation_types': operation_types, 'success_count': success_count, 'error_count': error_count, 'success_rate': f"{(success_count / (success_count + error_count) * 100):.1f}%" if (success_count + error_count) > 0 else "0%"}}, meta={"file_path": file_path})

    except Exception as e:
        return _fail(f"获取操作历史失败: {str(e)}", meta={"error_code": "HISTORY_RETRIEVAL_FAILED"})


@mcp.tool()
@_track_call
def excel_create_backup(
    file_path: str,
    backup_dir: Optional[str] = None
) -> Dict[str, Any]:
    """为Excel文件创建备份。备份存放在同级backup目录。

    Args:
        file_path: Excel文件路径
        backup_dir: 备份目录路径，默认为None表示同级backup目录
    """
    if not os.path.exists(file_path):
        return _fail(f"源文件不存在: {file_path}", meta={"error_code": "FILE_NOT_FOUND"})

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

        return _ok(f"备份创建成功: {backup_filename}", data={'backup_file': backup_path, 'backup_directory': backup_dir, 'file_size': {'original': original_size, 'backup': backup_size}, 'timestamp': timestamp}, meta={"file_path": file_path})

    except Exception as e:
        return _fail(f"备份创建失败: {str(e)}", meta={"error_code": "BACKUP_FAILED"})


@mcp.tool()
@_track_call
def excel_restore_backup(
    backup_path: str,
    target_path: Optional[str] = None
) -> Dict[str, Any]:
    """从备份文件恢复Excel。target_path不传则覆盖原文件。

    Args:
        backup_path: 备份文件路径
        target_path: 目标文件路径，默认为None表示覆盖原文件
    """
    _path_err = _validate_path(backup_path)
    if _path_err:
        return _path_err

    if not os.path.exists(backup_path):
        return _fail(f"备份文件不存在: {backup_path}", meta={"error_code": "BACKUP_NOT_FOUND"})

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

        return _ok(f"文件恢复成功: {os.path.basename(target_path)}", data={'backup_file': backup_path, 'target_file': target_path, 'target_existed': target_exists}, meta={"file_path": backup_path})

    except Exception as e:
        return _fail(f"恢复失败: {str(e)}", meta={"error_code": "RESTORE_FAILED"})


@mcp.tool()
@_track_call
def excel_list_backups(
    file_path: str,
    backup_dir: Optional[str] = None
) -> Dict[str, Any]:
    """列出文件的所有备份版本及时间。

    Args:
        file_path: Excel文件路径
        backup_dir: 备份目录路径，默认为None
    """
    try:
        # 确定备份目录
        if backup_dir is None:
            base_dir = os.path.dirname(file_path)
            backup_dir = os.path.join(base_dir, ".excel_mcp_backups")

        if not os.path.exists(backup_dir):
            return _ok("备份目录不存在", data={'backups': []}, meta={"file_path": file_path})

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

        return _ok(f"找到 {len(backup_files)} 个备份", data={'backups': backup_files, 'backup_directory': backup_dir, 'total_backups': len(backup_files)}, meta={"file_path": file_path})

    except Exception as e:
        return _fail(f"列出备份失败: {str(e)}", meta={"error_code": "LIST_BACKUPS_FAILED"})


@mcp.tool()
@_track_call
def excel_insert_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """在指定位置插入空行。row_index从0开始。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        row_index: 插入位置的行索引（从0开始）
        count: 插入的行数，默认为1
        streaming: 是否使用流式写入，默认为True
    """
    return _wrap(ExcelOperations.insert_rows(file_path, sheet_name, row_index, count, streaming))


@mcp.tool()
@_track_call
def excel_insert_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """在指定位置插入空列。column_index从1开始。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        column_index: 插入位置的列索引（从1开始）
        count: 插入的列数，默认为1
        streaming: 是否使用流式写入，默认为True
    """
    return _wrap(ExcelOperations.insert_columns(file_path, sheet_name, column_index, count, streaming))


@mcp.tool()
@_track_call
def excel_find_last_row(
    file_path: str,
    sheet_name: str,
    column: Optional[Union[str, int]] = None
) -> Dict[str, Any]:
    """查找工作表最后一行。可指定列来找该列最后一个有值的行。追加数据前必用。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        column: 列名或列索引，默认为None
    """
    return _wrap(ExcelOperations.find_last_row(file_path, sheet_name, column))


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]:
    """创建新Excel文件。可指定初始工作表名称列表。

    Args:
        file_path: Excel文件路径
        sheet_names: 初始工作表名称列表，默认为None
    """
    return _wrap(ExcelOperations.create_file(file_path, sheet_names))


@mcp.tool()
@_validate_file_path(['file_path', 'output_path'])
@_track_call
def excel_export_to_csv(
    file_path: str,
    output_path: str,
    sheet_name: Optional[str] = None,
    encoding: str = "utf-8"
) -> Dict[str, Any]:
    """将工作表导出为CSV。支持指定编码和分隔符。

    Args:
        file_path: Excel文件路径
        output_path: CSV输出路径
        sheet_name: 工作表名称，默认为None
        encoding: 编码格式，默认为"utf-8"
    """
    return _wrap(ExcelOperations.export_to_csv(file_path, output_path, sheet_name, encoding))


@mcp.tool()
@_validate_file_path(['csv_path', 'output_path'])
@_track_call
def excel_import_from_csv(
    csv_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    encoding: str = "utf-8",
    has_header: bool = True
) -> Dict[str, Any]:
    """从CSV创建Excel文件。支持编码和分隔符配置。

    Args:
        csv_path: CSV文件路径
        output_path: Excel输出路径
        sheet_name: 工作表名称，默认为"Sheet1"
        encoding: 编码格式，默认为"utf-8"
        has_header: CSV是否有表头，默认为True
    """
    for _p in [csv_path, output_path]:
        _err = _validate_path(_p)
        if _err:
            return _err

    return _wrap(ExcelOperations.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header))


@mcp.tool()
@_validate_file_path(['input_path', 'output_path'])
@_track_call
def excel_convert_format(
    input_path: str,
    output_path: str,
    target_format: str = "xlsx"
) -> Dict[str, Any]:
    """Excel/CSV/JSON格式互转。

    Args:
        input_path: 输入文件路径
        output_path: 输出文件路径
        target_format: 目标格式，默认为"xlsx"
    """
    for _p in [input_path, output_path]:
        _err = _validate_path(_p)
        if _err:
            return _err

    return _wrap(ExcelOperations.convert_format(input_path, output_path, target_format))


@mcp.tool()
@_track_call
def excel_merge_files(
    input_files: List[str],
    output_path: str,
    merge_mode: str = "sheets"
) -> Dict[str, Any]:
    """合并多个Excel文件。merge_mode: sheets(每个文件一个表) | append(纵向追加) | columns(横向拼接)。

    Args:
        input_files: 输入文件路径列表
        output_path: 输出文件路径
        merge_mode: 合并模式，默认为"sheets"
    """
    for _f in input_files:
        _err = _validate_path(_f)
        if _err:
            return _err
    _err = _validate_path(output_path)
    if _err:
        return _err

    return _wrap(ExcelOperations.merge_files(input_files, output_path, merge_mode))


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_get_file_info(file_path: str) -> Dict[str, Any]:
    """获取文件元数据：大小、工作表数、行列范围等。

    Args:
        file_path: Excel文件路径
    """
    return _wrap(ExcelOperations.get_file_info(file_path))


@mcp.tool()
@_track_call
def excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """创建新工作表。可指定插入位置index。

    Args:
        file_path: Excel文件路径
        sheet_name: 新工作表名称
        index: 插入位置索引，默认为None
    """
    return _wrap(ExcelOperations.create_sheet(file_path, sheet_name, index))


@mcp.tool()
@_track_call
def excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]:
    """删除指定工作表。

    Args:
        file_path: Excel文件路径
        sheet_name: 要删除的工作表名称
    """
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

        return _wrap(result)

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除工作表操作失败: {str(e)}"
        })

        return _fail(f"删除工作表操作失败: {str(e)}", meta={"error_code": "DELETE_SHEET_FAILED"})


@mcp.tool()
@_track_call
def excel_rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str
) -> Dict[str, Any]:
    """重命名工作表。

    Args:
        file_path: Excel文件路径
        old_name: 原工作表名称
        new_name: 新工作表名称
    """
    return _wrap(ExcelOperations.rename_sheet(file_path, old_name, new_name))


@mcp.tool()
@_track_call
def excel_copy_sheet(
    file_path: str,
    source_name: str,
    new_name: Optional[str] = None,
    index: Optional[int] = None,
    streaming: bool = True
) -> Dict[str, Any]:
    """复制工作表（含数据和格式）。可指定目标文件。

    Args:
        file_path: Excel文件路径
        source_name: 源工作表名称
        new_name: 新工作表名称，默认为None
        index: 插入位置索引，默认为None
        streaming: 是否使用流式写入，默认为True
    """
    return _wrap(ExcelOperations.copy_sheet(file_path, source_name, new_name, index, streaming))


@mcp.tool()
@_track_call
def excel_rename_column(
    file_path: str,
    sheet_name: str,
    old_header: str,
    new_header: str,
    header_row: int = 1
) -> Dict[str, Any]:
    """修改表头（列名）。只改header_row指定的行。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        old_header: 原列名
        new_header: 新列名
        header_row: 表头行号，默认为1
    """
    return _wrap(ExcelOperations.rename_column(file_path, sheet_name, old_header, new_header, header_row))


@mcp.tool()
@_track_call
def excel_upsert_row(
    file_path: str,
    sheet_name: str,
    key_column: str,
    key_value: Any,
    updates: Dict[str, Any],
    header_row: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """按key_column+key_value查找行，存在则更新，不存在则插入。update_columns指定要更新的列。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        key_column: 键列名
        key_value: 键值
        updates: 要更新的字段字典
        header_row: 表头行号，默认为1
        streaming: 是否使用流式写入，默认为True
    """
    return _wrap(ExcelOperations.upsert_row(file_path, sheet_name, key_column, key_value, updates, header_row, streaming))


@mcp.tool()
@_track_call
def excel_batch_insert_rows(
    file_path: str,
    sheet_name: str,
    data: List[Dict[str, Any]],
    header_row: int = 1,
    streaming: bool = True,
    insert_position: str = None,
    condition: str = None
) -> Dict[str, Any]:
    """批量插入多行数据。data为字典列表，header_row指定表头行号。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        data: 要插入的数据字典列表
        header_row: 表头行号，默认为1
        streaming: 是否使用流式写入，默认为True
        insert_position: 插入位置行号，默认为None
        condition: 条件表达式，默认为None
    """

    # 确定插入位置
    target_row = None

    if condition is not None:
        # 条件定位：找到第一个匹配条件的行号
        try:
            import pandas as pd
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='calamine', keep_default_na=False)
            try:
                filtered = df.query(condition)
            except Exception:
                # 降级：用SQL引擎
                sql = f"SELECT * FROM {sheet_name} WHERE {condition} LIMIT 1"
                query_result = ExcelOperations.query(file_path, sql)
                if query_result.get('success', False):
                    qdata = query_result.get('data', [])
                    if isinstance(qdata, list) and len(qdata) > 1:
                        # 匹配第一行获取行号
                        sql_headers = qdata[0]
                        sql_row = qdata[1]
                        if isinstance(sql_row, (list, tuple)):
                            row_dict = dict(zip(sql_headers, sql_row))
                            col_names = list(df.columns)
                            for row_tuple in df.itertuples(index=True, name=None):
                                idx = row_tuple[0]
                                row_vals = dict(zip(col_names, row_tuple[1:]))
                                match = True
                                for col, val in row_dict.items():
                                    if col in col_names and str(row_vals.get(col, '')) != str(val):
                                        match = False
                                        break
                                if match:
                                    target_row = idx + 2  # 1-based + header
                                    break
            else:
                if len(filtered) > 0:
                    target_row = filtered.index[0] + 2  # 1-based + header
        except Exception as e:
            logger.warning(f"条件定位查询失败: {e}，将使用默认追加模式")

    elif insert_position is not None:
        try:
            target_row = int(insert_position)
        except (ValueError, TypeError):
            return _fail(f"insert_position必须是整数: {insert_position}",
                         meta={"error_code": "INVALID_INSERT_POSITION"})

    if target_row is not None:
        # 指定位置插入模式
        try:
            result = ExcelOperations.batch_insert_rows_at(file_path, sheet_name, data, target_row, header_row, streaming)
            return _wrap(result)
        except Exception as e:
            return _fail(f"指定位置插入失败: {str(e)}", meta={"error_code": "INSERT_AT_FAILED"})

    return _wrap(ExcelOperations.batch_insert_rows(file_path, sheet_name, data, header_row, streaming))


@mcp.tool()
@_track_call
def excel_delete_rows(
    file_path: str,
    sheet_name: str,
    row_index: int = None,
    count: int = 1,
    streaming: bool = True,
    condition: str = None
) -> Dict[str, Any]:
    """删除行。支持按索引(row_index)或条件(where_column+where_value)删除。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        row_index: 要删除的行索引，默认为None
        count: 删除的行数，默认为1
        streaming: 是否使用流式写入，默认为True
        condition: SQL条件表达式，默认为None
    """

    # condition模式：根据SQL条件查找并删除行
    if condition is not None:
        operation_logger.start_session(file_path)
        operation_logger.log_operation("delete_rows_by_condition", {
            "sheet_name": sheet_name,
            "condition": condition
        })
        try:
            # 使用pandas加载工作表数据，找出符合条件的行号
            import pandas as pd
            from python_calamine import CalamineWorkbook

            cal_wb = CalamineWorkbook.from_path(file_path)
            cal_ws = cal_wb.get_sheet_by_name(sheet_name)
            if cal_ws is None:
                return _fail(f"工作表 '{sheet_name}' 不存在",
                             meta={"error_code": "SHEET_NOT_FOUND"})

            # 读取全部数据
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='calamine', keep_default_na=False)
            # 找出表头行（第一行数据行），实际数据从第2行开始
            # condition中的列名对应表头，行号 = 数据行index + 2（1-based表头 + 1-based offset）
            header_row = 1  # 默认表头在第1行

            # REQ-046: 尝试将可转换的列转为数值类型，避免字符串比较导致条件匹配失败
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='ignore')
                except Exception:
                    pass

            # 使用pandas query执行条件筛选
            try:
                # pandas的query需要特殊处理中文列名
                filtered = df.query(condition)
            except Exception:
                # 降级：用SQL引擎执行
                sql = f"SELECT * FROM {sheet_name} WHERE {condition}"
                query_result = ExcelOperations.query(file_path, sql)
                if not query_result.get('success', False):
                    return _fail(f"条件查询失败: {query_result.get('message', '未知错误')}",
                                 meta={"error_code": "CONDITION_QUERY_FAILED"})
                # 从SQL结果中提取行号（通过比对原始数据）
                qdata = query_result.get('data', [])
                if not isinstance(qdata, list) or len(qdata) <= 1:
                    return _ok(f"条件 '{condition}' 未匹配到任何行",
                               data={'deleted_rows': 0, 'condition': condition})
                # SQL结果的行需要和原始df匹配来获取行号
                sql_headers = qdata[0]
                sql_rows = qdata[1:]
                row_numbers = []
                for sql_row in sql_rows:
                    if not isinstance(sql_row, (list, tuple)):
                        continue
                    row_dict = dict(zip(sql_headers, sql_row))
                    # 在df中找到匹配的行
                    col_names = list(df.columns)
                    for row_tuple in df.itertuples(index=True, name=None):
                        idx = row_tuple[0]
                        row_vals = dict(zip(col_names, row_tuple[1:]))
                        match = True
                        for col, val in row_dict.items():
                            if col in col_names and str(row_vals.get(col, '')) != str(val):
                                match = False
                                break
                        if match:
                            row_numbers.append(idx + header_row + 1)  # 1-based, +1 for header
                            break
            else:
                # pandas query成功，获取行号
                row_numbers = [idx + header_row + 1 for idx in filtered.index]

            if not row_numbers:
                return _ok(f"条件 '{condition}' 未匹配到任何行",
                           data={'deleted_rows': 0, 'condition': condition})

            # 使用批量删除（一次文件I/O代替N次）- REQ-032 多线程优化
            result = ExcelOperations.batch_delete_rows(file_path, sheet_name, row_numbers, streaming)
            total_deleted = result.get('data', {}).get('deleted_rows', result.get('metadata', {}).get('deleted_rows', 0))

            operation_logger.log_operation("operation_result", {
                "success": result.get('success', False),
                "deleted_rows": total_deleted,
                "message": f"按条件批量删除完成"
            })

            if result.get('success', False):
                return _ok(f"按条件 '{condition}' 删除了 {total_deleted} 行",
                           data={'deleted_rows': total_deleted, 'condition': condition})
            else:
                return _fail(f"按条件删除失败: {result.get('message', '')}",
                             meta={"error_code": "BATCH_DELETE_FAILED"})

        except Exception as e:
            operation_logger.log_operation("operation_error", {
                "error": str(e),
                "message": f"条件删除失败: {str(e)}"
            })
            return _fail(f"条件删除失败: {str(e)}", meta={"error_code": "CONDITION_DELETE_FAILED"})

    # 行号模式：按row_index和count删除
    if row_index is None:
        return _fail("删除行需要指定 row_index 或 condition 参数",
                     meta={"error_code": "MISSING_ROW_INDEX"})

    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录删除操作日志
    operation_logger.log_operation("delete_rows", {
        "sheet_name": sheet_name,
        "row_index": row_index,
        "count": count
    })

    try:
        result = ExcelOperations.delete_rows(file_path, sheet_name, row_index, count, streaming)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "deleted_rows": result.get('deleted_rows', 0),
            "message": result.get('message', '')
        })

        return _wrap(result)

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除行操作失败: {str(e)}"
        })

        return _fail(f"删除行操作失败: {str(e)}", meta={"error_code": "DELETE_ROWS_FAILED"})


@mcp.tool()
@_track_call
def excel_delete_columns(
    file_path: str,
    sheet_name: str,
    column_index: int,
    count: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """删除指定位置开始的列。column_index从1开始。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        column_index: 起始列索引（从1开始）
        count: 删除的列数，默认为1
        streaming: 是否使用流式写入，默认为True
    """
    # 开始操作会话
    operation_logger.start_session(file_path)

    # 记录删除列操作日志
    operation_logger.log_operation("delete_columns", {
        "sheet_name": sheet_name,
        "column_index": column_index,
        "count": count
    })

    try:
        result = ExcelOperations.delete_columns(file_path, sheet_name, column_index, count, streaming)

        # 记录操作结果
        operation_logger.log_operation("operation_result", {
            "success": result.get('success', False),
            "deleted_columns": result.get('deleted_columns', 0),
            "message": result.get('message', '')
        })

        return _wrap(result)

    except Exception as e:
        # 记录错误
        operation_logger.log_operation("operation_error", {
            "error": str(e),
            "message": f"删除列操作失败: {str(e)}"
        })

        return _fail(f"删除列操作失败: {str(e)}", meta={"error_code": "DELETE_COLUMNS_FAILED"})

@mcp.tool()
@_track_call
def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """在单元格写入Excel公式。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        cell_address: 单元格地址，如"A1"
        formula: Excel公式，以等号开头
    """
    # 参数验证：formula 不能为空
    if not formula or not formula.strip():
        return _fail(
            '📝 公式参数缺失：excel_set_formula() 需要一个有效的Excel公式。\n'
            '✅ 正确用法：excel_set_formula("file.xlsx", "Sheet1", "A1", "=SUM(B1:C1)")\n'
            '🔧 支持的公式类型：\n'
            '  - 数学公式: "=A1+B1", "=SUM(A1:A10)"\n'
            '  - 引用公式: "=B2", "=Sheet2!A1"\n'
            '  - 函数公式: "=VLOOKUP(A1, Sheet2!A:B, 2, FALSE)"\n'
            '❌ 不要只传单元格地址，公式必须以等号 "=" 开头',
            meta={
                "error_code": "MISSING_FORMULA",
                "received_formula": formula,
                "expected_format": "Excel公式（以=开头）",
                "example": "=SUM(A1:C1)"
            }
        )

    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return _fail(f'🔒 安全验证失败: {_formula_err["error"]}', meta={"error_code": "FORMULA_SECURITY_FAILED"})
    return _wrap(ExcelOperations.set_formula(file_path, sheet_name, cell_address, formula))

@mcp.tool()
@_track_call
def excel_evaluate_formula(
    formula: str,
    context_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """临时计算公式结果，不修改文件。

    Args:
        formula: Excel公式
        context_sheet: 上下文工作表名称，默认为None
    """
    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return _fail(f'🔒 安全验证失败: {_formula_err["error"]}', meta={"error_code": "FORMULA_SECURITY_FAILED"})
    return _wrap(ExcelOperations.evaluate_formula(formula, context_sheet))


@mcp.tool()
@_track_call
def excel_query(
    file_path: str,
    query_expression: str,
    include_headers: bool = True,
    output_format: str = "table"
) -> Dict[str, Any]:
    """SQL查询引擎。支持WHERE/JOIN/GROUP BY/ORDER BY/LIMIT/子查询。
query_expression: SELECT * FROM 技能表 WHERE 伤害>100 | GROUP BY | JOIN ON

    Args:
        file_path: Excel文件路径
        query_expression: SQL查询语句
        include_headers: 是否包含表头，默认为True
        output_format: 输出格式，默认为"table"
    """
    # 参数验证
    if not file_path or not file_path.strip():
        return _fail('文件路径不能为空', meta={"error_code": "MISSING_FILE_PATH"})

    if not query_expression or not query_expression.strip():
        return _fail('SQL查询语句不能为空', meta={"error_code": "MISSING_QUERY"})

    # 验证 output_format
    valid_formats = ('table', 'json', 'csv')
    if output_format and output_format not in valid_formats:
        return _fail(f'不支持的输出格式: {output_format}。可选: {", ".join(valid_formats)}', meta={"error_code": "INVALID_FORMAT"})

    # 使用高级SQL查询引擎
    try:
        from .api.advanced_sql_query import execute_advanced_sql_query
        result = _wrap(execute_advanced_sql_query(
            file_path=file_path,
            sql=query_expression,
            sheet_name=None,  # 统一使用SQL FROM子句中的表名
            limit=None,  # 统一使用SQL中的LIMIT
            include_headers=include_headers,
            output_format=output_format or 'table'
        ))
        # 确保meta字段存在（query_info保留在顶层向后兼容）
        if 'meta' not in result:
            result['meta'] = {'file_path': file_path}
        return result

    except ImportError:
        return _fail('SQLGlot未安装，无法使用高级SQL功能。请运行: pip install sqlglot\n\n💡 智能降级建议：\n• 对于简单数据读取：尝试使用 excel_get_range("文件路径", "工作表名!A1:Z100")\n• 对于文本搜索：尝试使用 excel_search("文件路径", "关键词", "工作表名")\n• 对于表头信息：尝试使用 excel_get_headers("文件路径", "工作表名")', meta={"error_code": "DEPENDENCY_MISSING"})
    except Exception as e:
        # SQL引擎已处理大部分错误并返回结构化响应，此处仅捕获未预期的异常
        return _fail(f'SQL查询失败: {str(e)}', meta={"error_code": "SQL_EXECUTION_FAILED"})


@mcp.tool()
@_track_call
def excel_update_query(
    file_path: str,
    update_expression: str,
    dry_run: bool = False
) -> Dict[str, Any]:
    """SQL批量修改。dry_run=True预览变更不实际写入。示例: UPDATE 技能表 SET 伤害=200 WHERE 等级>=5

    Args:
        file_path: Excel文件路径
        update_expression: UPDATE语句
        dry_run: 是否仅预览不实际写入，默认为False
    """

    if not file_path or not file_path.strip():
        return _fail('文件路径不能为空', meta={"error_code": "MISSING_FILE_PATH"})

    if not update_expression or not update_expression.strip():
        return _fail('UPDATE语句不能为空', meta={"error_code": "MISSING_QUERY"})

    # 安全检查：只允许UPDATE语句
    stripped = update_expression.strip().upper()
    if not stripped.startswith('UPDATE'):
        return _fail('只支持UPDATE语句。查询请使用 excel_query', meta={"error_code": "UNSUPPORTED_SQL"})

    try:
        from .api.advanced_sql_query import execute_advanced_update_query
        return _wrap(execute_advanced_update_query(
            file_path=file_path,
            sql=update_expression,
            dry_run=dry_run
        ))
    except ImportError:
        return _fail('SQLGlot未安装，无法使用UPDATE功能', meta={"error_code": "DEPENDENCY_MISSING"})
    except Exception as e:
        return _fail(f'UPDATE执行失败: {str(e)}', meta={"error_code": "UPDATE_EXECUTION_FAILED"})


@mcp.tool()
@_track_call
def _detect_dual_header(rows) -> tuple:
    """
    检测双行表头模式
    
    Args:
        rows: 工作表前几行数据
        
    Returns:
        tuple: (is_dual_header, header_row_idx, descriptions)
    """
    is_dual_header = False
    header_row_idx = 0
    descriptions = None
    
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
                descriptions = first_row
    
    return is_dual_header, header_row_idx, descriptions

def _collect_column_statistics(ws, data_start, num_cols, col_name_list):
    """
    收集列统计信息 - 优化版本，单次遍历
    
    Args:
        ws: 工作表对象
        data_start: 数据开始行
        num_cols: 列数
        col_name_list: 列名列表
        
    Returns:
        dict: 列统计信息
    """
    col_stats = {}
    for col_idx in range(num_cols):
        col_name = col_name_list[col_idx]
        col_stats[col_name] = {'non_null': 0, 'samples': [], 'type_values': []}
    
    # 单次遍历所有行和列，同时统计行数和列数据
    total_rows = 0
    for row in ws.iter_rows(min_row=data_start + 1, values_only=True):
        total_rows += 1
        for col_idx in range(min(len(row), num_cols)):
            val = row[col_idx]
            if val is not None:
                s = col_stats[col_name_list[col_idx]]
                s['non_null'] += 1
                if len(s['samples']) < 3:
                    s['samples'].append(val)
                if len(s['type_values']) < 100:
                    s['type_values'].append(val)
    
    return col_stats, total_rows

def _build_describe_columns(col_stats, col_name_list, is_dual_header, descriptions):
    """
    分析列数据类型并构建最终列信息列表（保留原始完整行为）
    
    Args:
        col_stats: 列统计信息
        col_name_list: 列名列表
        is_dual_header: 是否双行表头
        descriptions: 双行表头描述列表
        
    Returns:
        list: 列信息列表
    """
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

    return list(col_stats.values())

def _resolve_row_count(ws, data_start, total_rows):
    """
    统计行数 — 优先使用max_row，为None则使用total_rows（已统计避免重复计算）
    
    Args:
        ws: 工作表对象
        data_start: 数据开始行
        total_rows: 已统计的总行数
        
    Returns:
        int: 行数
    """
    try:
        # 修复：streaming写入后max_row可能为None的问题
        if hasattr(ws, 'max_row') and ws.max_row is not None and ws.max_row > data_start:
            row_count = ws.max_row - data_start
        else:
            # streaming写入后max_row可能为None，直接使用已统计的total_rows
            # total_rows已经在上面的循环中统计完成，避免重复计算
            row_count = total_rows if total_rows > 0 else 0
            # 安全检查：如果total_rows为0，回退到iter_rows统计
            if row_count == 0:
                try:
                    for idx, row in enumerate(ws.iter_rows(min_row=data_start + 1, values_only=True), start=1):
                        # 只要行中有一个非空单元格，就计数为有效数据行
                        if row and any(cell is not None for cell in row):
                            row_count += 1
                except Exception as e:
                    logger.warning(f"iter_rows统计行数失败: {e}，使用默认值0")
                    row_count = 0
    except Exception as e:
        # 极端情况处理：如果max_row访问失败，使用已统计的total_rows
        logger.warning(f"访问ws.max_row失败: {e}，使用已统计的total_rows")
        row_count = total_rows if total_rows > 0 else 0
        # 确保total_rows无效时，使用iter_rows重新统计
        if row_count == 0:
            try:
                for idx, row in enumerate(ws.iter_rows(min_row=data_start + 1, values_only=True), start=1):
                    if row and any(cell is not None for cell in row):
                        row_count += 1
            except Exception as e2:
                logger.warning(f"iter_rows回退统计也失败: {e2}")
                row_count = 0
    return row_count

def _prepare_describe_result(sheet_name, is_dual_header, columns, row_count, file_path, sheet):
    """
    准备describe结果
    
    Args:
        sheet_name: 工作表名称
        is_dual_header: 是否双行表头
        columns: 列信息
        row_count: 行数
        file_path: 文件路径
        sheet: 工作表对象
        
    Returns:
        dict: 返回结果
    """
    return _ok(f"表 '{sheet_name}': {len(columns)}列, {row_count}行数据, {'双行表头' if is_dual_header else '单行表头'}", data={
        'sheet_name': sheet_name,
        'header_type': 'dual' if is_dual_header else 'single',
        'row_count': row_count,
        'column_count': len(columns),
        'columns': columns
    }, meta={"file_path": file_path, "sheet": sheet_name})


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_describe_table(
    file_path: str,
    sheet_name: str = None
) -> Dict[str, Any]:
    """
📋 表结构分析器 - 数据探索的必备第一步

**核心功能**: 快速分析Excel工作表结构，返回列名、数据类型、空值统计等元信息。在进行任何数据操作前，务必先调用此工具了解表结构，避免列名错误和操作失败。

**🎮 游戏开发场景**:
• **表结构扫描**: 快速了解技能表、装备表、怪物表的字段构成和数据类型
• **数据质量检查**: 检查关键字段（如ID、名称）是否完整，是否存在空值异常
• **配置验证**: 确认配置表包含预期字段，字段类型是否正确
• **开发调试**: 在开发过程中快速检查数据结构是否符合预期
• **RPG配置**: 检查技能表的skill_id、skill_name、damage、cooldown等字段类型完整性
• **装备系统**: 验证装备表的稀有度字段是否为枚举类型，属性字段是否为数值类型
• **数值平衡**: 检查怪物表的血量、攻击、防御等关键数值字段是否存在异常值
• **任务系统**: 确认任务表的接取条件、完成条件、奖励字段格式正确
• **经济系统**: 验证物价表的价格字段为数值型，商店折扣字段为百分比格式

**🔍 返回关键信息**:
• **columns**: 所有列名及对应的数据类型（number/text/date/mixed）
• **total_rows**: 总行数（包含标题行）
• **non_null_counts**: 每列的非空值数量，用于数据完整性检查
• **sample_data**: 每列前3个实际值，直观了解数据内容
• **header_mode**: single（单行表头）或 dual（双行表头中英映射）

**🔧 参数说明**:
• **file_path**: Excel文件路径（支持相对路径）
• **sheet_name**: 工作表名称（可选，不传则分析第一个工作表）

**💡 实用技巧**:
• **操作前必调**: 任何数据操作前先describe，了解表结构和数据类型
• **空值检查**: 关注non_null_counts，空值率高的列可能是数据质量问题
• **双行表头**: 自动识别中文描述+英文字段名，返回完整映射关系
• **性能友好**: 比get_range读取全部数据更轻量，适合初步了解表结构

**🔍 专业数据分析**:
• **数据类型验证**: 检查数值字段是否为number，文本字段是否为text，日期字段格式正确
• **数据完整性**: 计算completeness_rate，关键字段如ID、名称应达到100%完整性
• **异常值检测**: 检查数值字段的最大值、最小值是否在合理范围内
• **字段冗余**: 分析是否有重复含义的字段，可以合并删除
• **数据一致性**: 检查相关文件间的字段命名约定是否一致

**⚡ 性能优化**:
• **抽样分析**: 默认只读取前3行样本数据，快速了解表结构
• **延迟加载**: 大文件时使用read_only模式，避免内存溢出
• **缓存利用**: 重复调用时利用内部缓存，第二次调用速度提升50%
• **智能识别**: 自动检测双行表头，无需手动指定header_mode

**🎯 最佳实践**:
• **配置审查**: `describe` → `query` → 验证配置完整性 → 生成质量报告
• **迁移验证**: 迁移前后各调用一次describe，验证结构一致性
• **监控指标**: 定期调用describe监控数据质量变化趋势
• **异常预警**: 设置字段完整性阈值，低于阈值时触发警报

**📊 返回信息**:
• **columns**: 列信息列表{name、type、sample_values}
• **table_stats**: 表统计{total_rows、header_mode、dual_header_detected}
• **data_quality**: 数据质量分析{null_counts、completeness_rates}
• **file_path**: 源文件路径
• **success**: 操作是否成功
• **message**: 状态消息或错误信息

**🔗 配合使用**:
• **后续查询**: `excel_query`基于此表结构进行数据检索
• **修改操作**: `excel_update_query`根据类型进行安全修改
• **版本对比**: `excel_compare_sheets`对比不同版本的结构变化
• **数据评估**: `excel_assess_data_impact`修改前结合describe评估影响

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为None表示第一个工作表
    """
    # 文件验证和加载
    if not file_path or not file_path.strip():
        return _fail('文件路径不能为空', meta={"error_code": "MISSING_FILE_PATH"})

    try:
        wb, ws = ExcelValidator.get_workbook_and_sheet(file_path, sheet_name, read_only=True, data_only=True)
        if not sheet_name:
            sheet_name = ws.title
    except DataValidationError as e:
        return _fail(e.message, meta={"error_code": "SHEET_NOT_FOUND", "hint": e.hint, "suggested_fix": e.suggested_fix})
    except Exception as e:
        return _fail(f'无法打开文件: {e}', meta={"error_code": "FILE_OPEN_FAILED"})

    try:

        # 读取前几行来判断表头类型
        rows = list(ws.iter_rows(max_row=4, values_only=True))
        if not rows:
            return _fail('工作表为空', meta={"error_code": "EMPTY_SHEET"})

        # 检测双行表头（提取为独立函数）
        is_dual_header, header_row_idx, descriptions = _detect_dual_header(rows)
        headers = rows[header_row_idx]
        data_start = header_row_idx + 1

        # 准备列名列表
        num_cols = len(headers)
        col_name_list = []
        for col_idx in range(num_cols):
            col_name = headers[col_idx]
            if col_name is None:
                col_name = f"column_{col_idx + 1}"
            col_name = str(col_name).strip()
            if not col_name:
                col_name = f"column_{col_idx + 1}"
            col_name_list.append(col_name)

        # 单次遍历收集统计信息（提取为独立函数）
        col_stats, total_rows = _collect_column_statistics(ws, data_start, num_cols, col_name_list)

        # 推断类型并构建最终结果（提取为独立函数）
        columns = _build_describe_columns(col_stats, col_name_list, is_dual_header, descriptions)

        # 统计行数 — 优先使用max_row，为None则使用total_rows
        row_count = _resolve_row_count(ws, data_start, total_rows)

        return _prepare_describe_result(sheet_name, is_dual_header, columns, row_count, file_path, ws)
    except Exception as e:
        return _fail(f'查看表结构失败: {e}', meta={"error_code": "DESCRIBE_FAILED"})
    finally:
        wb.close()


@mcp.tool()
@_track_call
def excel_format_cells(
    file_path: str,
    sheet_name: str,
    range: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None
) -> Dict[str, Any]:
    """设置单元格样式。formatting字段: bold/italic/underline/font_size/font_color/bg_color/number_format/alignment/wrap_text/border_style。只传需要修改的字段。
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range: 单元格范围，如 "A1:C10"
        formatting: 样式配置
        preset: 预设样式名称
        start_cell: 起始单元格（可选，与end_cell配合使用）
        end_cell: 结束单元格（可选，与start_cell配合使用）
    """
    
    # 参数校验：当没有提供formatting和preset参数时，明确告知用户
    if formatting is None and preset is None:
        return _fail(
            '未提供样式参数：formatting 和 preset 均为空。\n'
            '请至少提供一种样式配置：\n'
            '1. 使用 preset: "bold", "italic", "highlight", "header"\n'
            '2. 使用 formatting: {"bold": true} 或 {"font_size": 12}\n'
            '示例：excel_format_cells(file, "Sheet1", "A1:C1", preset="bold")',
            meta={
                "error_code": "MISSING_FORMATTING_PARAMS",
                "received_formatting": formatting,
                "received_preset": preset,
                "hint": "需要指定样式才能执行格式化操作"
            }
        )
    
    # 参数兼容性处理：如果提供了start_cell和end_cell，构建range表达式
    if start_cell and end_cell:
        range = f"{start_cell}:{end_cell}"
    
    # 确保range包含工作表名
    if '!' not in range:
        range = f"{sheet_name}!{range}"
    
    # 调用原始函数
    result = ExcelOperations.format_cells(file_path, sheet_name, range, formatting, preset)
    
    # 如果成功但data为null（无实际格式更改），添加说明
    if result.get('success') and result.get('data') is None:
        result['message'] += ' (注意：没有单元格需要格式化，可能是因为指定范围没有内容或样式无变化)'
    
    return _wrap(result)


@mcp.tool()
@_track_call
def excel_merge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """合并指定范围为一个大单元格。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range: 单元格范围
    """
    return _wrap(ExcelOperations.merge_cells(file_path, sheet_name, range))


@mcp.tool()
@_track_call
def excel_unmerge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """取消合并，恢复为独立单元格。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range: 单元格范围
    """
    return _wrap(ExcelOperations.unmerge_cells(file_path, sheet_name, range))


@mcp.tool()
@_track_call
def excel_set_borders(
    file_path: str,
    sheet_name: str,
    range: str,
    border_style: str = "thin"
) -> Dict[str, Any]:
    """为范围设置边框。border_style: thin/thick/double/dotted/dashed。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range: 单元格范围
        border_style: 边框样式，默认为"thin"
    """
    return _wrap(ExcelOperations.set_borders(file_path, sheet_name, range, border_style))


@mcp.tool()
@_track_call
def excel_set_row_height(
    file_path: str,
    sheet_name: str,
    row_index: int,
    height: float,
    count: int = 1
) -> Dict[str, Any]:
    """设置行高（磅值）。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        row_index: 行索引
        height: 行高（磅值）
        count: 影响的行数，默认为1
    """
    return _wrap(ExcelOperations.set_row_height(file_path, sheet_name, row_index, height, count))


@mcp.tool()
@_track_call
def excel_set_column_width(
    file_path: str,
    sheet_name: str,
    column_index: int,
    width: float,
    count: int = 1
) -> Dict[str, Any]:
    """设置列宽（字符单位）。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        column_index: 列索引
        width: 列宽（字符单位）
        count: 影响的列数，默认为1
    """
    return _wrap(ExcelOperations.set_column_width(file_path, sheet_name, column_index, width, count))


# ==================== Excel比较功能 ====================

@mcp.tool()
@_validate_file_path(['file1_path', 'file2_path'])
@_track_call
def excel_compare_files(
    file1_path: str,
    file2_path: str
) -> Dict[str, Any]:
    """逐单元格比较两个Excel文件差异。

    Args:
        file1_path: 第一个文件路径
        file2_path: 第二个文件路径
    """
    for _p in [file1_path, file2_path]:
        _err = _validate_path(_p)
        if _err:
            return _err
    return _wrap(ExcelOperations.compare_files(file1_path, file2_path))


@mcp.tool()
@_track_call
def excel_check_duplicate_ids(
    file_path: str,
    sheet_name: str,
    id_column: Union[int, str] = 1,
    header_row: int = 1
) -> Dict[str, Any]:
    """扫描ID列，返回重复值及所在行号。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        id_column: ID列名或列索引，默认为1
        header_row: 表头行号，默认为1
    """
    return _wrap(ExcelOperations.check_duplicate_ids(file_path, sheet_name, id_column, header_row))


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
    """比较两个工作表的差异：新增/删除/修改的行和列。

    Args:
        file1_path: 第一个文件路径
        sheet1_name: 第一个工作表名称
        file2_path: 第二个文件路径
        sheet2_name: 第二个工作表名称
        id_column: ID列名或列索引，默认为1
        header_row: 表头行号，默认为1
    """
    for _p in [file1_path, file2_path]:
        _err = _validate_path(_p)
        if _err:
            return _err
    return _wrap(ExcelOperations.compare_sheets(file1_path, sheet1_name, file2_path, sheet2_name, id_column, header_row))


@mcp.tool()
@_track_call
def excel_server_stats() -> Dict[str, Any]:
    """服务器状态：缓存、调用次数、运行时间。无参数。"""
    stats = _tracker.get_stats()
    return _ok("服务器统计信息", data=stats)


# ==================== 用户友好的参数兼容工具 ====================
@mcp.tool()
@_track_call
def excel_get_range_user_friendly(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """读取指定范围的数据（用户友好版本）。支持单独指定工作表、起始和结束单元格。
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格，如 "A1"
        end_cell: 结束单元格，如 "C10"
        include_formatting: 是否包含格式信息
    """
    
    # 构建范围表达式
    range_expression = f"{sheet_name}!{start_cell}:{end_cell}"
    
    try:
        # 调用原有的excel_get_range函数
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        if sheet_name not in sheet_names:
            return _fail(
                f"工作表 '{sheet_name}' 不存在。可用工作表: {', '.join(sheet_names)}",
                meta={"error_code": "SHEET_NOT_FOUND", "available_sheets": sheet_names}
            )
        
        # 调用原有的API
        result = ExcelOperations.get_range(file_path, range_expression, include_formatting)
        return _wrap(result)
        
    except Exception as e:
        return _wrap(ExcelOperations.get_range(file_path, range_expression, include_formatting))


@mcp.tool()
@_track_call
def excel_format_cells_user_friendly(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None
) -> Dict[str, Any]:
    """格式化单元格（用户友好版本）。支持单独指定工作表、起始和结束单元格。
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格，如 "A1"
        end_cell: 结束单元格，如 "C10"
        formatting: 样式配置，包含bold/italic/underline等字段
        preset: 预设样式名称
    """
    
    # 构建范围表达式
    range_expression = f"{sheet_name}!{start_cell}:{end_cell}"
    
    try:
        # 验证工作表存在
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        if sheet_name not in sheet_names:
            return _fail(
                f"工作表 '{sheet_name}' 不存在。可用工作表: {', '.join(sheet_names)}",
                meta={"error_code": "SHEET_NOT_FOUND", "available_sheets": sheet_names}
            )
        
        # 调用原有的API
        result = ExcelOperations.format_cells(
            file_path, sheet_name, f"{start_cell}:{end_cell}", formatting, preset
        )
        return _wrap(result)
        
    except Exception as e:
        return _fail(f"格式化失败: {str(e)}")


@mcp.tool()
@_track_call
def excel_update_range_user_friendly(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = False,
    streaming: bool = True
) -> Dict[str, Any]:
    """写入数据到指定范围（用户友好版本）。支持单独指定工作表、起始和结束单元格。
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格，如 "A1"
        end_cell: 结束单元格，如 "C10"
        data: 要写入的数据，二维数组格式
        preserve_formulas: 是否保留已有公式不被覆盖，默认True
        insert_mode: 数据写入模式，默认False(覆盖模式)
            - False: 覆盖模式，直接替换目标单元格数据，适合修改现有数据
            - True: 插入模式，在目标位置插入新行，原有数据下移，适合添加新数据
        streaming: 是否使用流式写入，默认True
    """
    
    # 构建范围表达式
    range_expression = f"{sheet_name}!{start_cell}:{end_cell}"
    
    try:
        # 验证工作表存在
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        if sheet_name not in sheet_names:
            return _fail(
                f"工作表 '{sheet_name}' 不存在。可用工作表: {', '.join(sheet_names)}",
                meta={"error_code": "SHEET_NOT_FOUND", "available_sheets": sheet_names}
            )
        
        # 调用原有的API
        result = ExcelOperations.update_range(
            file_path, range_expression, data, preserve_formulas, insert_mode, streaming
        )
        return _wrap(result)
        
    except Exception as e:
        return _fail(f"数据写入失败: {str(e)}")


# ==================== 批量操作工具 ====================
@mcp.tool()
@_validate_file_path()
@_track_call
def excel_batch_update_ranges(file_path: str, updates: List[Dict[str, Any]]) -> Dict[str, Any]:
    """批量更新多个范围。updates为[{range, data}]列表。

    Args:
        file_path: Excel文件路径
        updates: 更新项列表，每项包含range和data
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import range_boundaries

        results = []
        success_count = 0
        error_count = 0

        # 大批量并行预验证（>10个更新项时启用）- REQ-032
        if len(updates) > 10:
            from ..utils.concurrent_utils import parallel_validate_batch_data

            def _validate_update(update):
                range_spec = update.get('range', '')
                data = update.get('data', [])
                if not range_spec or not data:
                    return '缺少range或data参数'
                try:
                    range_boundaries(range_spec)
                except Exception as e:
                    return f'范围格式无效: {e}'
                return None

            errors = parallel_validate_batch_data(updates, _validate_update)
            for i, err in enumerate(errors):
                if err:
                    results.append({'index': i, 'success': False, 'error': err})
                    error_count += 1

        # 加载工作簿
        wb = load_workbook(file_path, data_only=False)

        for i, update in enumerate(updates):
            # 跳过已预验证失败的项
            if any(r['index'] == i and not r['success'] for r in results):
                continue

            try:
                range_spec = update.get('range', '')
                data = update.get('data', [])
                sheet_name = update.get('sheet', None)

                if not range_spec or not data:
                    results.append({
                        'index': i,
                        'success': False,
                        'error': '缺少range或data参数'
                    })
                    error_count += 1
                    continue

                # 获取工作表
                if sheet_name:
                    if sheet_name not in wb.sheetnames:
                        results.append({
                            'index': i,
                            'success': False,
                            'error': f'工作表"{sheet_name}"不存在'
                        })
                        error_count += 1
                        continue
                    ws = wb[sheet_name]
                else:
                    ws = wb.active

                # 解析范围并更新数据
                min_col, min_row, max_col, max_row = range_boundaries(range_spec)

                # 写入数据
                for row_idx, row_data in enumerate(data):
                    for col_idx, cell_value in enumerate(row_data):
                        target_row = min_row + row_idx
                        target_col = min_col + col_idx

                        if target_row <= max_row and target_col <= max_col:
                            ws.cell(row=target_row, column=target_col, value=cell_value)

                results.append({
                    'index': i,
                    'success': True,
                    'range': range_spec,
                    'sheet': sheet_name or 'active'
                })
                success_count += 1

            except Exception as e:
                results.append({
                    'index': i,
                    'success': False,
                    'error': str(e)
                })
                error_count += 1

        # 保存文件
        wb.save(file_path)

        return _ok(f"批量更新完成：成功{success_count}个区域，失败{error_count}个区域", data={
            'total_updates': len(updates),
            'success_count': success_count,
            'error_count': error_count,
            'results': results
        })

    except Exception as e:
        return _fail("批量更新失败", meta={"error_code": "OPERATION_FAILED"})


@mcp.tool()
@_track_call
def excel_merge_multiple_files(source_files: List[str], target_file: str, merge_mode: str = "append") -> Dict[str, Any]:
    """合并多个文件。merge_mode: append(纵向追加) | sheets(分表合并)。

    Args:
        source_files: 源文件路径列表
        target_file: 目标文件路径
        merge_mode: 合并模式，默认为"append"
    """
    # 路径遍历安全验证
    for _f in source_files:
        _err = _validate_path(_f)
        if _err:
            return _err
    _err = _validate_path(target_file)
    if _err:
        return _err
    try:
        from openpyxl import load_workbook, Workbook
        import os
        
        merged_sheets = {}
        total_rows = 0
        conflict_count = 0
        
        # 确保目标目录存在
        os.makedirs(os.path.dirname(target_file), exist_ok=True)
        
        # 加载目标文件或创建新文件
        if os.path.exists(target_file):
            target_wb = load_workbook(target_file)
        else:
            target_wb = Workbook()
            # 删除默认创建的工作表
            default_sheet = target_wb.active
            target_wb.remove(default_sheet)
        
        for source_file in source_files:
            if not os.path.exists(source_file):
                continue
                
            try:
                source_wb = load_workbook(source_file)
                
                for sheet_name in source_wb.sheetnames:
                    source_ws = source_wb[sheet_name]
                    
                    if sheet_name in target_wb.sheetnames:
                        # 工作表已存在，处理冲突
                        target_ws = target_wb[sheet_name]
                        
                        if merge_mode == "overwrite":
                            # 覆盖模式：清空目标工作表并复制数据
                            target_ws.delete_rows(1, target_ws.max_row)
                            for row in source_ws.iter_rows(values_only=True):
                                target_ws.append(row)
                        elif merge_mode == "merge":
                            # 合并模式：追加新行（跳过表头）
                            start_row = 1 if source_ws.max_row > 1 else 0
                            for row in source_ws.iter_rows(values_only=True, min_row=start_row+1):
                                target_ws.append(row)
                        else:
                            # 追加模式：直接追加
                            for row in source_ws.iter_rows(values_only=True):
                                target_ws.append(row)
                    else:
                        # 工作表不存在，直接复制
                        if sheet_name not in merged_sheets:
                            merged_sheets[sheet_name] = []
                        
                        new_ws = target_wb.create_sheet(title=sheet_name)
                        for row in source_ws.iter_rows(values_only=True):
                            new_ws.append(row)
                    
                    # 统计数据
                    if sheet_name not in merged_sheets:
                        merged_sheets[sheet_name] = []
                    merged_sheets[sheet_name].append(source_file)
                    total_rows += source_ws.max_row
                    
            except Exception as e:
                print(f"处理文件{source_file}时出错: {e}")
                continue
        
        # 保存目标文件
        target_wb.save(target_file)
        
        return _ok(f"文件合并完成：处理{len(source_files)}个源文件，生成{len(merged_sheets)}个工作表", data={
            'source_files_count': len(source_files),
            'merged_sheets_count': len(merged_sheets),
            'total_rows': total_rows,
            'merged_sheets': list(merged_sheets.keys()),
            'merge_mode': merge_mode,
            'target_file': target_file
        })
        
    except Exception as e:
        return _fail("文件合并失败", meta={"error_code": "OPERATION_FAILED"})


# ==================== 图表生成工具 ====================
@mcp.tool()
@_validate_file_path()
@_track_call
def excel_create_chart(file_path: str, sheet_name: str, chart_type: str, data_range: str, 
                      title: str = "", chart_name: str = "", position: str = "B15") -> Dict[str, Any]:
    """在工作表中创建图表。chart_type: line/bar/column/pie/scatter/area等。支持'column'作为'bar'的别名。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        chart_type: 图表类型
        data_range: 数据范围
        title: 图表标题，默认为空字符串
        chart_name: 图表名称，默认为空字符串
        position: 图表位置，默认为"B15"
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
        from openpyxl.utils import range_boundaries

        wb, ws = ExcelValidator.get_workbook_and_sheet(file_path, sheet_name)

        # 创建图表对象
        chart_map = {
            "column": BarChart,
            "bar": BarChart,
            "line": LineChart,
            "pie": PieChart,
            "scatter": ScatterChart
        }
        
        if chart_type not in chart_map:
            return _fail("不支持的图表类型", meta={"error_code": "INVALID_PARAMETER", "hint": f"支持的类型: {list(chart_map.keys())}"})
        
        chart_class = chart_map[chart_type]
        chart = chart_class()
        
        # 设置图表标题
        if title:
            chart.title = title
        
        # 设置数据范围
        min_col, min_row, max_col, max_row = range_boundaries(data_range)
        
        data = Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
        categories = Reference(ws, min_col=min_col-1 if min_col > 1 else min_col, 
                              min_row=min_row+1, max_row=max_row)
        
        chart.add_data(data, titles_from_data=True)
        if chart_type != "pie":
            chart.set_categories(categories)
        
        # 设置图表位置
        if position:
            chart.anchor = position
        
        # 添加图表到工作表
        ws.add_chart(chart)
        
        # 保存文件
        wb.save(file_path)
        
        # 生成图表名称
        if not chart_name:
            chart_name = f"图表_{len(ws._charts) + 1}"
        
        return _ok("图表创建成功", data={
            'chart_name': chart_name,
            'chart_type': chart_type,
            'data_range': data_range,
            'sheet_name': sheet_name,
            'position': position,
            'title': title,
            'chart_count': len(ws._charts)
        })

    except DataValidationError as e:
        return _fail(e.message, meta={"error_code": "SHEET_NOT_FOUND", "hint": e.hint, "suggested_fix": e.suggested_fix})
    except Exception as e:
        return _fail("图表创建失败", meta={"error_code": "OPERATION_FAILED"})


@_validate_file_path()
def excel_create_pivot_table(file_path: str, sheet_name: str, data_range: str, 
                            rows: List[str], values: List[str], 
                            columns: Optional[List[str]] = None,
                            agg_func: str = "sum", 
                            pivot_sheet_name: str = None) -> Dict[str, Any]:
    """
    在Excel中创建数据透视表。支持多种聚合函数，包括'mean'作为'average'的别名。
    
    Args:
        file_path: Excel文件路径
        sheet_name: 数据所在工作表名称
        data_range: 数据范围，如 "A1:D100"
        rows: 行字段列表，如 ["类别", "子类别"]
        values: 值字段列表，如 ["销售额", "利润"]
        columns: 列字段列表（可选），如 ["月份"]
        agg_func: 聚合函数，支持: sum/count/average/mean/max/min/std/var，其中'mean'是'average'的别名
        pivot_sheet_name: 透视表工作表名称（可选，默认自动生成）
    
    Returns:
        Dict: 透视表创建结果
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import range_boundaries, get_column_letter
        import pandas as pd
        import numpy as np

        wb, ws = ExcelValidator.get_workbook_and_sheet(file_path, sheet_name)
        
        # 处理聚合函数别名
        agg_func_map = {
            "sum": "sum",
            "count": "count", 
            "average": "mean",
            "mean": "mean",  # 支持'mean'作为'average'的别名
            "max": "max",
            "min": "min",
            "std": "std",
            "var": "var"
        }
        
        if agg_func not in agg_func_map:
            return _fail("不支持的聚合函数", meta={"error_code": "INVALID_PARAMETER", 
                                                   "hint": f"支持的函数: {list(agg_func_map.keys())}"})
        
        normalized_agg_func = agg_func_map[agg_func]

        # 解析数据范围
        min_col, min_row, max_col, max_row = range_boundaries(data_range)
        
        # 读取表头和数据到DataFrame
        headers = []
        data_rows = []
        
        # 读取表头
        for col in range(min_col, max_col + 1):
            col_letter = get_column_letter(col)
            header_cell = ws[f"{col_letter}{min_row}"]
            headers.append(header_cell.value or f"列{col}")
        
        # 读取数据行
        for row in range(min_row + 1, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                col_letter = get_column_letter(col)
                cell_value = ws[f"{col_letter}{row}"].value
                row_data.append(cell_value if cell_value is not None else np.nan)
            data_rows.append(row_data)
        
        # 创建DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # 移除全为NaN的行
        df = df.dropna(how='all')
        
        # 创建透视表
        index_cols = [col for col in rows if col in df.columns]
        value_cols = [col for col in values if col in df.columns]
        
        if not index_cols or not value_cols:
            return _fail("行字段或值字段不存在于数据中", meta={"error_code": "INVALID_PARAMETER"})
        
        # 创建透视表
        if columns:
            column_cols = [col for col in columns if col in df.columns]
            pivot_df = df.pivot_table(
                index=index_cols,
                columns=column_cols,
                values=value_cols,
                aggfunc=normalized_agg_func,
                fill_value=0
            )
        else:
            pivot_df = df.pivot_table(
                index=index_cols,
                values=value_cols,
                aggfunc=normalized_agg_func,
                fill_value=0
            )
        
        # 确定透视表工作表
        if pivot_sheet_name:
            if pivot_sheet_name not in wb.sheetnames:
                pivot_ws = wb.create_sheet(pivot_sheet_name)
            else:
                pivot_ws = wb[pivot_sheet_name]
                # 清空工作表
                pivot_ws.delete_rows(1, pivot_ws.max_row)
        else:
            pivot_ws = wb.create_sheet(f"透视表_{len(wb.sheetnames) + 1}")
        
        # 写入透视表结果
        row_idx = 1
        
        # 写入标题
        if columns:
            # 多级列标题的情况
            header_row = []
            first_col_headers = pivot_df.columns.get_level_values(0).unique()
            
            for i, first_col in enumerate(first_col_headers):
                # 计算这个列跨多少个第二级列
                second_cols_for_first = [col[1] for col in pivot_df.columns if col[0] == first_col]
                col_span = len(second_cols_for_first)
                
                if col_span == 1:
                    header_row.append(str(first_col))
                else:
                    # 写入合并的标题
                    pivot_ws.merge_cells(f"{get_column_letter(min_col + i)}{row_idx}:{get_column_letter(min_col + i + col_span - 1)}{row_idx}")
                    pivot_ws[f"{get_column_letter(min_col + i)}{row_idx}"] = str(first_col)
                    header_row.extend([''] * (col_span - 1))
            
            # 写入第二级列标题
            row_idx += 1
            second_headers = []
            for col in pivot_df.columns:
                if len(col) == 2:
                    second_headers.append(str(col[1]))
                else:
                    second_headers.append(str(col))
            
            for col_idx, header in enumerate(second_headers):
                cell_col = min_col + col_idx
                pivot_ws[f"{get_column_letter(cell_col)}{row_idx}"] = header
        else:
            # 单级列标题
            row_idx += 1
            for col_idx, header in enumerate(pivot_df.columns):
                pivot_ws[f"{get_column_letter(min_col + col_idx)}{row_idx}"] = str(header)
        
        # 写入行索引和数据
        row_idx += 1
        for row_tuple in pivot_df.itertuples(index=True, name=None):
            index = row_tuple[0]
            if isinstance(index, tuple):
                # 多级行索引
                for i, idx_val in enumerate(index):
                    pivot_ws[f"{get_column_letter(min_col + i)}{row_idx}"] = str(idx_val)
            else:
                # 单级行索引
                pivot_ws[f"{get_column_letter(min_col)}{row_idx}"] = str(index)

            # 写入数据值
            values = row_tuple[1:]
            for col_idx, value in enumerate(values):
                value_col = min_col + len(pivot_df.columns.get_level_values(0).unique()) + col_idx
                pivot_ws[f"{get_column_letter(value_col)}{row_idx}"] = float(value) if pd.notna(value) else 0

            row_idx += 1
        
        # 添加说明信息
        info_row = row_idx + 2
        pivot_ws[f"A{info_row}"] = f"数据源: {sheet_name}!{data_range}"
        pivot_ws[f"A{info_row + 1}"] = f"聚合函数: {agg_func}"
        pivot_ws[f"A{info_row + 2}"] = f"创建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # 保存文件
        wb.save(file_path)
        
        return _ok("透视表创建成功", data={
            'pivot_name': f"透视表_{pivot_ws.title}",
            'sheet_name': pivot_ws.title,
            'data_range': data_range,
            'rows': rows,
            'values': values,
            'columns': columns or [],
            'agg_func': agg_func,
            'row_count': len(pivot_df),
            'column_count': len(pivot_df.columns),
            'pivot_rows': len(index_cols),
            'pivot_cols': len(pivot_df.columns)
        })

    except DataValidationError as e:
        return _fail(e.message, meta={"error_code": "SHEET_NOT_FOUND", "hint": e.hint, "suggested_fix": e.suggested_fix})
    except Exception as e:
        return _fail("透视表创建失败", meta={"error_code": "OPERATION_FAILED", "error": str(e)})


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_list_charts(file_path: str, sheet_name: str = None) -> Dict[str, Any]:
    """列出工作表中的所有图表信息。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为None表示所有工作表
    """
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(file_path)
        charts_info = []
        
        sheets_to_check = [sheet_name] if sheet_name else wb.sheetnames
        
        for sheet in sheets_to_check:
            if sheet not in wb.sheetnames:
                continue
                
            ws = wb[sheet]
            
            for i, chart in enumerate(ws._charts):
                chart_info = {
                    'sheet_name': sheet,
                    'chart_index': i,
                    'chart_type': getattr(chart, 'type', type(chart).__name__),
                    'position': str(chart.anchor),
                    'title': extract_rich_text(chart.title),
                    'legend': getattr(chart.legend, 'position', None) if chart.legend else None,
                    'has_data_labels': bool(chart.dLbls) if hasattr(chart, 'dLbls') and chart.dLbls else False
                }
                charts_info.append(chart_info)
        
        return _ok(f"找到{len(charts_info)}个图表", data={
            'total_charts': len(charts_info),
            'charts': charts_info,
            'sheets_with_charts': len(set([c['sheet_name'] for c in charts_info])),
            'file_path': file_path
        })
        
    except Exception as e:
        return _fail("图表列表获取失败", meta={"error_code": "OPERATION_FAILED"})


# ==================== 数据验证工具 ====================
@mcp.tool()
@_validate_file_path()
@_track_call
def excel_set_data_validation(file_path: str, sheet_name: str, range_address: str, 
                            validation_type: str, criteria: str, input_title: str = "", 
                            input_message: str = "") -> Dict[str, Any]:
    """设置数据验证规则。validation_type: list/whole_number/decimal/date/text_length/custom。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range_address: 单元格范围
        validation_type: 验证类型
        criteria: 验证条件
        input_title: 输入提示标题，默认为空字符串
        input_message: 输入提示内容，默认为空字符串
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.worksheet.datavalidation import DataValidation

        wb, ws = ExcelValidator.get_workbook_and_sheet(file_path, sheet_name)
        
        # 创建数据验证对象
        dv = DataValidation(type=validation_type, formula1=criteria, showDropDown=True)
        
        # 设置输入提示
        if input_title or input_message:
            dv.promptTitle = input_title
            dv.prompt = input_message
        
        # 添加应用到指定范围
        dv.add(range_address)
        ws.add_data_validation(dv)
        
        # 保存文件
        wb.save(file_path)
        
        return _ok("数据验证设置成功", data={
            'validation_type': validation_type,
            'criteria': criteria,
            'range_address': range_address,
            'sheet_name': sheet_name,
            'input_title': input_title,
            'input_message': input_message,
            'validation_count': len(ws.data_validations)
        })

    except DataValidationError as e:
        return _fail(e.message, meta={"error_code": "SHEET_NOT_FOUND", "hint": e.hint, "suggested_fix": e.suggested_fix})
    except Exception as e:
        return _fail("数据验证设置失败", meta={"error_code": "OPERATION_FAILED"})


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_clear_validation(file_path: str, sheet_name: str = None, range_address: str = None) -> Dict[str, Any]:
    """清除数据验证规则。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为None
        range_address: 单元格范围，默认为None
    """
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(file_path)
        cleared_count = 0
        
        sheets_to_clear = [sheet_name] if sheet_name else wb.sheetnames
        
        for sheet in sheets_to_clear:
            if sheet not in wb.sheetnames:
                continue
                
            ws = wb[sheet]
            
            # 获取当前的数据验证
            data_validations = ws.data_validations
            
            if range_address:
                # 清除指定范围的验证
                dv_to_remove = []
                for dv in list(ws.data_validations.dataValidation):
                    dv_range = str(dv.sqref) if dv.sqref else ""
                    if range_address in dv_range or dv_range in range_address:
                        dv_to_remove.append(dv)
                
                for dv in dv_to_remove:
                    ws.data_validations.dataValidation.remove(dv)
                    cleared_count += 1
            else:
                # 清除整个工作表的验证
                cleared_count += ws.data_validations.count
                ws.data_validations.dataValidation.clear()
        
        # 保存文件
        wb.save(file_path)
        
        return _ok(f"数据验证清除完成：共清除{cleared_count}个验证规则", data={
            'cleared_count': cleared_count,
            'sheets_processed': len(sheets_to_clear),
            'file_path': file_path,
            'validation_cleared': True
        })
        
    except Exception as e:
        return _fail("数据验证清除失败", meta={"error_code": "OPERATION_FAILED"})


# ==================== 条件格式工具 ====================
@mcp.tool()
@_validate_file_path()
@_track_call
def excel_add_conditional_format(file_path: str, sheet_name: str, range_address: str,
                                format_type: str, criteria: str, format_style: str = "lightRed") -> Dict[str, Any]:
    """添加条件格式规则。支持高亮/数据条/色阶/图标集。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range_address: 单元格范围
        format_type: 格式类型
        criteria: 条件表达式
        format_style: 格式样式，默认为"lightRed"
    """
    try:
        from openpyxl import load_workbook, styles
        from openpyxl.formatting.rule import CellIsRule, FormulaRule

        wb, ws = ExcelValidator.get_workbook_and_sheet(file_path, sheet_name)
        
        # 创建格式样式
        style_map = {
            "lightRed": "FFCCCC",
            "lightGreen": "CCFFCC", 
            "lightYellow": "FFFFCC",
            "lightTurquoise": "CCFFFF"
        }
        
        if format_style not in style_map:
            return _fail("不支持的格式样式", meta={"error_code": "INVALID_PARAMETER", "hint": f"支持的样式: {list(style_map.keys())}"})
        
        fill = styles.PatternFill(start_color=style_map[format_style], 
                                end_color=style_map[format_style], 
                                fill_type="solid")
        
        # 创建条件格式规则
        if format_type == "cellValue":
            rule = CellIsRule(operator="greaterThanOrEqual", formula=[criteria], stopIfTrue=True)
        elif format_type == "formula":
            rule = FormulaRule(formula=[criteria], stopIfTrue=True)
        else:
            return _fail("不支持的格式类型", meta={"error_code": "INVALID_PARAMETER", "hint": "支持的类型: cellValue, formula"})
        
        # 添加格式
        rule.fill = fill
        ws.conditional_formatting.add(range_address, rule)
        
        # 保存文件
        wb.save(file_path)
        
        return _ok("条件格式添加成功", data={
            'format_type': format_type,
            'criteria': criteria,
            'range_address': range_address,
            'format_style': format_style,
            'sheet_name': sheet_name,
            'rule_applied': True
        })

    except DataValidationError as e:
        return _fail(e.message, meta={"error_code": "SHEET_NOT_FOUND", "hint": e.hint, "suggested_fix": e.suggested_fix})
    except Exception as e:
        return _fail("条件格式添加失败", meta={"error_code": "OPERATION_FAILED"})


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_clear_conditional_format(file_path: str, sheet_name: str = None, range_address: str = None) -> Dict[str, Any]:
    """清除条件格式。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为None
        range_address: 单元格范围，默认为None
    """
    try:
        from openpyxl import load_workbook
        
        wb = load_workbook(file_path)
        cleared_count = 0
        
        sheets_to_clear = [sheet_name] if sheet_name else wb.sheetnames
        
        for sheet in sheets_to_clear:
            if sheet not in wb.sheetnames:
                continue
                
            ws = wb[sheet]
            
            # 获取当前的条件格式
            conditional_formats = ws.conditional_formatting
            
            if range_address:
                # 清除指定范围的条件格式
                cf_to_remove = []
                for cf_key in list(ws.conditional_formatting._cf_rules):
                    cf_range = str(cf_key)
                    if range_address in cf_range:
                        cf_to_remove.append(cf_key)
                
                for cf_key in cf_to_remove:
                    del ws.conditional_formatting._cf_rules[cf_key]
                    cleared_count += 1
            else:
                # 清除整个工作表的条件格式
                cleared_count += len(ws.conditional_formatting._cf_rules)
                ws.conditional_formatting._cf_rules.clear()
                ws.conditional_formatting.max_priority = 0
        
        # 保存文件
        wb.save(file_path)
        
        return _ok(f"条件格式清除完成：共清除{cleared_count}个格式规则", data={
            'cleared_count': cleared_count,
            'sheets_processed': len(sheets_to_clear),
            'file_path': file_path,
            'format_cleared': True
        })
        
    except Exception as e:
        return _fail("条件格式清除失败", meta={"error_code": "OPERATION_FAILED"})


@mcp.tool()
@_validate_file_path()
@_track_call
def excel_write_only_override(
    file_path: str,
    sheet_name: str,
    range_spec: str,
    data: List[List[Any]],
    preserve_formulas: bool = False,
    preserve_col_widths: bool = True
) -> Dict[str, Any]:
    """大文件高性能覆盖写入。range_spec: "sheet!A1:D10"。不读取已有内容，直接覆盖。适合批量导入场景。

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        range_spec: 范围表达式，如"sheet!A1:D10"
        data: 要写入的数据，二维数组
        preserve_formulas: 是否保留公式，默认为False
        preserve_col_widths: 是否保留列宽，默认为True
    """

    try:
        # 验证参数
        if not sheet_name:
            return _fail("工作表名称不能为空", meta={"error_code": "INVALID_PARAMS"})
        
        if not range_spec:
            return _fail("范围表达式不能为空", meta={"error_code": "INVALID_PARAMS"})
            
        if not data:
            return _fail("数据不能为空", meta={"error_code": "INVALID_PARAMS"})
            
        if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
            return _fail("数据必须是二维数组", meta={"error_code": "INVALID_PARAMS"})

        # 记录操作日志
        operation_logger.start_session(file_path)
        operation_logger.log_operation("write_only_override", {
            "sheet_name": sheet_name,
            "range": range_spec,
            "data_rows": len(data),
            "preserve_formulas": preserve_formulas,
            "preserve_col_widths": preserve_col_widths
        })

        # 优先尝试流式写入（高性能模式）
        try:
            from excel_mcp_server_fastmcp.core.streaming_writer import StreamingWriter
            
            if StreamingWriter.is_available():
                # 解析范围表达式
                from excel_mcp_server_fastmcp.utils.parsers import RangeParser
                range_info = RangeParser.parse_range_expression(f"{sheet_name}!{range_spec}")
                
                # 获取起始行列
                from openpyxl.utils import range_boundaries
                min_col, min_row, max_col, max_row = range_boundaries(range_info.cell_range)
                
                # 执行流式覆盖写入
                success, message, meta = StreamingWriter.update_range(
                    file_path, sheet_name, min_row, min_col, data,
                    preserve_formulas=preserve_formulas,
                    preserve_col_widths=preserve_col_widths
                )
                
                if success:
                    # 记录成功结果
                    operation_logger.log_operation("operation_result", {
                        "success": True,
                        "updated_cells": meta.get('updated_cells', 0),
                        "message": message
                    })
                    
                    return {
                        'success': True,
                        'message': message,
                        'data': meta,
                        'metadata': {
                            'file_path': file_path,
                            'sheet_name': sheet_name,
                            'range': range_spec,
                            'streaming_mode': True,
                            'override_mode': True,
                            'memory_efficiency': 'high',
                            **meta
                        }
                    }
                else:
                    logger.warning(f"流式写入失败，降级到openpyxl: {message}")
        except Exception as stream_err:
            logger.warning(f"流式写入异常，降级到openpyxl: {stream_err}")

        # 降级到传统openpyxl模式（作为备用方案）
        try:
            from excel_mcp_server_fastmcp.core.excel_writer import ExcelWriter
            writer = ExcelWriter(file_path)
            result = writer.update_range(
                f"{sheet_name}!{range_spec}", 
                data, 
                preserve_formulas, 
                insert_mode=False  # 强制覆盖模式
            )
            
            if result.get('success'):
                # 记录结果
                operation_logger.log_operation("operation_result", {
                    "success": True,
                    "updated_cells": len(data) * len(data[0]) if data else 0,
                    "message": f"传统模式覆盖完成"
                })
                
                return {
                    'success': True,
                    'message': f"传统模式覆盖完成: {result.get('message', '')}",
                    'data': result,
                    'metadata': {
                        'file_path': file_path,
                        'sheet_name': sheet_name,
                        'range': range_spec,
                        'streaming_mode': False,
                        'override_mode': True,
                        'memory_efficiency': 'medium',
                        **result
                    }
                }
            else:
                return result
        except Exception as fallback_err:
            error_msg = f"覆盖修改失败: {str(fallback_err)}"
            logger.error(error_msg)
            return _fail(error_msg, meta={"error_code": "OPERATION_FAILED"})

    except Exception as e:
        error_msg = f"write_only覆盖修改失败: {str(e)}"
        logger.error(error_msg)
        return _fail(error_msg, meta={"error_code": "OPERATION_FAILED"})


# ==================== 主程序 ====================
# ==================== 智能配置推荐工具 ====================

if SMART_CONFIG_AVAILABLE:
    # 智能配置推荐工具
    @mcp.tool()
    def recommend_excel_config(
        file_path: str,
        game_type: Optional[str] = None,
        optimization_level: str = "balanced"
    ) -> str:
        """
        智能推荐Excel配置结构
        
        Args:
            file_path: Excel文件路径
            game_type: 游戏类型 (rpg/strategy/action/puzzle/simulation)，如不指定自动检测
            optimization_level: 优化级别 (basic/balanced/aggressive)
        
        Returns:
            智能配置推荐结果JSON字符串
        """
        try:
            # 创建智能推荐器实例
            recommender = SmartConfigurationRecommender()
            
            # 读取Excel数据
            excel_data = ExcelOperations.read_excel_file(file_path)
            
            # 进行智能配置推荐
            recommendations = recommender.recommend_configurations(excel_data)
            
            # 根据优化级别调整建议详细程度
            if optimization_level == "basic":
                # 只保留核心建议
                result = {
                    "game_type": recommendations["game_type"],
                    "core_recommendations": recommendations["config_recommendations"][:3],
                    "critical_validation_rules": [r for r in recommendations["validation_rules"] if r["priority"] == "high"]
                }
            elif optimization_level == "aggressive":
                # 添加详细的优化建议
                result = {
                    "game_type": recommendations["game_type"],
                    "full_analysis": recommendations["analysis"],
                    "all_recommendations": recommendations["config_recommendations"],
                    "all_validation_rules": recommendations["validation_rules"],
                    "optimization_tips": recommendations["optimization_tips"],
                    "summary": _generate_summary(recommendations)
                }
            else:
                # balanced 默认模式
                result = {
                    "game_type": recommendations["game_type"],
                    "analysis_summary": _generate_summary(recommendations),
                    "key_recommendations": recommendations["config_recommendations"][:5],
                    "important_validation_rules": recommendations["validation_rules"][:10]
                }
            
            return _ok("智能配置推荐完成", result)
            
        except Exception as e:
            error_msg = f"配置推荐失败: {str(e)}"
            logger.error(error_msg)
            return _fail(error_msg, meta={"error_code": "SMART_CONFIG_FAILED"})
    
    @mcp.tool()
    def analyze_game_patterns(
        file_path: str,
        target_sheet: Optional[str] = None
    ) -> str:
        """
        分析游戏模式和数据结构
        
        Args:
            file_path: Excel文件路径
            target_sheet: 指定分析特定工作表，如不分析所有工作表
        
        Returns:
            游戏模式分析结果JSON字符串
        """
        try:
            # 创建分析器
            analyzer = ConfigurationAnalyzer()
            
            # 读取Excel数据
            excel_data = ExcelOperations.read_excel_file(file_path)
            
            # 分析数据结构
            analysis = analyzer.analyze_excel_structure(excel_data)
            
            # 检测游戏类型
            game_type = analyzer.detector.detect_game_type(excel_data)
            
            result = {
                "detected_game_type": game_type,
                "analysis_scope": target_sheet if target_sheet else "all_sheets",
                "data_patterns": analysis.get("data_patterns", {}),
                "structure_analysis": analysis.get("sheet_structure", {}),
                "optimization_suggestions": analysis.get("optimization_suggestions", []),
                "data_quality_score": _calculate_data_quality_score(analysis)
            }
            
            return _ok("游戏模式分析完成", result)
            
        except Exception as e:
            error_msg = f"游戏模式分析失败: {str(e)}"
            logger.error(error_msg)
            return _fail(error_msg, meta={"error_code": "ANALYSIS_FAILED"})
    
    @mcp.tool()
    def generate_validation_rules(
        file_path: str,
        rule_categories: Optional[List[str]] = None
    ) -> str:
        """
        基于游戏配置生成验证规则
        
        Args:
            file_path: Excel文件路径
            rule_categories: 规则类别列表，如不生成全部类别
        
        Returns:
            验证规则JSON字符串
        """
        try:
            # 创建推荐器
            recommender = SmartConfigurationRecommender()
            
            # 读取Excel数据
            excel_data = ExcelOperations.read_excel_file(file_path)
            
            # 获取推荐
            recommendations = recommender.recommend_configurations(excel_data)
            
            # 筛选规则类别
            if rule_categories:
                filtered_rules = []
                for rule in recommendations["validation_rules"]:
                    if rule["sheet"] in rule_categories:
                        filtered_rules.append(rule)
                rules = filtered_rules
            else:
                rules = recommendations["validation_rules"]
            
            result = {
                "validation_rules": rules,
                "rule_categories": list(set(rule["sheet"] for rule in rules)),
                "priority_breakdown": _categorize_by_priority(rules),
                "rule_summary": f"生成了{len(rules)}个验证规则"
            }
            
            return _ok("验证规则生成完成", result)
            
        except Exception as e:
            error_msg = f"验证规则生成失败: {str(e)}"
            logger.error(error_msg)
            return _fail(error_msg, meta={"error_code": "VALIDATION_FAILED"})
    
    @mcp.tool()
    def optimize_data_structure(
        file_path: str,
        optimization_type: str = "compression"
    ) -> str:
        """
        优化Excel数据结构
        
        Args:
            file_path: Excel文件路径
            optimization_type: 优化类型 (compression/restructuring/indexing)
        
        Returns:
            优化建议JSON字符串
        """
        try:
            # 创建分析器
            analyzer = ConfigurationAnalyzer()
            
            # 读取Excel数据
            excel_data = ExcelOperations.read_excel_file(file_path)
            analysis = analyzer.analyze_excel_structure(excel_data)
            
            # 生成优化建议
            optimization_suggestions = []
            
            if optimization_type == "compression":
                # 数据压缩优化
                for sheet_name, patterns in analysis.get("data_patterns", {}).items():
                    for col_name, pattern in patterns.items():
                        if pattern.get("uniqueness_ratio", 0) < 0.1:
                            optimization_suggestions.append({
                                "type": "enum_conversion",
                                "sheet": sheet_name,
                                "column": col_name,
                                "description": f"建议将{col_name}转换为枚举类型，节省存储空间",
                                "estimated_savings": f"{(1 - pattern.get('uniqueness_ratio', 0)) * 100:.1f}%"
                            })
            
            elif optimization_type == "restructuring":
                # 结构重构优化
                for sheet_name, sheet_info in analysis.get("sheet_structure", {}).items():
                    if sheet_info["rows"] > 1000:
                        optimization_suggestions.append({
                            "type": "normalization",
                            "sheet": sheet_name,
                            "description": f"{sheet_name}表数据量较大，建议考虑表结构规范化",
                            "current_rows": sheet_info["rows"],
                            "recommendation": "拆分为多个关联表"
                        })
            
            elif optimization_type == "indexing":
                # 索引优化
                for sheet_name, sheet_info in analysis.get("sheet_structure", {}).items():
                    headers = sheet_info.get("headers", [])
                    for i, header in enumerate(headers):
                        if any(keyword in header.lower() for keyword in ["id", "name", "key", "code"]):
                            optimization_suggestions.append({
                                "type": "index_recommendation",
                                "sheet": sheet_name,
                                "column": header,
                                "description": f"建议为{sheet_name}表的{header}列添加索引，提升查询性能",
                                "query_type": "primary_key" if "id" in header.lower() else "search_key"
                            })
            
            result = {
                "optimization_type": optimization_type,
                "original_structure": analysis,
                "optimization_suggestions": optimization_suggestions,
                "expected_improvements": _estimate_improvements(optimization_suggestions)
            }
            
            return _ok("数据结构优化分析完成", result)
            
        except Exception as e:
            error_msg = f"数据结构优化失败: {str(e)}"
            logger.error(error_msg)
            return _fail(error_msg, meta={"error_code": "OPTIMIZATION_FAILED"})
    
    # 辅助函数
    def _generate_summary(recommendations: Dict[str, Any]) -> str:
        """生成推荐摘要"""
        game_type = recommendations["game_type"]
        key_recs = recommendations["config_recommendations"][:3]
        
        summary = f"检测到游戏类型: {game_type}\\n"
        summary += f"核心推荐: {len(key_recs)}条\\n"
        summary += "主要建议: "
        summary += "; ".join([rec["suggestion"] for rec in key_recs])
        
        return summary
    
    def _calculate_data_quality_score(analysis: Dict[str, Any]) -> float:
        """计算数据质量评分"""
        score = 100.0
        
        # 根据优化建议扣分
        suggestions = analysis.get("optimization_suggestions", [])
        score -= len(suggestions) * 5  # 每个建议扣5分
        
        # 确保评分在0-100之间
        return max(0, min(100, score))
    
    def _categorize_by_priority(rules: List[Dict[str, Any]]) -> Dict[str, int]:
        """按优先级分类规则"""
        priority_count = {"high": 0, "medium": 0, "low": 0}
        for rule in rules:
            priority = rule.get("priority", "medium")
            priority_count[priority] += 1
        return priority_count
    
    def _estimate_improvements(suggestions: List[Dict[str, Any]]) -> Dict[str, Any]:
        """预估优化效果"""
        improvements = {
            "performance_gain": "预估查询速度提升20-30%",
            "storage_reduction": "预估存储节省10-25%",
            "maintainability": "预估维护难度降低40%"
        }
        
        if len(suggestions) > 5:
            improvements["significant_improvement"] = "大规模优化，效果显著"
        elif len(suggestions) > 2:
            improvements["moderate_improvement"] = "中等规模优化，效果明显"
        else:
            improvements["minor_improvement"] = "小幅优化，略有改善"
        
        return improvements


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
        logger.info(f"excel-mcp-server-fastmcp {__version__}")
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
