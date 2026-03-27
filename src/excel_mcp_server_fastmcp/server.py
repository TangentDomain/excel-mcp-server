#!/usr/bin/env python3
"""
Excel MCP Server - 基于 FastMCP 和 openpyxl 实现

📋 **统一返回格式说明**
==============================

所有工具都采用统一的JSON返回格式，确保AI客户端可以可靠解析：

✅ **成功返回格式**:
{
  "success": true,
  "message": "操作成功说明",
  "data": {
    // 实际数据，根据工具不同而异
  },
  "meta": {
    "error_code": "SUCCESS",
    "execution_time_ms": 150,
    "cache_hit": true,
    "file_size_mb": 2.5
  }
}

❌ **失败返回格式**:
{
  "success": false,
  "message": "详细的错误信息💡 建议修复方法",
  "data": null,
  "meta": {
    "error_code": "SHEET_NOT_FOUND",
    "suggested_fix": "请使用excel_list_sheets确认工作表名称",
    "file_path": "配置.xlsx",
    "sheet_name": "技能表"
  }
}

🔍 **字段说明**:
- success: 操作是否成功
- message: 结果描述（失败时包含💡修复提示）
- data: 实际数据内容（成功时）
- meta: 元信息（错误码、执行时间、缓存状态等）

🎮 **游戏开发优化**:
- 中文友好：支持中文字段名和描述
- 智能错误：错误信息直接提供修复建议
- 缓存优化：重复操作自动提速
- 内存控制：大文件使用流式写入

---

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
    instructions=r"""🎮 游戏开发Excel配置表管理专家 — 44个工具

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
- 默认覆盖: update_range默认覆盖，保留数据用insert_mode=True

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
    """统一成功响应: {success, message, data, meta}"""
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
    """统一错误响应: {success, message, meta}。自动附加error_code对应的修复提示。"""
    r: dict = {"success": False, "message": message}
    if meta:
        r["meta"] = meta
    # 自动附加集中式错误提示（仅当message中还没有💡提示时）
    error_code = (meta or {}).get('error_code', '')
    hint = _ERROR_HINTS.get(error_code, '')
    if hint and '💡' not in message:
        r["message"] = message + f'\n💡 {hint}'
    return r


def _wrap(result: dict, meta: dict = None) -> dict:
    """包装Operations层返回，metadata→meta，添加上下文meta，统一success字段"""
    if not isinstance(result, dict):
        return result
    # 统一error→message，确保AI只需检查message键
    err_val = result.get('error')
    if isinstance(err_val, str) and not result.get('message'):
        result['message'] = result.pop('error')
    if "success" not in result:
        result["success"] = True
    if "metadata" in result:
        m = result.pop("metadata")
        if isinstance(m, dict) and m:
            merged = {**m, **(meta or {})}
            result["meta"] = merged
            meta = None  # 已合并，不再重复设置
    if meta and "meta" not in result:
        result["meta"] = meta

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
@_track_call
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    """
📋 Excel工作表清单 - 获取文件中的所有工作表名称

**核心功能**: 快速列出Excel文件中所有工作表的名称，支持多文件表结构了解。在进行SQL查询前，建议先用此工具确认目标工作表存在。

**🎮 游戏开发场景**:
• **表结构扫描**: 快速了解技能表、装备表、怪物表等存在情况
• **多表操作**: 确认目标工作表存在后再进行查询或修改操作
• **文件检查**: 验证Excel文件是否包含预期的工作表
• **配置管理**: 了解配置文件包含哪些类型的配置表

**📊 返回信息**:
• **sheets**: 工作表名称列表（数组格式）
• **success**: 操作是否成功
• **message**: 状态消息或错误信息

**🔧 参数说明**:
• **file_path**: Excel文件路径（支持相对路径）

**⚡ 使用建议**:
• **必用前置**: 在执行excel_query前先用此工具确认工作表存在
• **批量操作**: 多个文件操作前先用此工具了解表结构
• **错误预防**: 避免"工作表不存在"的错误

**🚫 注意事项**:
• 只返回工作表名称，不包含表结构信息
• 了解表结构需要配合excel_describe_table
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.list_sheets(file_path))


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
🔍 智能文本搜索 - 快速定位Excel中的目标内容

**核心功能**: 在Excel文件中搜索文本、数值或公式，支持正则表达式。精确匹配单元格位置，支持灵活的搜索策略。

**🎮 游戏开发场景**:
• **技能查找**: "火" 搜索所有火系技能
• **装备定位**: "传说" 查找所有传说装备
• **问题排查**: "测试" 查找所有测试数据便于清理
• **关键字定位**: "CD" 或 "冷却" 找到冷却时间字段
• **数值搜索**: "100+" 查找大于100的数值配置

**🔧 搜索模式**:
• **文本搜索**: 直接搜索"火球术"、"治疗"等关键词
• **正则表达式**: "等级[1-5]" 匹配"等级1"到"等级5"
• **大小写敏感**: 区分"mage"和"Mage"，适合精确匹配
• **全词匹配**: "火"只匹配完整单词"火"，不匹配"火星"
• **公式搜索**: 查找包含特定公式的单元格
• **范围限制**: 在指定区域"A1:C100"内搜索

**📊 搜索结果**:
• **位置信息**: 精确的行列位置(row/column)
• **单元格值**: 匹配的实际内容
• **工作表**: 来源工作表名称
• **文件路径**: 文件完整路径（多文件搜索时）

**💡 使用策略**:
• **模糊查找**: 用"火"搜所有火系技能
• **精确匹配**: 用"火球术"搜特定技能
• **模式匹配**: 用"[0-9]+"搜所有数值
• **批量查找**: 先搜索定位，再批量修改
• **数据清理**: 用"测试"、"调试"等词查找需要清理的数据

**🔗 配合使用**:
• 定位后修改: 搜索结果→excel_update_query批量修改
• 查看详情: excel_describe_table了解字段含义
• 数据验证: 搜索后检查数据完整性
• 版本对比: 搜索相同关键字对比版本差异

**⚡ 性能提示**:
• 大文件建议指定range缩小搜索范围
• 正则表达式在大文件中可能较慢
• 重复搜索会自动缓存，第二次更快
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
🔎 目录批量搜索 - 在文件夹下所有Excel文件中搜索内容

**核心功能**: 递归扫描目录下所有Excel文件，返回匹配的文件名+单元格位置+值。支持正则、大小写、全词匹配、文件名过滤。
搜索单个文件请用excel_search。

**🎮 游戏开发场景**:
• **全局配置搜索**: 在整个配置目录中查找某个数值（如"攻击力:500"出现在哪些文件）
• **文本替换前置**: 找到需要修改的文件和位置，再逐个修改
• **跨表关联检查**: 查找引用了某个ID的所有配置表
• **废弃资源排查**: 搜索已下线活动ID，确认是否还有残留引用

**🔧 参数说明**:
• **pattern**: 搜索文本（支持正则表达式）
• **use_regex**: True启用正则匹配
• **file_pattern**: 文件名过滤（如"*.xlsx"只搜xlsx文件）
• **max_files**: 最大搜索文件数（防止卡死）

**⚡ 使用建议**:
• 先用file_pattern缩小范围（如只搜"技能*.xlsx"），再搜内容
• 大目录搜索较慢，建议缩小范围或设置max_files
    """
    _path_err = _validate_path(directory_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.search_directory(directory_path, pattern, case_sensitive, whole_word, use_regex, include_values, include_formulas, recursive, file_extensions, file_pattern, max_files))


@mcp.tool()
@_track_call
def excel_get_range(
    file_path: str,
    range: str,
    include_formatting: bool = False
) -> Dict[str, Any]:
    """
📥 极速数据读取 - 精确获取Excel数据

**核心功能**: 使用calamine引擎极速读取指定范围的原始数据。适合获取精确的单元格数据，速度快于传统方式100倍。

**⚡ 性能优势**:
• **超高速读取**: calamine引擎比传统方式快2300倍，2000行仅需230ms
• **智能缓存**: 同一文件重复查询自动提速30-100倍
• **内存友好**: 读取仅针对指定范围，不加载整个文件
• **格式可选**: 只需数据→set include_formatting=False，需样式→设为True

**🎮 游戏开发场景**:
• **配置读取**: 读取技能列表前10行进行预览
• **数据提取**: 装备表中特定范围的属性值
• **批量导出**: 技能表中指定范围的技能数据
• **测试验证**: 获取特定单元格的数据进行验证

**📋 使用场景**:
• ✅ **已知范围精确读取**: 如"A1:C10"、"技能表!B2:D50"  
• ✅ **批量数据导出**: 获取连续的数值、文本、公式数据
• ✅ **快速数据验证**: 修改前先读取确认数据正确性
• ✅ **精确控制**: 需要精确指定单元格位置的操作

**❌ 不适用场景**:
• ❌ 数据查询分析→用`excel_query`进行筛选和聚合
• ❌ 条件修改→用`excel_update_query`进行批量修改
• ❌ 全表统计→用`excel_describe_table`了解结构

**🔧 参数说明**:
• **range**: 范围表达式，支持"A1:B10"、"Sheet1!A1:C10"、"技能表!B:D"
• **include_formatting**: 是否包含字体、颜色、边框等格式信息

**💡 使用建议**:
• 小范围数据用此工具，复杂分析用excel_query
• 同一文件多次读取自动缓存，第二次更快
• 修改数据后再次读取可确认修改结果
• 双行表头会自动处理，列名映射在结果中
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

        return _fail(f"范围表达式验证失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})

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

    return _wrap(result)


@mcp.tool()
@_track_call
def excel_get_headers(
    file_path: str,
    sheet_name: Optional[str] = None,
    header_row: int = 1,
    max_columns: Optional[int] = None
) -> Dict[str, Any]:
    """
📋 智能表头提取器 - 数据结构的导航仪

**核心功能**: 快速提取Excel工作表的表头信息，特别针对游戏开发的双行表头优化。提供清晰的字段名和描述映射。

**🎮 游戏开发场景**:
• **技能表分析**: 提取技能ID、名称、类型、伤害、冷却等字段信息
• **装备表检查**: 获取装备ID、品质、属性、套装、获取方式等字段名
• **怪物数据审查**: 收集怪物ID、名称、等级、血量、攻击、防御等字段
• **配置表确认**: 验证技能配置、装备配置、掉落配置的表头结构

**🔍 核心功能**:
• **双行表头支持**: 自动识别第1行中文描述+第2行英文字段名
• **批量表头获取**: 不传sheet_name获取所有工作表的表头
• **自定义行号**: header_row参数指定表头起始位置（1-based）
• **列数限制**: max_columns限制返回的列数，避免信息过载
• **格式保持**: 保持原始表头格式和结构

**📊 返回信息**:
• **字段名**: 英文字段名（如skill_name, damage）
• **字段描述**: 中文描述（如"技能名称"、"伤害值"）
• **工作表名**: 来源工作表的完整名称
• **字段顺序**: 按Excel中的实际顺序排列
• **字段数量**: 总共有多少个字段列

**💡 使用策略**:
1️⃣ **初见新表**: `excel_get_headers("配置.xlsx")` 快速了解所有表结构
2️⃣ **聚焦分析**: `excel_get_headers("配置.xlsx", "技能表")` 查看特定表的字段
3️⃣ **自定义位置**: `header_row=2` 从第2行开始读取表头
4️⃣ **精简查看**: `max_columns=10` 只显示前10个字段

**🔗 配合使用**:
• 结构了解: 结合excel_describe_table获取完整类型和样本信息
• 数据查询: 基于字段名进行excel_query数据检索
• 批量修改: 使用字段名进行excel_update_query条件修改
• 版本对比: excel_compare_sheets对比不同版本的表头变化

**⚡ 性能提示**:
• 比excel_describe_table更快，专门针对表头信息
• 大文件建议指定max_columns减少数据量
• 重复读取相同文件有缓存，第二次更快

**🎯 选择指南**:
• • 看表头结构→用excel_get_headers
• • 看完整结构(类型+样本)→用excel_describe_table  
• • 看数据详情→用excel_get_range
• • 看分析统计→用excel_query
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    if sheet_name is None:
        return _wrap(ExcelOperations.get_all_headers(file_path, header_row, max_columns))
    return _wrap(ExcelOperations.get_headers(file_path, sheet_name, header_row, max_columns))


@mcp.tool()
@_track_call
def excel_update_range(
    file_path: str,
    range: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = True,
    streaming: bool = True
) -> Dict[str, Any]:
    """
✏️ 精确范围写入 - 批量数据的快速写入器

**核心功能**: 将二维数组数据写入Excel的指定范围，支持覆盖、插入和流式模式。适合精确位置的批量数据写入。

**🎮 游戏开发场景**:
• **配置批量导入**: 一次性写入数十条技能/装备数据到配置表
• **数值批量调整**: 按列批量修改伤害、冷却、消耗等数值
• **新表初始化**: 创建新的配置表并填充初始数据
• **数据迁移**: 从一个表提取数据写入另一个表

**🔧 写入模式**:
• **覆盖模式**: 直接覆盖目标范围内的所有数据（streaming=True时使用流式覆盖）
• **插入模式**: 在目标位置插入新数据，原数据下移（streaming=True时使用流式插入）
• **流式写入**: streaming=True（默认）使用calamine引擎，内存占用低，适合大文件
• **传统写入**: streaming=False使用openpyxl，保留所有格式但内存占用高

**📋 参数详解**:
• **data**: 二维数组，外层为行，内层为列值
• **range**: 目标范围（如"Sheet1!A1:C10"、"A:D"）
• **preserve_formulas**: True时保留已有公式不被覆盖
• **insert_mode**: True时插入新行而非覆盖
• **streaming**: True时使用流式引擎，性能更好

**💡 使用建议**:
• 条件修改→用`excel_update_query`（SQL语法更直观）
• 按ID合并→用`excel_upsert_row`（智能upsert）
• 追加到末尾→用`excel_batch_insert_rows`（自动定位末尾）
• 精确位置写入→用本工具（最灵活）

**⚡ 性能提示**:
• streaming=True时，覆盖模式和插入模式都使用流式写入，内存占用低
• preserve_formulas=False时性能最佳，插入模式下会自动使用流式插入
• 大文件强烈建议使用streaming=True，内存占用与文件大小无关
• 大文件建议使用streaming减少内存占用
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
    """
🛡️ 数据影响评估器 - 修改前的安全网

**核心功能**: 在修改或删除数据前评估潜在影响，不实际执行任何操作。帮助了解操作会影响到多少数据、什么类型的数据、以及潜在风险。

**🎮 游戏开发场景**:
• **数值修改前评估**: 修改技能伤害前了解会影响多少技能
• **批量删除前检查**: 删除测试数据前确认不会误删正式数据
• **配置表修改预检**: 修改装备配置前评估影响范围
• **数据清理验证**: 清理无效数据前确认安全

**📊 评估模式**:
• **快速预览** (detailed=False): 行数+列数+风险等级+当前数据快照
• **全面评估** (detailed=True): 数据类型分析+结果预测+安全建议+风险评级

**🔍 评估内容**:
• **影响范围**: 受影响的行数、列数、单元格总数
• **数据类型**: 当前数据的类型分布（数字/文本/日期/空值）
• **风险等级**: LOW/MEDIUM/HIGH三级风险评级
• **安全建议**: 基于数据内容提供操作建议
• **结果预测**: 预测操作后的数据变化

**💡 使用建议**:
1️⃣ 修改前先评估: `excel_assess_data_impact("技能表.xlsx", "A1:C100")` 
2️⃣ 删除前必评估: 避免误删重要数据
3️⃣ 结合备份: 重要操作前先用excel_create_backup创建备份
4️⃣ 大操作分段: 影响范围过大时考虑分批操作
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err

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
    """
📋 操作历史 - 查看最近的Excel操作记录

**核心功能**: 获取最近的操作历史，包含操作类型、文件路径、成功/失败状态。可按文件过滤。

**🎮 游戏开发场景**:
• **操作审计**: 查看某个配置表最近被修改了什么
• **错误排查**: 出问题时回顾最近的操作记录
• **团队协作**: 确认外包修改了哪些配置
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
    """
🛡️ 智能备份系统 - 数据安全的守护者

**核心功能**: 创建Excel文件的时间戳备份，保护重要配置数据。自动管理备份目录，提供完整的备份信息。

**🎮 游戏开发场景**:
• **版本控制**: 修改技能表前自动备份，方便版本回退
• **安全测试**: 调整数值配置前备份，避免误操作导致数据丢失
• **团队协作**: 重要配置修改前备份，方便团队协作和问题追踪
• **数据迁移**: 游戏版本更新前完整备份，确保可回滚到稳定版本

**🔧 备份特性**:
• **智能路径**: 默认存储在文件同目录的`.excel_mcp_backups/`文件夹
• **时间戳命名**: 自动添加`YYYYMMDD_HHMMSS`时间戳，避免覆盖
• **完整信息**: 记录原始大小、备份大小、备份路径、时间戳
• **自动目录**: 备份目录不存在时自动创建
• **文件完整性**: 使用shutil.copy2保持文件元数据

**📊 备份信息**:
• **备份文件**: 完整备份路径
• **备份目录**: 备份存储的根目录
• **文件对比**: 原始文件大小 vs 备份文件大小
• **时间戳**: 精确的备份时间
• **文件名**: 带时间戳的备份文件名

**💡 使用建议**:
1️⃣ **重要操作前**: 修改数值配置前必备份
2️⃣ **版本管理**: 大版本更新前创建完整备份
3️⃣ **团队协作**: 多人协作时共享备份信息
4️⃣ **定期清理**: 定期清理旧备份释放空间
5️⃣ **路径自定义**: 大项目可指定统一备份目录

**🔗 配合使用**:
• 修改前安全网: `excel_create_backup("技能表.xlsx")` 创建备份
• 版本对比: `excel_compare_sheets` 对比备份与当前版本
• 恢复操作: `excel_restore_backup` 从备份恢复数据
• 历史管理: `excel_list_backups` 查看所有历史版本

**⚠️ 安全提示**:
• 备份是修改前的最后保障，养成备份好习惯
• 重要数据建议保留多个历史版本
• 备份文件占用磁盘空间，注意定期清理
• 备份文件也应放在安全的位置，避免误删

**📈 最佳实践**:
• 修改前备份→修改→验证→确认成功
• 定期检查备份文件的可读性和完整性
• 为不同类型的配置表建立不同的备份策略
• 重要项目考虑建立自动化备份流程
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
💾 恢复备份 - 从备份还原Excel文件

**核心功能**: 将备份文件恢复到原位置或指定位置。修改前建议先用excel_create_backup创建备份。

**🎮 游戏开发场景**:
• **误操作回滚**: 改错配置后从备份恢复
• **版本回退**: 新版本有问题时回退到备份版本
• **数据灾难恢复**: 配置表损坏时从备份恢复
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
    """
📂 备份列表 - 查看文件的所有备份版本

**核心功能**: 列出指定Excel文件的所有历史备份，包含备份时间和文件大小。恢复前先用此工具选择备份版本。

**🎮 游戏开发场景**:
• **版本选择**: 修改前查看有哪些备份版本可选
• **备份清理**: 确认旧备份可以清理
• **团队协作**: 确认外包是否创建了备份
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
    """
➕ 插入空行 - 在指定位置插入空白行

**核心功能**: 在工作表指定行号前插入空行，后续行自动下移。适合需要腾出空间插入新数据的场景。

**🎮 游戏开发场景**:
• **表头扩展**: 在表头和数据之间插入分隔行或说明行
• **区域预留**: 为即将添加的数据预留空间
• **排序前置**: 插入行后填充数据再排序

**🔧 参数说明**:
• **row_index**: 插入位置（在此行之前插入）
• **count**: 插入行数
• **streaming**: True=流式写入（快）/ False=传统模式

**⚡ 使用建议**:
• 批量插入数据行请用excel_batch_insert_rows（按列名匹配，更智能）
• 插入前建议先用excel_get_range查看当前内容确认位置
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
➕ 插入空列 - 在指定位置插入空白列

**核心功能**: 在工作表指定列号前插入空列，后续列自动右移。适合需要新增数据维度的场景。

**🎮 游戏开发场景**:
• **新增属性列**: 给装备表添加"强化等级上限"等新属性列
• **多语言扩展**: 添加新的语言翻译列
• **数据分离**: 插入列后将数据从复合列拆分到独立列

**🔧 参数说明**:
• **column_index**: 插入位置（1-based，1=A列，在此列之前插入）
• **count**: 插入列数

**⚡ 使用建议**:
• 插入列后记得更新表头（用excel_update_range修改第1行）
• 插入位置不对不会自动撤销，建议先确认列号
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.insert_columns(file_path, sheet_name, column_index, count, streaming))


@mcp.tool()
@_track_call
def excel_find_last_row(
    file_path: str,
    sheet_name: str,
    column: Optional[Union[str, int]] = None
) -> Dict[str, Any]:
    """
📍 末行定位器 - 追加数据前的位置确认

**核心功能**: 快速查找工作表中最后一行有数据的行号。在追加新数据前使用此工具确定插入位置，避免覆盖已有数据。

**🎮 游戏开发场景**:
• **技能追加**: 新增技能前定位最后一行，避免覆盖
• **装备导入**: 批量导入装备数据前确认末行位置
• **数据清理**: 确认数据范围后进行清理操作
• **表格维护**: 定期检查表格数据量

**💡 使用建议**:
• 追加单行: `excel_find_last_row("技能表.xlsx", "技能表")` → 返回末行号 → 在末行+1写入
• 按列定位: `column="A"` 查找A列最后有数据的行
• 批量追加: 直接用`excel_batch_insert_rows`自动定位末行
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.find_last_row(file_path, sheet_name, column))


@mcp.tool()
@_track_call
def excel_create_file(
    file_path: str,
    sheet_names: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
🆕 创建文件 - 创建新的空Excel文件

**核心功能**: 创建新的空Excel文件，可指定工作表名称列表。创建后用excel_update_range写入数据。

**🎮 游戏开发场景**:
• **新建配置表**: 创建技能表/装备表/怪物表等标准配置文件
• **模板初始化**: 从零开始搭建配置表结构
• **自动化管线**: 脚本化创建配置文件

**🔧 参数说明**:
• **sheet_names**: 工作表名称列表（如["技能表","装备表"]），默认创建"Sheet"

**⚡ 使用建议**:
• 创建后用excel_update_range写入表头和数据
• 需要多工作表可在创建时指定sheet_names
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.create_file(file_path, sheet_names))


@mcp.tool()
@_track_call
def excel_export_to_csv(
    file_path: str,
    output_path: str,
    sheet_name: Optional[str] = None,
    encoding: str = "utf-8"
) -> Dict[str, Any]:
    """
📤 导出CSV - 将Excel工作表导出为CSV文件

**核心功能**: 将指定工作表导出为CSV格式，支持utf-8/gbk编码选择。

**🎮 游戏开发场景**:
• **版本控制**: 导出CSV后用git diff追踪配置变更（xlsx是二进制，diff不可读）
• **程序对接**: 导出CSV供游戏引擎或构建工具读取
• **数据分析**: 导出后用Python脚本做批量数值分析
• **配置备份**: 导出为纯文本格式做快照备份

**🔧 参数说明**:
• **encoding**: utf-8（默认）/ gbk（兼容旧版Excel中文）

**⚡ 使用建议**:
• 需要导出多个工作表请分别调用，或用excel_convert_format(json)一次导出全部
• 中文Excel导出有时需要gbk编码才能正常显示
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.export_to_csv(file_path, output_path, sheet_name, encoding))


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
📥 CSV导入 - 从CSV文件创建Excel文件

**核心功能**: 将CSV文件转换为Excel格式，支持自定义编码和表头处理。使用流式写入，大文件性能好。

**🎮 游戏开发场景**:
• **配置迁移**: 从其他工具（如Google Sheets导出、数据库导出）的CSV转换为Excel配置表
• **策划数据导入**: 策划用Excel编辑数据后导出CSV，再导入为标准格式
• **本地化导入**: 翻译团队提交的CSV多语言文件转为Excel配置
• **数据管线对接**: 自动化脚本生成的CSV报表转为Excel供策划查看

**🔧 参数说明**:
• **csv_path**: CSV文件路径
• **output_path**: 输出Excel文件路径
• **encoding**: 文件编码（中文CSV常用gbk，默认utf-8）
• **has_header**: CSV是否有表头行（默认True）

**⚡ 使用建议**:
• 中文CSV乱码时尝试encoding="gbk"
• 导入后建议用excel_describe_table检查表结构
• 需要追加到已有Excel请用excel_merge_files(append模式)
    """
    for _p in [csv_path, output_path]:
        _err = _validate_path(_p)
        if _err:
            return _err

    return _wrap(ExcelOperations.import_from_csv(csv_path, output_path, sheet_name, encoding, has_header))


@mcp.tool()
@_track_call
def excel_convert_format(
    input_path: str,
    output_path: str,
    target_format: str = "xlsx"
) -> Dict[str, Any]:
    """
🔄 格式转换 - Excel/CSV/JSON格式互转

**核心功能**: 将Excel文件转换为其他格式（xlsx/xlsm/csv/json），支持双向转换。

**🎮 游戏开发场景**:
• **xlsx→csv**: 导出配置表供版本控制diff（CSV比xlsx更易做文本对比）
• **xlsx→json**: 导出为JSON供程序直接读取或前端展示
• **csv→xlsx**: 将策划导出的CSV配置转为标准Excel格式
• **版本对比准备**: 将xlsx转为csv后用git diff检查配置变更

**🔧 参数说明**:
• **target_format**: 目标格式 xlsx/xlsm/csv/json

**⚡ 使用建议**:
• xlsx→csv适合版本管理（文本diff友好）
• 复杂格式（合并单元格、公式）转csv/json可能丢失信息
• json输出为行列二维数组格式
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
    """
🔗 合并文件 - 将多个Excel文件合并为一个

**核心功能**: 合并多个Excel文件，支持三种模式。适合将分散的配置表整合。

**🎮 游戏开发场景**:
• **sheets模式**: 将多个配置文件（技能表、装备表、怪物表）合并到一个文件，方便统一管理
• **append模式**: 合并外包团队分批提交的同类型配置（如多个版本的怪物表追加合并）
• **horizontal模式**: 将不同维度的数据横向拼接（如基础属性表+附加属性表按行对齐合并）

**📊 三种合并模式**:
• **sheets**: 每个输入文件作为独立工作表（适合不同类型的配置表）
• **append**: 纵向追加行（适合同结构配置表，如多个版本的道具表）
• **horizontal**: 横向拼接列（适合不同维度的数据合并）

**🔧 参数说明**:
• **input_files**: 输入文件路径数组（2个及以上）
• **merge_mode**: sheets/append/horizontal

**⚡ 使用建议**:
• append模式要求文件结构（列名）一致，否则数据会错位
• 合并前可用excel_compare_sheets检查结构差异
• 合并后建议用excel_check_duplicate_ids检查ID重复
    """
    for _f in input_files:
        _err = _validate_path(_f)
        if _err:
            return _err

    return _wrap(ExcelOperations.merge_files(input_files, output_path, merge_mode))


@mcp.tool()
@_track_call
def excel_get_file_info(file_path: str) -> Dict[str, Any]:
    """
ℹ️ 文件信息 - 获取Excel文件元数据

**核心功能**: 获取Excel文件的基本信息：文件大小、工作表列表、工作表数量、格式、创建/修改时间。

**🎮 游戏开发场景**:
• **文件确认**: 操作前确认文件格式和内容概况
• **配置盘点**: 快速了解项目有多少配置表
• **文件诊断**: 确认文件是否损坏（能否正常读取）
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.get_file_info(file_path))


@mcp.tool()
@_track_call
def excel_create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
📄 创建工作表 - 在Excel文件中新建工作表

**核心功能**: 在已有Excel文件中创建新工作表，可指定名称和位置。

**🎮 游戏开发场景**:
• **配置分类**: 将不同类型的配置放在不同工作表（如"技能表""装备表""怪物表"合在一个文件）
• **数据分区**: 将不同版本/活动的配置放在独立工作表
• **模板扩展**: 在配置文件中添加新的数据区域

**🔧 参数说明**:
• **index**: 插入位置（0=最前面，None=最后面）
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.create_sheet(file_path, sheet_name, index))


@mcp.tool()
@_track_call
def excel_delete_sheet(
    file_path: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
🗑️ 删除工作表 - 移除指定工作表

**核心功能**: 删除Excel文件中的指定工作表。⚠️ 删除不可恢复，建议先备份。

**🎮 游戏开发场景**:
• **清理废弃配置**: 删除已下线活动的工作表
• **文件瘦身**: 移除不需要的临时数据工作表
• **结构重组**: 合并文件后删除多余工作表

**⚡ 使用建议**:
• 删除前用excel_list_sheets确认工作表名
• 重要操作前先用excel_create_backup创建备份
• 至少保留一个工作表，否则会报错
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
    """
✏️ 重命名工作表 - 修改工作表名称

**核心功能**: 修改工作表名称。新名称不能与已有工作表重复。

**🎮 游戏开发场景**:
• **规范命名**: 将"Sheet1"改为"技能表"等有意义的名称
• **版本管理**: 给工作表加版本号后缀（如"装备表_v2"）
• **多语言标识**: 添加语言标识（如"怪物表_CN""怪物表_EN"）

**⚡ 使用建议**:
• 先用excel_list_sheets查看当前名称
• 新名称不能包含特殊字符（反斜杠/斜杠/问号/星号/方括号）
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.rename_sheet(file_path, old_name, new_name))


@mcp.tool()
@_track_call
def excel_copy_sheet(
    file_path: str,
    source_name: str,
    new_name: Optional[str] = None,
    index: Optional[int] = None
) -> Dict[str, Any]:
    """
📋 复制工作表 - 创建工作表副本（含数据和格式）

**核心功能**: 复制工作表的所有内容（数据+格式+公式），用于创建配置表变体。不指定new_name时自动生成"源表名_副本"。

**🎮 游戏开发场景**:
• **副本变体**: 复制怪物表做"精英版"怪物（属性翻倍）
• **活动版本**: 复制装备表做"活动版"装备（不同数值配置）
• **模板复用**: 基于已有配置表创建新类型配置
• **A/B测试**: 复制一份修改数值做对比测试

**🔧 参数说明**:
• **new_name**: 新工作表名称（默认"源表名_副本"）
• **index**: 插入位置（None=最后面）

**⚡ 使用建议**:
• 复制后用excel_update_range修改副本的数值
• 需要保留历史版本时建议用excel_create_backup替代
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.copy_sheet(file_path, source_name, new_name, index))


@mcp.tool()
@_track_call
def excel_rename_column(
    file_path: str,
    sheet_name: str,
    old_header: str,
    new_header: str,
    header_row: int = 1
) -> Dict[str, Any]:
    """
✏️ 重命名列 - 修改表头（列名）

**核心功能**: 修改指定列的表头值。支持双行表头（header_row=2修改英文字段名）。列名不存在时自动提示实际列名。

**🎮 游戏开发场景**:
• **字段重命名**: 将"atk"改为"attack_power"使列名更清晰
• **规范统一**: 统一不同策划使用的列名风格
• **多语言表头**: 修改第1行中文描述或第2行英文字段名

**🔧 参数说明**:
• **header_row**: 表头行号（双行表头设为2修改英文字段名）

**⚡ 使用建议**:
• 不确定列名时直接传一个可能的名字，系统会提示实际列名
• 重命名后相关引用（公式、代码）需要同步更新
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
🔄 智能 Upsert - 策划配置合并的核心工具

**核心功能**: 按键列查找行，存在则更新，不存在则插入。这是游戏配置表管理的核心操作，实现"合并导入"的智能逻辑。

**🎮 游戏开发场景**:
• **技能配置合并**: 导入新技能时，技能ID已存在则更新属性，不存在则新增
• **装备批量更新**: 导入装备数据时，按装备ID智能更新或新增
• **怪物数据同步**: 批量更新怪物信息，存在的更新属性，不存在的添加
• **配置表维护**: 版本升级时配置的智能合并和更新

**🔧 Upsert逻辑**:
1️⃣ **查找**: 按key_column和key_value查找指定行
2️⃣ **判断**: 存在→更新指定字段；不存在→插入新行
3️⃣ **合并**: 只更新指定字段，不修改其他已有字段
4️⃣ **效率**: 一次操作完成查找和更新，比分开操作更快

**📋 参数详解**:
• **key_column**: 匹配列名（如"skill_id"、"equip_id"）
• **key_value**: 匹配值（如1001、"T_001"）
• **updates**: 要更新的字段映射（如{"damage": 200, "cooldown": 3}）
• **header_row**: 表头行号（默认1，支持双行表头）
• **streaming**: 流式写入，大文件性能更好

**🔍 执行效果**:
• **更新场景**: 找到技能ID=1001的行，只更新damage和coolddown字段
• **插入场景**: 未找到技能ID=1002的行，创建新行包含所有updates数据
• **混合场景**: 同一批数据中部分更新、部分新增

**💡 使用示例**:
• 技能更新: `excel_upsert_row("技能表.xlsx", "技能表", "skill_id", 1001, {"damage": 200, "cooldown": 3})`
• 装备导入: `excel_upsert_row("装备表.xlsx", "装备表", "equip_id", "E001", {"品质": "传说", "价格": 10000})`
• 怪物同步: `excel_upsert_row("怪物表.xlsx", "怪物表", "monster_id", 500, {"血量": 1000, "攻击": 50})`

**⚡ 性能优势**:
• **智能缓存**: 查找后自动缓存，重复操作更快
• **流式写入**: streaming=True对大文件性能提升显著
• **批量优化**: 比分开执行"查找+更新/插入"操作更高效
• **内存控制**: 按需加载，不消耗过多内存

**🛡️ 安全机制**:
• **字段映射**: 只更新指定字段，不会误改其他数据
• **类型检查**: 自动处理数据类型转换
• **错误提示**: 查找失败时提供详细的错误信息
• **事务保护**: 失败时自动回滚，确保文件完整性

**🔗 配合使用**:
• 数据验证: upsert前用excel_query确认数据状态
• 合并预览: 先小批量测试确认逻辑正确
• 完整更新: 大规模数据更新前建议备份
• 结果确认: upsert后用excel_query确认结果

**🚧 注意事项**:
• 确保key_column是唯一标识字段（如ID、编码等）
• 双行表头会自动处理，列名映射智能匹配
• 更新只针对指定字段，不影响其他数据
• 大文件建议使用streaming提升性能
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.upsert_row(file_path, sheet_name, key_column, key_value, updates, header_row, streaming))


@mcp.tool()
@_track_call
def excel_batch_insert_rows(
    file_path: str,
    sheet_name: str,
    data: List[Dict[str, Any]],
    header_row: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """
📦 批量插入行 - 将多行数据追加到工作表末尾

**核心功能**: 批量导入数据行，自动按列名匹配写入。streaming=True（默认）使用流式写入，大文件性能更好，但不保留单元格格式。

**🎮 游戏开发场景**:
• **批量配置导入**: 策划一次导入几十条技能/装备/怪物配置数据
• **活动数据填充**: 批量添加限时活动道具、任务奖励等配置行
• **版本合并**: 将外包团队的新配置行合入主表
• **数据迁移**: 从其他系统导出的配置批量导入Excel

**📊 返回信息**:
• **inserted_count**: 实际插入的行数
• **start_row/end_row**: 插入的起始/结束行号
• **unknown_columns**: 表中不存在的列名（数据被忽略）
• **mode**: streaming（流式）或 standard（传统）

**🔧 参数说明**:
• **data**: 行数据数组，每行为{列名: 值}字典（列名需与表头一致）
• **header_row**: 表头所在行号（默认1，双行表头设为2）
• **streaming**: True=流式写入（快，不保留格式）/ False=传统模式（保留格式）

**⚡ 使用建议**:
• 先用excel_get_headers确认列名，避免unknown_columns
• 超过100行数据时streaming模式优势明显（内存降低90%+）
• 需要按ID更新已有行请用excel_upsert_row
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.batch_insert_rows(file_path, sheet_name, data, header_row, streaming))


@mcp.tool()
@_track_call
def excel_delete_rows(
    file_path: str,
    sheet_name: str,
    row_index: int,
    count: int = 1,
    streaming: bool = True
) -> Dict[str, Any]:
    """
🗑️ 删除行 - 删除工作表中指定位置的行

**核心功能**: 删除指定行号范围的行，后续行自动上移。streaming=True（默认）使用流式写入。

**🎮 游戏开发场景**:
• **清理废弃配置**: 删除已下线活动的道具/任务配置行
• **版本裁剪**: 删除测试用的临时配置数据
• **批量清理**: 配合excel_query找出不需要的行号后批量删除
• **ID重构**: 删除重复ID行（配合excel_check_duplicate_ids使用）

**🔧 参数说明**:
• **row_index**: 起始行号（1-based）
• **count**: 删除行数（默认1）
• **streaming**: True=流式写入（快，不保留格式）/ False=传统模式

**⚡ 使用建议**:
• 删除前建议先用excel_get_range查看目标行的内容，避免误删
• 大量删除（>100行）时streaming模式优势明显
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
    """
🗑️ 删除列 - 删除工作表中指定位置的列

**核心功能**: 删除指定列号范围的列，后续列自动左移。streaming=True（默认）使用流式写入。

**🎮 游戏开发场景**:
• **清理废弃字段**: 删除已下线功能的配置列（如"旧版技能ID"）
• **字段精简**: 移除不需要的中间计算列
• **结构重组**: 合并配置表后删除重复列

**🔧 参数说明**:
• **column_index**: 起始列号（1-based，1=A列）
• **count**: 删除列数
• **streaming**: True=流式写入（快，不保留格式）/ False=传统模式

**⚡ 使用建议**:
• 删除前建议先用excel_get_range查看列内容，避免误删
• 大量列删除（>10列）时streaming模式优势明显
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
    """
🧮 设置公式 - 在单元格中写入Excel公式

**核心功能**: 在指定单元格设置公式（不含等号），保存后返回计算结果。公式会随数据变化自动更新。

**🎮 游戏开发场景**:
• **自动计算**: 设置"总攻击=基础攻击+装备攻击"等自动计算列
• **统计汇总**: 在表末尾添加SUM/AVERAGE/COUNT等汇总行
• **条件判断**: 用IF公式做条件判断（如"IF(等级>10,'高级','普通')"）

**🔧 参数说明**:
• **cell_address**: 目标单元格（如"A1"）
• **formula**: 公式（不含等号，如"SUM(A1:A10)"）

**⚡ 使用建议**:
• 公式会随文件保存，下次打开仍有效
• 只做临时计算不修改文件请用excel_evaluate_formula
• 复杂公式建议先在Excel中验证再写入
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
🧮 计算公式 - 临时执行公式（不修改文件）

**核心功能**: 临时执行Excel公式并返回结果，不写入文件。可做快速计算器，支持基础统计函数和范围引用。

**🎮 游戏开发场景**:
• **快速计算**: "这个怪物的DPS是多少？" → 直接算不用打开文件
• **数值验证**: 验证策划给的数值公式是否正确
• **平衡估算**: "如果攻击力提升20%，DPS变化多少？"

**🔧 支持的公式**:
• 算术: `100*1.2+50`
• 范围统计: `SUM(A1:A10)`, `AVERAGE(B1:B50)`, `MAX(C1:C100)`
• 高级统计: `MEDIAN`, `STDEV`, `PERCENTILE`, `COUNTIF`, `SUMIF`

**⚡ 使用建议**:
• 不需要引用文件数据的纯计算直接写公式即可
• 需要引用文件数据时配合context_sheet参数指定工作表
• 结果有缓存，重复计算更快
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
    """
🚀 SQL查询引擎 - 游戏配置表的超强分析工具

**核心功能**: 执行标准SQL查询Excel数据，支持复杂分析、聚合统计、关联查询。这是最重要的数据分析工具，优先使用它而非excel_get_range。

**🎮 游戏开发场景**:
• 数值平衡: `SELECT 类型, AVG(伤害/冷却) as dpm FROM 技能 GROUP BY 类型`
• 数据质量: `SELECT * FROM 怪物 WHERE 血量 < 10 OR 攻击 = 0`
• 版本对比: `SELECT skill_id, skill_name, damage as 旧版 FROM 技能_v1 UNION SELECT skill_id, skill_name, damage as 新版 FROM 技能_v2`
• 关联分析: `SELECT 技能名, 装备名 FROM 技能 a JOIN 装备 b ON a.equipment_id = b.id`

**📊 支持语法**:
• **基础查询**: SELECT/DISTINCT/别名(AS)/数学表达式(伤害*1.2)
• **条件过滤**: WHERE/AND/OR/LIKE/IN/NOT IN/BETWEEN/IS NULL/NOT
• **高级查询**: WHERE子查询/FROM子查询/CASE WHEN/COALESCE/EXISTS/CTE/UNION/窗口函数
• **聚合统计**: COUNT/SUM/AVG/MAX/MIN/GROUP BY/HAVING/ORDER BY/LIMIT/OFFSET
• **关联查询**: INNER JOIN/LEFT JOIN/RIGHT JOIN/CROSS JOIN（同文件内工作表）
• **字符串函数**: UPPER/LOWER/TRIM/LENGTH/CONCAT/REPLACE/SUBSTRING
• **双行表头**: 自动识别中文描述+英文字段名，支持中英文混合查询

**🔍 独特优势**:
• 中文列名友好: `SELECT 技能名称, 伤害值 FROM 技能 WHERE 等级 > 10`
• 智能缓存: 同文件重复查询提速30-100倍
• 错误提示: 拼写错误时自动推荐相似列名
• 数据安全: 只读操作，不修改文件

**📤 输出格式**:
• table: Markdown表格格式（默认）
• json: 结构化JSON数据
• csv: CSV逗号分隔格式

**💡 使用建议**:
• 新手: 先用excel_describe_table了解表结构
• 复杂分析: 使用GROUP BY和HAVING进行统计
• 大表查询: 利用缓存提升后续查询速度
• 中文查询: 直接使用双行表头的中文描述
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
        return _wrap(execute_advanced_sql_query(
            file_path=file_path,
            sql=query_expression,
            sheet_name=None,  # 统一使用SQL FROM子句中的表名
            limit=None,  # 统一使用SQL中的LIMIT
            include_headers=include_headers,
            output_format=output_format or 'table'
        ))

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
    """
⚙️ SQL批量修改器 - 精确控制数据变更

**核心功能**: 使用SQL语句进行条件批量修改，支持复杂条件判断和安全预览。这是批量修改的首选工具，比手动修改更安全高效。

**🎮 游戏开发场景**:
• **数值调整**: `UPDATE 技能 SET 伤害 = 伤害 * 1.2 WHERE 元素 = '火'` (火系技能伤害+20%)
• **版本升级**: `UPDATE 装备 SET 等级 = 等级 + 1 WHERE 品质 = '传说'` (传说装备升级)
• **数据修正**: `UPDATE 怪物 SET 血量 = 100 WHERE 血量 < 10` (修正异常血量数据)
• **平衡性调整**: `UPDATE 技能 SET 冷却 = 冷却 - 1 WHERE 职业 = '战士'` (战士技能优化)

**🔧 SET语法支持**:
• **常量值**: `SET 伤害 = 500` 直接设置固定值
• **算术表达式**: `SET 伤害 = 伤害 * 1.1` (当前值×1.1)
• **数学运算**: `SET 冷却 = 冷却 + 1, 消耗 = 消耗 - 5` (多字段同时修改)
• **条件更新**: `SET 品质 = '精品' WHERE 原价 > 10000`

**🔍 WHERE语法支持**:
• **基础条件**: `WHERE 元素 = '火'` / `WHERE 等级 > 10`
• **范围条件**: `WHERE 血量 BETWEEN 50 AND 100`
• **模糊匹配**: `WHERE 技能名 LIKE '%治疗%'` / `WHERE 装备名 LIKE '传说%'`
• **多条件**: `WHERE 元素 IN ('火', '水', '风')` / `WHERE 职业 = '战士' AND 等级 >= 20`
• **空值处理**: `WHERE 描述 IS NULL` / `WHERE 副作用 IS NOT NULL`
• **复合条件**: `WHERE 元素 = '火' AND 等级 > 5`

**🛡️ 安全机制**:
• **预览模式**: dry_run=True只预览修改范围，不实际修改
• **事务保护**: 失败自动回滚，确保文件不损坏
• **备份机制**: 修改前自动创建备份
• **错误提示**: 详细错误信息，AI可自动修复

**💡 最佳实践**:
1️⃣ **先预览后执行**: `dry_run=True` 预览修改范围和数据影响
2️⃣ **精准条件**: 用明确的WHERE条件避免误修改
3️⃣ **小批量测试**: 先测试小范围数据验证修改逻辑
4️⃣ **备份确认**: 重要修改前确认文件已备份
5️⃣ **验证结果**: 修改后用excel_query确认修改成功

**🚫 注意事项**:
• 只支持UPDATE语句，不支持INSERT/DELETE
• 修改前建议先用excel_describe_table了解数据结构
• 复杂修改建议分步骤执行，避免一次修改过多数据
• 大文件修改可能较慢，建议错峰执行

**⚡ 配合使用**:
• 修改前预览: dry_run=True确认修改范围
• 修改后验证: excel_query检查修改结果
• 安全保障: excel_create_backup创建备份
• 版本对比: excel_compare_sheets对比版本差异
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
def excel_describe_table(
    file_path: str,
    sheet_name: str = None
) -> Dict[str, Any]:
    """
📋 表结构分析器 - 数据探索的第一步

**核心功能**: 快速分析Excel工作表结构，了解列名、数据类型、数据分布。这是所有数据分析的第一步，使用前必查！

**🎮 游戏开发场景**:
• 技能表分析: 了解技能ID、名称、类型、伤害、冷却等字段结构和类型
• 装备表审查: 检查装备ID、品质、属性、套装等字段的完整性
• 怪物数据验证: 确认怪物ID、等级、血量、攻击等字段的类型和分布
• 配置表检查: 验证字段命名规范和数据类型一致性

**🔍 智能检测**:
• **双行表头自动识别**: 第1行中文描述+第2行英文字段名，自动中英映射
• **数据类型推断**: 自动识别number/text/date/mixed类型
• **空值统计**: 准确统计每列的非空值数量
• **样本展示**: 每列显示前3个实际值，了解数据内容
• **行数统计**: 快速获取总行数，把握数据规模

**📊 返回信息**:
• 字段名: 英文字段名（如skill_name）
• 数据类型: number/文字/日期/混合
• 非空统计: 实际有数据的单元格数量
• 示例值: 每列的前3个真实值
• 中英映射: 双行表头时自动映射中文描述
• 总行数: 整个表的数据行数量

**💡 使用流程**:
1️⃣ **初见新表**: `excel_describe_table("技能表.xlsx")` 了解整体结构
2️⃣ **聚焦重点**: 指定sheet_name分析特定工作表
3️⃣ **数据规划**: 根据类型选择合适的查询和分析方法
4️⃣ **质量检查**: 通过非空统计和示例值发现数据问题

**🎯 配合使用**:
• 后续查询: `excel_query`基于此表结构进行数据检索
• 修改操作: `excel_update_query`根据类型进行安全修改
• 版本对比: `excel_compare_sheets`对比不同版本的结构变化
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    if not file_path or not file_path.strip():
        return _fail('文件路径不能为空', meta={"error_code": "MISSING_FILE_PATH"})

    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    except Exception as e:
        return _fail(f'无法打开文件: {e}', meta={"error_code": "FILE_OPEN_FAILED"})

    try:
        # 选择工作表
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                return _fail(f'工作表 "{sheet_name}" 不存在。可用工作表: {wb.sheetnames}', meta={"error_code": "SHEET_NOT_FOUND"})
            ws = wb[sheet_name]
        else:
            ws = wb.worksheets[0]
            sheet_name = ws.title

        # 读取前几行来判断表头类型
        rows = list(ws.iter_rows(max_row=4, values_only=True))
        if not rows:
            return _fail('工作表为空', meta={"error_code": "EMPTY_SHEET"})

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
        return _ok(f"表 '{sheet_name}': {len(columns)}列, {row_count}行数据, {'双行表头' if is_dual_header else '单行表头'}", data={
            'sheet_name': sheet_name,
            'header_type': 'dual' if is_dual_header else 'single',
            'row_count': row_count,
            'column_count': len(columns),
            'columns': columns
        }, meta={"file_path": file_path, "sheet": sheet_name})
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
    preset: Optional[str] = None
) -> Dict[str, Any]:
    """
🎨 单元格样式美化器 - 配置表的可视化

**核心功能**: 格式化单元格样式，支持快速预设和完全自定义。让配置表更易读、更专业。

**🎮 游戏开发场景**:
• **异常高亮**: 用"warning"预设标记异常数值配置
• **版本标记**: 用"highlight"标记新版本修改的数据
• **品质区分**: 不同品质装备用不同背景色
• **数据审查**: 标记需要检查的数据行

**⚡ 快速预设** (preset参数):
• **highlight**: 黄色背景，标记待检查的数据
• **warning**: 红色背景，标记异常或问题数据
• **success**: 绿色背景，标记已确认或正常数据

**🔧 自定义格式** (formatting参数):
• **字体**: bold/italic/underline/size/color
• **背景色**: bg_color（十六进制颜色值）
• **边框**: border_style/border_color
• **对齐**: horizontal/vertical alignment
• **数字格式**: number_format（如"#,##0"千分位）

**💡 使用建议**:
• 快速标记→用preset参数一键高亮
• 边框设置→用excel_set_borders更专业
• 条件格式→结合excel_query筛选后再格式化
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.format_cells(file_path, sheet_name, range, formatting, preset))


@mcp.tool()
@_track_call
def excel_merge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
🔗 合并单元格 - 将范围合并为一个大单元格

**核心功能**: 将指定范围的单元格合并为一个，仅保留左上角的值。常用于标题行、分类标签等场景。

**🎮 游戏开发场景**:
• **标题行**: 合并A1:E1作为配置表标题（如"技能配置表 v2.0"）
• **分类标签**: 在数据区域插入分类行并合并（如"近战武器"横跨多列）
• **表头美化**: 双行表头中合并描述行

**🔧 参数说明**:
• **range**: 范围表达式（如"A1:E1"）

**⚡ 使用建议**:
• 合并后只有左上角单元格的值保留，其他单元格值丢失
• 取消合并请用excel_unmerge_cells
• 合并后的单元格不能单独编辑子区域
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.merge_cells(file_path, sheet_name, range))


@mcp.tool()
@_track_call
def excel_unmerge_cells(
    file_path: str,
    sheet_name: str,
    range: str
) -> Dict[str, Any]:
    """
🔓 取消合并 - 恢复已合并的单元格为独立单元格

**核心功能**: 取消指定范围内的合并，恢复为独立单元格。合并前每个子单元格的值会变为空（只有左上角保留了原始值）。

**🎮 游戏开发场景**:
• **结构调整**: 需要在合并区域中单独编辑某个单元格时先取消合并
• **数据提取**: 取消合并后才能单独读取子区域数据
• **格式重置**: 重新规划表头结构前取消所有合并

**⚡ 使用建议**:
• 取消合并前先确认不需要保留合并效果
• 合并请用excel_merge_cells
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.unmerge_cells(file_path, sheet_name, range))


@mcp.tool()
@_track_call
def excel_set_borders(
    file_path: str,
    sheet_name: str,
    range: str,
    border_style: str = "thin"
) -> Dict[str, Any]:
    """
📏 设置边框 - 为指定范围添加单元格边框

**核心功能**: 为指定范围的单元格添加边框线。支持多种样式，简单高亮也可用excel_format_cells的preset参数。

**🎮 游戏开发场景**:
• **表格分隔**: 给汇总行添加上下边框区分数据区
• **区域标注**: 给特定区域添加边框（如"注意：以下数据需要审核"）
• **打印美化**: 给需要打印的配置表添加边框线

**🔧 边框样式**:
• **thin**: 细线（默认，最常用）
• **thick**: 粗线（强调分隔）
• **medium**: 中等（介于thin和thick之间）
• **double**: 双线（标题分隔）
• **dotted**: 点线（装饰性）
• **dashed**: 虚线（辅助线）

**⚡ 使用建议**:
• 只需要简单高亮背景色请用excel_format_cells的preset参数
• 边框对整个范围统一应用，不支持单边设置
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
📏 设置行高 - 调整行的高度

**核心功能**: 调整指定行的高度（磅值）。可同时调整连续多行。

**🎮 游戏开发场景**:
• **标题行加高**: 标题行高度设为30-40磅更醒目
• **数据行紧凑**: 数据行高度设为18-20磅更紧凑
• **说明行**: 多行文字的说明行需要更高（如40-60磅）

**🔧 参数说明**:
• **height**: 行高（磅值，默认约15磅）
• **count**: 同时调整的连续行数
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
📏 设置列宽 - 调整列的宽度

**核心功能**: 调整指定列的宽度（字符单位）。可同时调整连续多列。

**🎮 游戏开发场景**:
• **ID列收窄**: ID列通常只需要10-12字符宽度
• **描述列加宽**: 描述/备注列需要30-50字符宽度
• **统一列宽**: 一次性调整所有数据列为统一宽度

**🔧 参数说明**:
• **column_index**: 起始列号（1-based）
• **width**: 列宽（字符单位，默认约8.43字符）
• **count**: 同时调整的连续列数

**⚡ 使用建议**:
• 中文字符约占2个字符宽度
• 调整后如果显示"####"说明宽度不够
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    return _wrap(ExcelOperations.set_column_width(file_path, sheet_name, column_index, width, count))


# ==================== Excel比较功能 ====================

@mcp.tool()
@_track_call
def excel_compare_files(
    file1_path: str,
    file2_path: str
) -> Dict[str, Any]:
    """
🔍 文件对比 - 逐单元格比较两个Excel文件的所有工作表差异

**核心功能**: 深度对比两个文件的所有工作表，输出结构差异和逐单元格值变化（旧值→新值）。适合检查配置表整体改动。

**🎮 游戏开发场景**:
• **版本diff**: 对比新旧版本配置表，快速了解策划改了哪些数值
• **外包验收**: 对比外包交付的配置表与原始模板的差异
• **回归检查**: 更新配置后对比前后版本，确认只改了预期内容
• **多语言校验**: 对比中文和英文配置表的结构一致性

**📊 返回信息**:
• **结构差异**: 新增/删除的工作表、新增/删除的列
• **值变化**: 每个单元格的旧值→新值（按工作表分组）
• **统计**: 总差异单元格数

**⚡ 使用建议**:
• 只关心记录级变化（新增/删除/修改的行）→ 用excel_compare_sheets
• 本工具是单元格级对比，适合精确检查每个数值变化
• 文件较大时对比可能较慢，建议只对比特定工作表
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
    """
🔍 检查重复ID - 扫描配置表ID列，返回重复值及所在行号

**核心功能**: 快速检测指定列的重复值，返回每个重复值出现的行号。数据质量检查必备。

**🎮 游戏开发场景**:
• **配置入库前校验**: 导入新配置前检查技能ID/装备ID是否重复（重复会导致游戏运行时覆盖）
• **合并后验证**: 合并多个配置表后检查ID冲突
• **外包交付验收**: 验证外包提交的配置表是否有ID重复（常见低级错误）
• **策划自查工具**: 策划批量编辑后快速发现重复配置

**📊 返回信息**:
• **duplicates**: 重复值列表（值→行号数组）
• **total_duplicates**: 重复值总数
• **affected_rows**: 受影响的行数

**🔧 参数说明**:
• **id_column**: 列号（数字）或列名（字符串），默认第1列
• **header_row**: 表头行号（双行表头设为2）

**⚡ 使用建议**:
• 也可用SQL: SELECT ID, COUNT(*) as c FROM 表 GROUP BY ID HAVING c>1
• 建议在每次配置表修改后执行，防止重复ID进入版本库
• 检查多列组合唯一性可用excel_query的GROUP BY多列
    """
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
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
    """
📊 智能工作表对比 - 版本差异分析器

**核心功能**: 按ID列精确对比两个工作表，输出对象级差异（新增/删除/修改的记录）。游戏配置版本管理的核心工具。

**🎮 游戏开发场景**:
• **版本对比**: 对比v1.0和v1.1的技能配置表差异
• **数值审核**: 对比策划提交的数值修改和当前版本
• **合并验证**: 确认多人修改的配置表没有冲突
• **回滚检查**: 对比备份版本和当前版本确认回滚范围

**🔍 对比结果**:
• **新增记录**: 新版本中有但旧版本中没有的数据
• **删除记录**: 旧版本中有但新版本中没有的数据
• **修改记录**: 两个版本都存在但字段值不同的数据
• **未变记录**: 两个版本完全相同的数据

**📋 参数说明**:
• **id_column**: 用于匹配行的标识列（默认第1列，支持列名如"skill_id"）
• **header_row**: 表头行号（默认1，支持双行表头）
• **对比维度**: 文件级（跨文件对比）+ 工作表级（同文件不同表）

**💡 使用建议**:
1️⃣ **版本对比**: 对比修改前备份和当前版本
2️⃣ **数值审核**: 对比策划提交的修改和基准版本
3️⃣ **合并验证**: 多人协作时确认修改一致性
4️⃣ **数据迁移**: 验证数据迁移的正确性

**🔗 配合使用**:
• 修改前备份→对比→确认差异→发布
• 策划提交数值→与基准版对比→审核通过→合并
• 配合excel_query筛选特定类型的差异记录

**⚠️ 注意事项**:
• 逐单元格对比请用excel_compare_files
• 大表对比可能较慢，建议缩小范围
• ID列必须是唯一标识，否则对比结果不准确
    """
    for _p in [file1_path, file2_path]:
        _err = _validate_path(_p)
        if _err:
            return _err
    return _wrap(ExcelOperations.compare_sheets(file1_path, sheet1_name, file2_path, sheet2_name, id_column, header_row))


@mcp.tool()
@_track_call
def excel_server_stats() -> Dict[str, Any]:
    """
📊 服务器性能监控 - Excel MCP的健康守护者

**核心功能**: 获取MCP服务器运行统计：每个工具的调用次数、平均耗时、错误率和错误分类。返回全局error_types统计（按错误类型分类的计数），用于监控和调试。

**🎮 游戏开发场景**:
• **性能瓶颈分析**: 监控技能表查询、装备表更新等高频工具的性能表现，识别慢操作
• **错误率监控**: 追踪数据导入、表结构修改等关键操作的错误率，确保数据质量
• **资源使用追踪**: 监控大文件处理、批量数据操作时的内存和CPU使用情况
• **API调用统计**: 分析游戏配置管理、数据同步等场景的API调用模式，优化接口设计

**返回信息**: 工具调用统计、错误分类统计、性能指标、API调用频率
**参数说明**: 无参数，返回全局服务器统计信息
**使用建议**: 在进行大规模数据处理前检查服务器状态，性能异常时及时干预
    """
    stats = _tracker.get_stats()
    return _ok("服务器统计信息", data=stats)


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
