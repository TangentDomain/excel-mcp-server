"""
高级SQL查询引擎 - 基于SQLGlot实现的SQL查询支持

支持功能:
- 基础查询: SELECT, DISTINCT, 别名
- 条件筛选: WHERE, LIKE, IN, BETWEEN, AND/OR, EXISTS, 子查询
- 聚合统计: COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 排序限制: ORDER BY, LIMIT, OFFSET
- 算术运算: 加减乘除
- 条件表达式: CASE WHEN, COALESCE/IFNULL
- 表关联: INNER JOIN, LEFT JOIN, RIGHT JOIN, FULL JOIN, CROSS JOIN(同文件内工作表关联 + 跨文件关联)
- 子查询: WHERE col IN (SELECT ...), 标量子查询, EXISTS
- CTE: WITH ... AS (SELECT ...)
- 字符串函数: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT
- 窗口函数: ROW_NUMBER, RANK, DENSE_RANK(OVER PARTITION BY ... ORDER BY ...)
- 集合操作: UNION, UNION ALL, EXCEPT, INTERSECT

不支持功能:
- FROM子查询(FROM (SELECT ...) AS alias)
"""

import os
import re
import json
import csv
import io
import time
import difflib
import operator
import shutil
import tempfile
import logging
from contextlib import contextmanager
from typing import Dict, List, Any, Optional, Union, Tuple, Generator
import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)

# SQLGlot导入 - 核心SQL解析引擎
try:
    import sqlglot
    from sqlglot import expressions as exp
    from sqlglot.errors import ParseError, UnsupportedError
    SQLGLOT_AVAILABLE = True
except ImportError:
    SQLGLOT_AVAILABLE = False
    logger.warning("SQLGlot未安装,将使用基础pandas查询功能")
    # 创建虚拟类型注解以避免NameError
    class exp:
        class Expression: pass
        class Select: pass
        class Subquery: pass
        class With: pass
        class Window: pass
        class From: pass
        class Table: pass
        class Where: pass
        class Column: pass
        class Literal: pass
        class EQ: pass
        class NEQ: pass
        class GT: pass
        class GTE: pass
        class LT: pass
        class LTE: pass
        class And: pass
        class Or: pass
        class Like: pass
        class In: pass
        class Paren: pass  # 括号表达式
        # class IsNull: pass  # SQLGlot中可能不使用这个名称
        # class NotNull: pass  # SQLGlot中可能不使用这个名称
        class Order: pass
        class Ordered: pass
        class Having: pass
        class Alias: pass
        class AggFunc: pass

# Excel处理导入
import openpyxl
from pathlib import Path

# 流式写入导入
try:
    from ..core.streaming_writer import StreamingWriter
except ImportError:
    StreamingWriter = None

# 配置常量
from ..utils.config import (
    MAX_CACHE_SIZE,
    MAX_QUERY_CACHE_SIZE,
    QUERY_CACHE_TTL,
    CACHE_TARGET_MEMORY_MB,
    MAX_RESULT_ROWS,
    STREAMING_WRITE_MIN_ROWS,
    STREAMING_WRITE_MIN_CHANGES,
    STREAMING_WRITE_MIN_FILE_SIZE_MB,
    MARKDOWN_TABLE_MAX_ROWS
)


class StructuredSQLError(Exception):
    """结构化SQL错误,为AI提供可自动修复的错误信息.
    
    Attributes:
        error_code: 机器可读的错误分类码
        message: 人类可读的错误描述
        hint: AI修复建议(可选)
        context: 错误上下文,如可用列名,表名等(可选)
    """
    def __init__(self, error_code: str, message: str, hint: str = "", context: dict = None):
        self.error_code = error_code
        self.message = message
        self.hint = hint
        self.context = context or {}
        super().__init__(message)


def _unsupported_error_hint(err_detail: str) -> str:
    """为UnsupportedError提供替代建议."""
    err_upper = err_detail.upper()
    
    if 'INSERT' in err_upper or 'DELETE' in err_upper or 'DROP' in err_upper or 'ALTER' in err_upper or 'CREATE' in err_upper:
        return '此工具仅支持SELECT查询.数据修改请使用excel_update_query(UPDATE语句).'
    if 'NATURAL JOIN' in err_upper:
        return '不支持NATURAL JOIN,请改用显式ON条件:JOIN 表2 ON 表1.列 = 表2.列'
    if 'FETCH' in err_upper or 'NEXT' in err_upper:
        return '不支持FETCH/NEXT语法,请用LIMIT:SELECT ... LIMIT 10'
    if 'RECURSIVE' in err_upper:
        return '不支持递归CTE(WITH RECURSIVE).请改用普通CTE或子查询.'
    if 'LATERAL' in err_upper:
        return '不支持LATERAL JOIN.请改用子查询或CTE.'
    if 'WINDOW' in err_upper and 'OVER' not in err_upper:
        return 'WINDOW子句请改为直接在窗口函数后写OVER:ROW_NUMBER() OVER (PARTITION BY ... ORDER BY ...)'
    
    return '请参考工具描述中"SQL已支持功能"列表,使用支持的操作.对于复杂计算,可考虑分步查询.'


def _parse_error_hint(err_str: str, sql: str) -> str:
    """根据SQLGlot ParseError和原始SQL,生成AI可自动修复的提示.
    
    覆盖常见SQL书写错误:
    - 关键字拼写错误
    - 关键字顺序错误
    - 缺少关键字
    - 缺少逗号/括号
    - Excel函数名误用
    - 中文标点混用
    - 引号配对错误
    """
    sql_upper = sql.strip().upper()
    hint = ""
    
    # === 关键字拼写错误 ===
    typos = [
        ('SELEC ', 'SELECT'), ('SELEC$', 'SELECT'),
        ('FORM ', 'FROM'), ('FORM$', 'FROM'),
        ('WHER ', 'WHERE'), ('WHER$', 'WHERE'),
        ('GROUPBY', 'GROUP BY'), ('GROUP  BY', 'GROUP BY'),
        ('ORDERBY', 'ORDER BY'), ('ORDER  BY', 'ORDER BY'),
        ('HAVNIG', 'HAVING'), ('HAVIN', 'HAVING'),
        ('INNTER', 'INNER'), ('LEFR', 'LEFT'), ('RIGTH', 'RIGHT'),
        ('JOINT', 'JOIN'), ('OUDER', 'OUTER'),
        ('DISTIN T', 'DISTINCT'), ('DISTNCT', 'DISTINCT'),
        ('BETWEE N', 'BETWEEN'), ('BETWEN', 'BETWEEN'),
        ('NOTNULL', 'NOT NULL'), ('ISNUL', 'IS NULL'),
        ('LIK E', 'LIKE'), ('LIEK', 'LIKE'),
        ('EXIS TS', 'EXISTS'), ('EXIST ', 'EXISTS'),
        ('LIMITT', 'LIMIT'), ('OFFEST', 'OFFSET'),
        ('ASCEND', 'ASC'), ('DSCEND', 'DESC'),
        ('CROS', 'CROSS'), ('FUL L', 'FULL'),
        ('UNIO N', 'UNION'), ('UNON', 'UNION'),
        ('INTERSE CT', 'INTERSECT'), ('EXCEP T', 'EXCEPT'),
        ('CONCATENATE', 'CONCAT'), ('SUBSTITUE', 'REPLACE'),
    ]
    for typo, correct in typos:
        if typo.rstrip('$') in sql_upper:
            # 用$匹配行尾
            if typo.endswith('$') and not sql_upper.rstrip(';').endswith(typo.rstrip('$')):
                continue
            hint = f'可能是拼写错误,"{typo.rstrip().rstrip("$")}" 应为 "{correct}"'
            return hint
    
    # === 关键字顺序错误 ===
    # SELECT ... FROM ... WHERE ... GROUP BY ... HAVING ... ORDER BY ... LIMIT
    order_keywords = ['SELECT', 'FROM', 'WHERE', 'GROUP BY', 'HAVING', 'ORDER BY', 'LIMIT']
    found_positions = []
    for kw in order_keywords:
        # GROUP BY / ORDER BY 需要特殊处理
        if ' ' in kw:
            parts = kw.split()
            pos = sql_upper.find(parts[0])
            if pos != -1:
                # 检查后面是否跟着第二个词
                after = sql_upper[pos+len(parts[0]):].lstrip()
                if after.startswith(parts[1]):
                    found_positions.append((pos, kw))
        else:
            pos = sql_upper.find(kw)
            if pos != -1 and (pos == 0 or not sql_upper[pos-1].isalpha()):
                found_positions.append((pos, kw))
    
    # 按出现位置排序,然后检查顺序是否符合SQL标准
    found_positions.sort(key=lambda x: x[0])
    for i in range(len(found_positions) - 1):
        pos1, kw1 = found_positions[i]
        pos2, kw2 = found_positions[i + 1]
        idx1 = order_keywords.index(kw1)
        idx2 = order_keywords.index(kw2)
        if idx1 > idx2:
            hint = f'SQL关键字顺序错误:"{kw1}"出现在"{kw2}"之前,但标准顺序要求"{kw1}"在"{kw2}"之后.正确顺序: {" -> ".join(order_keywords)}'
            return hint
    
    # === 缺少关键字 ===
    # 有GROUP BY但没有聚合函数
    if 'GROUP BY' in sql_upper:
        agg_funcs = ['COUNT(', 'SUM(', 'AVG(', 'MIN(', 'MAX(', 'COUNT (', 'SUM (', 'AVG (', 'MIN (', 'MAX (']
        has_agg = any(af in sql_upper for af in agg_funcs)
        if not has_agg:
            hint = 'GROUP BY通常与聚合函数一起使用(如COUNT/SUM/AVG/MIN/MAX).如果只是去重,请用SELECT DISTINCT.'
            return hint
    
    # 有JOIN但缺少ON
    if re.search(r'\bJOIN\b', sql_upper) and ' ON ' not in sql_upper and not re.search(r'\bCROSS\s+JOIN\b', sql_upper):
        hint = 'JOIN缺少ON条件.例如:... JOIN 表2 ON 表1.id = 表2.id.如果是笛卡尔积,请用CROSS JOIN.'
        return hint
    
    # UPDATE语句出现在SELECT查询中
    if 'UPDATE' in sql_upper and 'SET' in sql_upper and 'SELECT' in sql_upper:
        hint = '不能在SELECT查询中使用UPDATE.批量修改请使用excel_update_query工具.'
        return hint
    
    # === 缺少逗号检测 ===
    # SELECT a b FROM -> SELECT a, b FROM(两个标识符之间只有空格没有逗号)
    select_match = re.search(r'\bSELECT\s+(.+?)\bFROM\b', sql_upper, re.DOTALL)
    if select_match:
        select_raw = sql[select_match.start(1):select_match.end(1)]
        # 检查原始SQL中两个标识符之间是否缺少逗号
        # 模式:单词 + 空格(非逗号) + 单词,其中两个都不是SQL关键字
        keywords_in_select = {'AS', 'DISTINCT', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'AND', 'OR', 'NOT', 'IN', 'BETWEEN', 'LIKE', 'IS', 'NULL', 'TRUE', 'FALSE', 'COUNT', 'SUM', 'AVG', 'MIN', 'MAX', 'UPPER', 'LOWER', 'TRIM', 'LENGTH', 'CONCAT', 'REPLACE', 'SUBSTRING', 'LEFT', 'RIGHT', 'COALESCE', 'IFNULL', 'CAST', 'ROW_NUMBER', 'RANK', 'DENSE_RANK', 'OVER', 'PARTITION', 'ASC', 'DESC', 'ON'}
        # 匹配:标识符 + 空格 + 标识符(中间无逗号)
        adjacent_pairs = re.finditer(r'([A-Za-z_]\w*)\s+([A-Za-z_]\w*)', select_raw)
        for m in adjacent_pairs:
            t1, t2 = m.group(1), m.group(2)
            if t1.upper() not in keywords_in_select and t2.upper() not in keywords_in_select:
                # 检查它们之间没有逗号(finditer已经保证了没有逗号,因为逗号不是\w)
                hint = f'SELECT子句中"{t1}"和"{t2}"之间可能缺少逗号.列之间用逗号分隔:SELECT {t1}, {t2}'
                return hint
    
    # === 括号配对检测 ===
    paren_count = sql.count('(') - sql.count(')')
    if paren_count > 0:
        hint = f'SQL中有{paren_count}个未闭合的括号.请检查每个左括号"("都有对应的右括号")".'
        return hint
    if paren_count < 0:
        hint = f'SQL中有多余的{abs(paren_count)}个右括号")".请删除多余的括号.'
        return hint
    
    # === 引号配对检测 ===
    single_quotes = len(re.findall(r"(?<!')'(?!')", sql))
    if single_quotes % 2 != 0:
        hint = 'SQL中的单引号数量为奇数,可能有未闭合的引号.字符串值需要用单引号包裹,如 \'值\'.'
        return hint
    
    # === 中文标点混用 ===
    cn_punctuation = {',': ',', '(': '(', ')': ')', ':': ':', ';': ';'}
    for cn, en in cn_punctuation.items():
        if cn in sql:
            hint = f'SQL中使用了中文标点"{cn}",应改为英文标点"{en}".'
            return hint
    
    # === Excel函数名误用 ===
    excel_funcs = {
        'SUMIF': '请用 CASE WHEN ... THEN ... END 替代 SUMIF',
        'COUNTIF': '请用 COUNT(CASE WHEN ... THEN 1 END) 替代 COUNTIF',
        'VLOOKUP': '请用 JOIN 替代 VLOOKUP',
        'IF': '请用 CASE WHEN ... THEN ... ELSE ... END 替代 IF 函数',
        'IFS': '请用 CASE WHEN ... THEN ... ELSE ... END 替代 IFS',
    }
    for func, suggestion in excel_funcs.items():
        if re.search(r'\b' + func + r'\s*\(', sql_upper):
            hint = f'Excel函数"{func}"不是SQL语法.{suggestion}.'
            return hint
    
    # === 子查询缺少别名 ===
    subquery_pattern = re.search(r'\(\s*SELECT\b.+?\)\s*$', sql.strip(), re.IGNORECASE | re.DOTALL)
    if subquery_pattern:
        end_part = sql.strip()[subquery_pattern.end():].strip()
        # 如果子查询后没有别名(没有内容,或内容不是 AS/标识符)
        if not end_part or (not re.match(r'^AS\b', end_part, re.IGNORECASE) and not re.match(r'^[A-Za-z_]\w*$', end_part)):
            hint = 'FROM子查询或UNION结果需要别名.例如:FROM (SELECT ...) AS subquery'
            return hint
    
    # === 通用建议 ===
    if 'SUBSTRING' in sql_upper and '(' in sql:
        # 检查SUBSTRING参数是否正确
        substr_match = re.search(r'SUBSTRING\s*\((.+?)\)', sql, re.IGNORECASE)
        if substr_match:
            args = [a.strip() for a in substr_match.group(1).split(',')]
            if len(args) == 2:
                hint = 'SUBSTRING需要3个参数:SUBSTRING(列, 起始位置, 长度).如果要从位置N取到末尾,请用SUBSTRING(列, N, LENGTH(列)-N+1).'
                return hint
    
    return hint


def _classify_value_error(err_str: str) -> str:
    """将ValueError分类为标准错误码.优先匹配更具体的模式."""
    err_upper = err_str.upper()
    # 更具体的模式优先匹配
    if "列 '" in err_str or "COLUMN" in err_upper:
        return "column_not_found"
    if "表 '" in err_str or "TABLE" in err_upper:
        return "table_not_found"
    if "窗口函数" in err_str or "WINDOW" in err_upper:
        return "window_function_error"
    if "JOIN" in err_upper:
        return "join_error"
    if "子查询" in err_str or "SUBQUERY" in err_upper:
        return "subquery_error"
    if "UNION" in err_upper:
        return "union_error"
    if "CTE" in err_upper or "WITH" in err_upper:
        return "cte_error"
    if "GROUP BY" in err_upper:
        return "group_by_error"
    if "ORDER BY" in err_upper:
        return "order_by_error"
    # 通用模式放后面
    if "不支持" in err_str or "UNSUPPORTED" in err_upper:
        return "unsupported_feature"
    if "表达式" in err_str:
        return "expression_error"
    return "execution_error"


def _generate_value_error_hint(err_str: str) -> str:
    """根据ValueError内容生成AI修复建议."""
    # 列不存在
    if "列 '" in err_str and "可用列" in err_str:
        return "请检查列名拼写,或先用excel_get_headers查看可用列名."
    # 表不存在
    if "表 '" in err_str and "可用表" in err_str:
        return "请检查表名拼写,或先用excel_list_sheets查看可用工作表."
    # FROM子查询
    if "FROM子查询" in err_str:
        return "请检查FROM子查询中的SQL语法和表名.FROM子查询需要别名:FROM (SELECT ...) AS alias."
    # JOIN相关
    if "JOIN表" in err_str and "不存在" in err_str:
        return "请检查JOIN的表名是否正确,先用excel_list_sheets确认可用工作表."
    if "JOIN缺少ON条件" in err_str:
        return "JOIN必须包含ON条件,例如:... JOIN 表2 ON 表1.id = 表2.id."
    if "没有列 '" in err_str:
        return "请检查ON条件中的列名,确认列属于哪个表."
    # 窗口函数
    if "不支持的窗口函数" in err_str:
        return "仅支持 ROW_NUMBER,RANK,DENSE_RANK 窗口函数."
    if "需要 ORDER BY" in err_str:
        return "该窗口函数必须包含 ORDER BY 子句."
    # UNION
    if "UNION" in err_str and "SELECT" in err_str:
        return "请确保UNION两侧的SELECT列数一致."
    # 数学表达式
    if "数学表达式" in err_str:
        return "请检查数学运算符和操作数是否正确."
    # 字符串函数
    if "字符串函数" in err_str:
        return "请检查函数名和参数.支持的字符串函数:UPPER/LOWER/TRIM/LENGTH/CONCAT/REPLACE/SUBSTRING/LEFT/RIGHT."
    return ""


def _generate_value_error_suggested_fix(err_str: str, sql: str) -> str:
    """根据ValueError内容生成具体的修复SQL建议(如果可能)."""
    # 列不存在:尝试提取建议的列名并替换
    if "列 '" in err_str and "你是否想用" in err_str:
        import re
        # 提取建议的列名
        suggestion_match = re.search(r'你是否想用:\s*(.+?)\?', err_str)
        if suggestion_match:
            suggested_col = suggestion_match.group(1).strip().split(',')[0].strip()
            # 提取错误的列名
            col_match = re.search(r"列 '(.+?)'", err_str)
            if col_match:
                wrong_col = col_match.group(1)
                return sql.replace(wrong_col, suggested_col)
    # 表不存在:尝试提取建议的表名
    if "表 '" in err_str and "你是否想用" in err_str:
        import re
        suggestion_match = re.search(r'你是否想用:\s*(.+?)\?', err_str)
        if suggestion_match:
            suggested_table = suggestion_match.group(1).strip().split(',')[0].strip()
            col_match = re.search(r"表 '(.+?)'", err_str)
            if col_match:
                wrong_table = col_match.group(1)
                return sql.replace(wrong_table, suggested_table)
    return ""


def _safe_float_comparison(left, right, op):
    """安全比较函数,处理None值,避免 '<=' not supported between instances of 'int' and 'NoneType' 错误
    
    REQ-034优化:
    - 不等式比较(>/>=/</<=)保持精确,不使用epsilon
    - 仅 == 比较使用自适应epsilon处理浮点精度问题(如0.1+0.2approx0.3)
    - 支持极端边界值:0.001秒精度,99999大数值,NULL处理
    """
    if left is None or right is None:
        return False
    
    try:
        left_float = float(left)
        right_float = float(right)
        
        if op == '>':
            return left_float > right_float
        elif op == '>=':
            return left_float >= right_float
        elif op == '<':
            return left_float < right_float
        elif op == '<=':
            return left_float <= right_float
        elif op == '==' or op == '=':
            # 仅等值比较使用epsilon
            max_val = max(abs(left_float), abs(right_float))
            epsilon = max(max_val * 1e-9, 1e-10)
            return abs(left_float - right_float) <= epsilon
        else:
            return False
    except (TypeError, ValueError):
        return False


class AdvancedSQLQueryEngine:
    """高级SQL查询引擎,支持完整的SQL语法"""

    def __init__(self, disable_streaming_aggregate: bool = False):
        """
        初始化SQL查询引擎

        Args:
            disable_streaming_aggregate: 禁用流式聚合优化(大文件处理)
        """
        self.disable_streaming_aggregate = disable_streaming_aggregate
        # DataFrame缓存:{file_path: (mtime, worksheets_data, header_descriptions)}
        self._df_cache = {}
        self._max_cache_size = MAX_CACHE_SIZE  # 最大缓存文件数,防止内存泄漏
        # 列名映射缓存:{file_path: {原始列名: 清洗列名}}
        # 与_df_cache同步,避免缓存命中时_original_to_clean_cols为空
        self._col_map_cache = {}

        # 性能优化:查询结果缓存 {hash(sql): (result_df, file_mtime)}
        self._query_result_cache = {}
        self._max_query_cache_size = MAX_QUERY_CACHE_SIZE  # 最大查询缓存数,防止内存泄漏
        self._query_cache_ttl = QUERY_CACHE_TTL  # 查询缓存TTL

        if not SQLGLOT_AVAILABLE:
            raise ImportError("SQLGlot未安装,请运行: pip install sqlglot")

    def clear_cache(self):
        """清除DataFrame缓存,释放内存
        
        清除DataFrame缓存和查询结果缓存,释放内存占用.
        下次查询会重新加载Excel数据.
        
        Args:
            无
        
        Returns:
            None
        """
        self._df_cache.clear()
        self._query_result_cache.clear()

    def _get_query_cache_key(self, sql: str, file_path: str, sheet_name: Optional[str] = None) -> str:
        """生成查询缓存键"""
        import hashlib
        cache_data = f"{sql}|{file_path}|{sheet_name or ''}"
        return hashlib.md5(cache_data.encode()).hexdigest()

    def _get_cached_query_result(self, cache_key: str, file_mtime: float) -> Optional[pd.DataFrame]:
        """获取缓存的查询结果"""
        if cache_key in self._query_result_cache:
            cached_time, cached_df, cached_mtime = self._query_result_cache[cache_key]
            # 检查缓存是否过期(文件是否被修改)
            current_time = time.time()
            if current_time - cached_time < self._query_cache_ttl and cached_mtime == file_mtime:
                return cached_df
            else:
                # 缓存过期,删除
                del self._query_result_cache[cache_key]
        return None

    def _cache_query_result(self, cache_key: str, result_df: pd.DataFrame, file_mtime: float):
        """缓存查询结果"""
        # LRU淘汰:超过最大缓存数时删除最早的缓存
        if len(self._query_result_cache) >= self._max_query_cache_size:
            oldest_key = next(iter(self._query_result_cache))
            del self._query_result_cache[oldest_key]
        
        self._query_result_cache[cache_key] = (time.time(), result_df, file_mtime)

    def execute_sql_query(
        self,
        file_path: str,
        sql: str,
        sheet_name: Optional[str] = None,
        limit: Optional[int] = None,
        include_headers: bool = True,
        output_format: str = "table"
    ) -> Dict[str, Any]:
        """

        Args:
            file_path: Excel文件路径
            sql: SQL查询语句，支持完整的SELECT语法包括WHERE、JOIN、GROUP BY、HAVING、ORDER BY等
            sheet_name: 工作表名称（可选，默认使用第一个）
            limit: 限制返回行数，用于控制结果集大小
            include_headers: 是否包含表头，True时返回数据第一行为列名
            output_format: 输出格式，支持 table/json/csv 三种格式

        Returns:
            Dict: 查询结果，包含以下字段:
                - success (bool): 查询是否成功
                - message (str): 结果消息或错误描述
                - data (list): 查询结果数据（二维数组）
                - query_info (dict): 查询详细信息
                    - original_rows: 原始数据总行数
                    - filtered_rows: 过滤后行数
                    - returned_rows: 实际返回行数（可能因截断减少）
                    - truncated: 是否被截断（超过最大行数限制）
                    - query_applied: 是否应用了查询条件
                    - sql_query: 执行的SQL语句
                    - columns_returned: 返回的列数
                    - available_tables: 可用的工作表列表
                    - returned_columns: 返回的列名列表
                    - data_types: 各列的数据类型推断
                    - execution_time_ms: 查询执行耗时（毫秒）
                    - markdown_table: Markdown格式表格（如数据存在）
                    - suggestion: 空结果时的智能建议（如结果为空）
                    - json_output: JSON格式输出（output_format=json时）
                    - csv_output: CSV格式输出（output_format=csv时）
        """
        try:
            # 验证文件存在性
            if not os.path.exists(file_path):
                return {
                    'success': False,
                    'message': f'文件不存在: {file_path}',
                    'data': [],
                    'query_info': {'error_type': 'file_not_found'}
                }

            # 检查文件大小并处理大文件(支持2GB+文件)
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > 2048:
                return {
                    'success': False,
                    'message': f'文件过大 ({file_size_mb:.2f}MB),建议使用小于2GB的文件',
                    'data': [],
                    'query_info': {'error_type': 'file_too_large', 'size_mb': file_size_mb}
                }
            if file_size_mb > 500:
                logger.info(f"大文件查询: {file_path} ({file_size_mb:.1f}MB),启用分块处理优化")

            file_mtime = os.path.getmtime(file_path)
            
            # 加载Excel数据(带缓存)
            # 重置列名映射(每次查询重新构建)
            self._original_to_clean_cols = {}
            worksheets_data = self._load_data_with_cache(file_path, sheet_name)

            if not worksheets_data:
                return {
                    'success': False,
                    'message': '无法加载Excel数据或文件为空',
                    'data': [],
                    'query_info': {'error_type': 'data_load_failed'}
                }

            # 跨文件引用解析:FROM 表名@'path' 语法
            # 在sqlglot解析前处理,加载外部文件并合并worksheets_data
            if "@'" in sql or '@"' in sql:
                sql, worksheets_data = self._resolve_cross_file_references(
                    sql, file_path, worksheets_data
                )

            # 清理ANSI转义序列(终端粘贴可能带入的不可见字符)
            sql = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', sql)
            # 清理残余控制字符
            sql = re.sub(r'[\x00-\x1f\x7f]', '', sql)
            # 清理残余的ANSI括号伪影:未配对的[后紧跟非ASCII字符
            # 有效的SQL Server标识符 [名称] 有配对的],ANSI伪影 [中文 无配对
            if sql.count('[') != sql.count(']'):
                # 存在未配对括号,清理[后紧跟非ASCII字符的情况
                sql = re.sub(r'\[(?=[^\x00-\x7F])', '', sql)

            # 中文列名替换:将SQL中的中文列名替换为英文列名(在解析前)
            sql = self._replace_cn_columns_in_sql(sql, worksheets_data)

            # DESCRIBE命令友好提示
            sql_stripped = sql.strip().upper()
            if sql_stripped.startswith('DESCRIBE ') or sql_stripped.startswith('DESC '):
                table_hint = sql.strip().split(None, 1)[-1].strip(';').strip('"\'`') if len(sql.strip().split()) > 1 else ''
                hint = f'请使用 excel_describe_table 工具查看表结构'
                if table_hint:
                    hint += f'(工作表: {table_hint})'
                return {
                    'success': False,
                    'message': f'DESCRIBE不是SQL查询语法.{hint}',
                    'data': [],
                    'query_info': {'error_type': 'describe_not_sql', 'hint': 'use_excel_describe_table'}
                }

            # 解析和执行SQL
            _query_start = time.time()
            try:
                # 预处理:将双引号引用的原始列名替换为清洗后的列名
                # 解决用户写 SELECT "Player Name" 但内部列名已变为 Player_Name 的问题
                sql = self._preprocess_quoted_identifiers(sql)

                parsed_sql = sqlglot.parse_one(sql, dialect="mysql")

                # 验证SQL支持范围
                validation_result = self._validate_sql_support(parsed_sql)
                if not validation_result['valid']:
                    error_msg = validation_result.get('error', '不支持的SQL语法')
                    hint = _unsupported_error_hint(error_msg)
                    qi = {'error_type': 'unsupported_sql', 'details': validation_result}
                    if hint:
                        qi['hint'] = hint
                    return {
                        'success': False,
                        'message': f'不支持的SQL语法: {error_msg}' + (f'\n💡 {hint}' if hint else ''),
                        'data': [],
                        'query_info': qi
                    }

                # 执行查询(UNION/UNION ALL/EXCEPT/INTERSECT 或普通 SELECT)
                if isinstance(parsed_sql, exp.Union):
                    result_data = self._execute_union(parsed_sql, worksheets_data, limit)
                elif isinstance(parsed_sql, (exp.Except, exp.Intersect)):
                    result_data = self._execute_except_intersect(parsed_sql, worksheets_data, limit)
                else:
                    result_data = self._execute_query(parsed_sql, worksheets_data, limit)
                _query_elapsed = (time.time() - _query_start) * 1000

                # 格式化结果(传入parsed_sql和WHERE前数据用于空结果智能建议)
                has_group_by = not isinstance(parsed_sql, (exp.Union, exp.Except, exp.Intersect)) and parsed_sql.args.get('group') is not None
                has_having = parsed_sql.args.get('having') is not None
                result = self._format_query_result(
                    result_data,
                    file_path,
                    sql,
                    worksheets_data,
                    include_headers,
                    has_group_by=has_group_by,
                    has_having=has_having,
                    parsed_sql=parsed_sql,
                    df_before_where=self._df_before_where,
                    output_format=output_format
                )
                # 注入执行时间
                result['query_info']['execution_time_ms'] = round(_query_elapsed, 1)
                
                return result

            except StructuredSQLError as e:
                qi = {
                    'error_type': e.error_code,
                    'hint': e.hint,
                    'context': e.context,
                    'details': e.message
                }
                # 为列名/表名错误生成suggested_fix
                suggested_fix = ""
                if e.error_code in ('column_not_found', 'table_not_found') and e.context:
                    wrong_name = e.context.get('column_requested') or e.context.get('table_requested', '')
                    available = e.context.get('available_columns') or e.context.get('available_tables') or []
                    if wrong_name and available:
                        matches = difflib.get_close_matches(wrong_name, available, n=1, cutoff=0.4)
                        if matches:
                            suggested_fix = sql.replace(wrong_name, matches[0], 1)
                if suggested_fix:
                    qi['suggested_fix'] = suggested_fix
                msg = e.message
                if e.hint:
                    msg += f'\n💡 {e.hint}'
                if suggested_fix:
                    msg += f'\n🔧 建议修复SQL: {suggested_fix}'
                return {
                    'success': False,
                    'message': msg,
                    'data': [],
                    'query_info': qi
                }
            except ParseError as e:
                err_str = str(e)
                hint = _parse_error_hint(err_str, sql)
                qi = {
                    'error_type': 'syntax_error',
                    'details': err_str,
                    'hint': hint
                }
                return {
                    'success': False,
                    'message': f'SQL语法错误: {err_str}' + (f'\n💡 {hint}' if hint else ''),
                    'data': [],
                    'query_info': qi
                }
            except UnsupportedError as e:
                err_detail = str(e)
                # 为不支持的SQL功能提供替代建议
                hint = _unsupported_error_hint(err_detail)
                qi = {
                    'error_type': 'unsupported_feature',
                    'details': err_detail,
                    'hint': hint
                }
                return {
                    'success': False,
                    'message': f'不支持的SQL功能: {err_detail}' + (f'\n💡 {hint}' if hint else ''),
                    'data': [],
                    'query_info': qi
                }
            except ValueError as e:
                err_str = str(e)
                # 对常见ValueError生成智能修复建议
                hint = _generate_value_error_hint(err_str)
                error_code = _classify_value_error(err_str)
                suggested_fix = _generate_value_error_suggested_fix(err_str, sql)
                qi = {
                    'error_type': error_code,
                    'details': err_str,
                    'hint': hint
                }
                if suggested_fix:
                    qi['suggested_fix'] = suggested_fix
                msg = err_str
                if hint:
                    msg += f'\n💡 {hint}'
                if suggested_fix:
                    msg += f'\n🔧 建议修复SQL: {suggested_fix}'
                return {
                    'success': False,
                    'message': msg,
                    'data': [],
                    'query_info': qi
                }
            except Exception as e:
                return {
                    'success': False,
                    'message': f'SQL执行错误: {str(e)}',
                    'data': [],
                    'query_info': {'error_type': 'execution_error', 'details': str(e)}
                }

        except Exception as e:
            return {
                'success': False,
                'message': f'查询引擎错误: {str(e)}',
                'data': [],
                'query_info': {'error_type': 'engine_error', 'details': str(e)}
            }

    def _resolve_cross_file_references(
        self, sql: str, primary_file_path: str, primary_worksheets: Dict[str, pd.DataFrame]
    ) -> Tuple[str, Dict[str, pd.DataFrame]]:
        """
        解析SQL中的跨文件引用(@'path'语法),加载外部文件数据并合并到worksheets_data

        语法:FROM 技能表@'/path/to/file1.xlsx' s JOIN 掉落表@'/path/to/file2.xlsx' d ON ...
        支持单引号和双引号包裹路径,支持相对路径(相对于主文件目录)

        Args:
            sql: 原始SQL语句
            primary_file_path: 主文件路径
            primary_worksheets: 主文件的worksheets_data

        Returns:
            Tuple[str, Dict]: (清理后的SQL, 合并后的worksheets_data)
        """
        # 正则匹配 @'path' 或 @"path" 模式
        # 路径可以包含字母,数字,./,../,_,-,空格等
        cross_file_pattern = re.compile(
            r"""@(['"])(.*?)\1""",
            re.DOTALL
        )

        matches = list(cross_file_pattern.finditer(sql))
        if not matches:
            return sql, primary_worksheets

        merged_data = dict(primary_worksheets)
        cleaned_sql = sql
        primary_dir = os.path.dirname(os.path.abspath(primary_file_path))
        loaded_files = {}  # filepath -> worksheets_data(避免重复加载)

        # 从后向前替换,避免索引偏移
        for match in reversed(matches):
            quote_char = match.group(1)
            ref_path = match.group(2).strip()

            # 解析相对路径(相对于主文件目录)
            if not os.path.isabs(ref_path):
                ref_path = os.path.join(primary_dir, ref_path)

            ref_path = os.path.normpath(ref_path)

            # 验证文件存在
            if not os.path.exists(ref_path):
                raise ValueError(
                    f"跨文件引用的文件不存在: {ref_path}."
                    f"请检查路径是否正确(支持绝对路径和相对于主文件的相对路径)"
                )

            # 加载文件(带缓存,避免重复加载)
            if ref_path not in loaded_files:
                ext_worksheets = self._load_data_with_cache(ref_path)
                if not ext_worksheets:
                    raise ValueError(f"无法加载跨文件引用的Excel数据: {ref_path}")
                loaded_files[ref_path] = ext_worksheets

            ext_worksheets = loaded_files[ref_path]

            # 合并工作表数据,处理名称冲突
            # 冲突时:主文件优先(已存在的不覆盖),外部文件重命名添加文件前缀
            file_basename = os.path.splitext(os.path.basename(ref_path))[0]
            for sheet_name, df in ext_worksheets.items():
                if sheet_name in merged_data:
                    # 名称冲突:为外部文件的工作表添加文件前缀
                    prefixed_name = f"{file_basename}.{sheet_name}"
                    if prefixed_name in merged_data:
                        # 即使加前缀也冲突,添加数字后缀
                        counter = 2
                        while f"{prefixed_name}_{counter}" in merged_data:
                            counter += 1
                        prefixed_name = f"{prefixed_name}_{counter}"
                    merged_data[prefixed_name] = df
                else:
                    merged_data[sheet_name] = df

            # 从SQL中移除 @'path' 部分,保留表名和别名
            cleaned_sql = cleaned_sql[:match.start()] + cleaned_sql[match.end():]

        return cleaned_sql, merged_data

    def _load_data_with_cache(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[Dict[str, pd.DataFrame]]:
        """
        带缓存的Excel数据加载(公共方法,供execute_sql_query和execute_update_query复用)

        使用mtime检测文件变更,LRU淘汰防止内存泄漏.
        大文件(>500MB)使用分块读取减少内存峰值.

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称(可选)

        Returns:
            worksheets_data字典,加载失败返回None
        """
        mtime = os.path.getmtime(file_path)
        cache_key = f"{file_path}|{sheet_name or ''}"
        if cache_key in self._df_cache:
            cached_mtime, cached_data, cached_desc = self._df_cache[cache_key]
            if cached_mtime == mtime:
                self._header_descriptions = cached_desc
                # 缓存命中时,重置列名映射为当前文件的正确映射,避免其他文件的映射干扰
                self._original_to_clean_cols = {}
                if cache_key in self._col_map_cache:
                    self._original_to_clean_cols.update(self._col_map_cache[cache_key])
                return cached_data
            else:
                # 文件已修改,重新加载
                worksheets_data = self._load_excel_data(file_path, sheet_name)
                self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
                # 保存列名映射到缓存
                self._col_map_cache[cache_key] = dict(self._original_to_clean_cols)
                return worksheets_data
        else:
            worksheets_data = self._load_excel_data(file_path, sheet_name)
            self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
            # 保存列名映射到缓存
            self._col_map_cache[cache_key] = dict(self._original_to_clean_cols)
            # LRU淘汰:超过最大缓存数时删除最早缓存的文件
            while len(self._df_cache) > self._max_cache_size:
                evicted_key = next(iter(self._df_cache))
                self._df_cache.pop(evicted_key)
                self._col_map_cache.pop(evicted_key, None)
            return worksheets_data

    def _estimate_cache_memory_mb(self) -> float:
        """估算当前缓存占用的内存(MB)"""
        total = 0.0
        for _key, (_mtime, worksheets_data, _desc) in self._df_cache.items():
            for _sheet, df in worksheets_data.items():
                total += df.memory_usage(deep=True).sum() / 1024 / 1024
        return total

    def evict_cache_by_memory(self, target_mb: float = None):
        """内存感知缓存淘汰:当缓存总内存超过阈值时,淘汰最早的缓存项

        Args:
            target_mb: 目标最大缓存内存(MB),默认使用 CACHE_TARGET_MEMORY_MB
        """
        if target_mb is None:
            target_mb = CACHE_TARGET_MEMORY_MB
        while self._estimate_cache_memory_mb() > target_mb and self._df_cache:
            self._df_cache.pop(next(iter(self._df_cache)))
            logger.info(f"缓存内存淘汰后剩余: {self._estimate_cache_memory_mb():.1f}MB")

    def _load_excel_data(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, pd.DataFrame]:
        """
        加载Excel数据到DataFrame字典,支持游戏配置表双行表头

        游戏配置表通常有双行表头:
          第1行:中文描述(如"技能ID","技能名称")
          第2行:字段名(如"skill_id","skill_name")
        
        本方法自动检测双行表头,用第二行(字段名)做列名,
        第一行(描述)保存在 self._header_descriptions 中.

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称(可选)

        Returns:
            Dict[str, pd.DataFrame]: 工作表名到DataFrame的映射
        """
        worksheets_data = {}
        self._header_descriptions = {}  # {sheet_name: {field_name: description}}

        try:
            # 性能优化:用calamine替代openpyxl读取(Rust引擎,速度提升10-50倍)
            # calamine一次性读取所有sheet数据,无需二次打开文件
            from python_calamine import CalamineWorkbook

            cal_wb = CalamineWorkbook.from_path(file_path)
            all_sheet_names = cal_wb.sheet_names

            if sheet_name:
                sheets_to_load = [sheet_name] if sheet_name in all_sheet_names else []
            else:
                sheets_to_load = all_sheet_names

            # 批量检测所有sheet的双行表头(calamine读取前两行,毫秒级)
            header_info = {}  # {sheet: (is_dual_header, first_row_values, second_row_values)}
            for sheet in sheets_to_load:
                try:
                    cal_ws = cal_wb.get_sheet_by_name(sheet)
                    # 跳过空工作表(无数据行)
                    if cal_ws.height == 0:
                        header_info[sheet] = (False, None, None)
                        continue
                    rows_iter = cal_ws.iter_rows()
                    first_row = next(rows_iter, None)
                    second_row = next(rows_iter, None)

                    is_dual_header = False
                    if first_row and second_row:
                        second_row_values = [str(v).strip() if v is not None else '' for v in second_row]
                        first_row_values = [str(v).strip() if v is not None else '' for v in first_row]

                        non_empty_second = [v for v in second_row_values if v]
                        non_empty_first = [v for v in first_row_values if v]

                        if len(non_empty_second) >= 2:
                            second_all_field = all(
                                re.match(r'^[a-zA-Z_][a-zA-Z0-9_.#]*$', v)
                                for v in non_empty_second
                            )
                            first_all_field = all(
                                re.match(r'^[a-zA-Z_][a-zA-Z0-9_#]*$', v)
                                for v in non_empty_first
                            ) if non_empty_first else False
                            if second_all_field and not first_all_field:
                                is_dual_header = True

                    header_info[sheet] = (is_dual_header, first_row_values, second_row_values)
                except Exception:
                    header_info[sheet] = (False, None, None)

            # 批量读取所有sheet数据(pd.read_excel + calamine引擎)
            for sheet, (is_dual_header, first_row_values, second_row_values) in header_info.items():
                if is_dual_header:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='calamine',
                        header=1,
                        keep_default_na=False
                    )
                    # 注意:desc_map 在 _clean_dataframe 之后构建(列名可能被清洗)
                    # 先记录原始映射关系,后面清洗后再构建最终映射
                    raw_desc_pairs = []
                    if second_row_values and first_row_values:
                        for col_idx, fname in enumerate(second_row_values):
                            fname = fname.strip() if fname else ''
                            desc = first_row_values[col_idx].strip() if col_idx < len(first_row_values) else ''
                            if fname and desc and desc != fname:
                                raw_desc_pairs.append((col_idx, fname, desc))

                else:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='calamine',
                        keep_default_na=False
                    )
                    raw_desc_pairs = []

                # 清洗 DataFrame(列名中的特殊字符会被替换)
                original_columns = list(df.columns)
                df = self._clean_dataframe(df)
                cleaned_columns = list(df.columns)

                # 构建中文描述映射(用清洗后的列名)
                if raw_desc_pairs:
                    desc_map = {}
                    for col_idx, fname, desc in raw_desc_pairs:
                        if col_idx < len(cleaned_columns):
                            desc_map[cleaned_columns[col_idx]] = desc
                    self._header_descriptions[sheet] = desc_map

                df = self._optimize_dtypes(df)
                worksheets_data[sheet] = df

        except Exception as e:
            logger.error(f"加载Excel数据失败: {e}")
            return {}

        return worksheets_data

    def _clean_dataframe(self, df) -> pd.DataFrame:
        """
        清理DataFrame数据

        Args:
            df: 原始DataFrame

        Returns:
            pd.DataFrame: 清理后的DataFrame
        """
        # 删除完全为空的行和列
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')

        # 重置索引
        df = df.reset_index(drop=True)

        # 清理列名
        clean_columns = {}
        for col in df.columns:
            clean_col = str(col).strip()
            # 处理Unicode编码
            if '\\u' in clean_col:
                try:
                    clean_col = clean_col.encode('raw_unicode_escape').decode('unicode_escape')
                except Exception:
                    pass

            # 清理特殊字符,但保持中文
            clean_col = re.sub(r'[^\w\u4e00-\u9fff\s]', '_', clean_col)
            clean_col = re.sub(r'\s+', '_', clean_col)

            # 确保列名不为空且不以数字开头
            if not clean_col or clean_col.isspace():
                clean_col = f"column_{len(clean_columns) + 1}"
            elif clean_col[0].isdigit():
                clean_col = f"col_{clean_col}"

            clean_columns[col] = clean_col

        df = df.rename(columns=clean_columns)

        # 保存原始列名到清洗后列名的映射,用于SQL预处理
        if not hasattr(self, '_original_to_clean_cols') or self._original_to_clean_cols is None:
            self._original_to_clean_cols = {}
        self._original_to_clean_cols.update(clean_columns)

        # 保持原始数据不做空值替换
        # pandas groupby 默认跳过 NaN 行,不需要手动处理

        return df

    def _preprocess_quoted_identifiers(self, sql: str) -> str:
        """
        预处理SQL中的双引号标识符,将原始列名替换为清洗后的列名

        当Excel列名含空格或特殊字符时(如"Player Name"),_clean_column_names会将其
        转换为下划线形式(Player_Name).用户在SQL中使用双引号引用原始列名时,
        sqlglot(MySQL方言)会将其解析为字符串字面量而非列引用.

        使用AST方法精确替换:只在列引用位置(SELECT/ORDER BY/GROUP BY/HAVING,
        WHERE比较左侧)替换,保留WHERE值位置的字符串字面量不变.
        例如:SELECT "Player Name" FROM t WHERE type = "Player Name"
        -> SELECT `Player_Name` FROM t WHERE type = 'Player Name'
        (SELECT中替换为列引用,WHERE值位置保留为字符串)

        Args:
            sql: 原始SQL查询语句

        Returns:
            str: 预处理后的SQL语句
        """
        col_map = getattr(self, '_original_to_clean_cols', None)
        if not col_map:
            return sql

        # 只处理原始名与清洗名不同的列(即含空格或特殊字符的列名)
        changed_cols = {orig: clean for orig, clean in col_map.items() if orig != clean}
        if not changed_cols:
            return sql

        try:
            import sqlglot
            from sqlglot import exp as sg_exp

            parsed = sqlglot.parse_one(sql, dialect="mysql")
            if parsed is None:
                return self._fallback_preprocess(sql, changed_cols)

            # 在SELECT表达式中替换:双引号字符串 -> 列引用
            select = parsed.find(sg_exp.Select)
            if select:
                new_exprs = []
                for e in select.expressions:
                    new_exprs.append(self._literal_to_column(e, changed_cols))
                select.set("expressions", new_exprs)

            # 在ORDER BY中替换
            order = parsed.find(sg_exp.Order)
            if order:
                new_ordered = []
                for o in order.expressions:
                    new_ordered.append(self._literal_to_column(o, changed_cols))
                order.set("expressions", new_ordered)

            # 在GROUP BY中替换
            group = parsed.find(sg_exp.Group)
            if group:
                new_group = []
                for g in group.expressions:
                    new_group.append(self._literal_to_column(g, changed_cols))
                group.set("expressions", new_group)

            # 在HAVING中替换
            having = parsed.find(sg_exp.Having)
            if having:
                self._replace_having_literals(having.this, changed_cols)

            # 在WHERE子句中:只替换比较操作左侧的字面量(列引用位置)
            where = parsed.find(sg_exp.Where)
            if where:
                self._replace_where_left_literals(where.this, changed_cols)

            result_sql = parsed.sql(dialect="mysql")
            return result_sql

        except Exception:
            # AST解析失败,回退到简单替换(仅SELECT子句)
            return self._fallback_preprocess(sql, changed_cols)

    def _literal_to_column(self, node, changed_cols):
        """
        如果节点是匹配原始列名的字符串字面量,替换为列引用

        Args:
            node: sqlglot AST节点
            changed_cols: 原始列名到清洗列名的映射

        Returns:
            替换后的节点(列引用或原始节点)
        """
        from sqlglot import exp as sg_exp
        if isinstance(node, sg_exp.Literal) and node.is_string:
            lit_val = node.this
            if lit_val in changed_cols:
                return sg_exp.Column(
                    this=sg_exp.Identifier(this=changed_cols[lit_val])
                )
        return node

    def _replace_where_left_literals(self, node, changed_cols):
        """
        替换WHERE子句中比较操作左侧的字符串字面量为列引用

        只处理比较操作(=, !=, >, <, >=, <=)的左侧,保留右侧作为字符串值.
        递归处理AND/OR组合条件.

        Args:
            node: WHERE子句的AST节点
            changed_cols: 原始列名到清洗列名的映射
        """
        from sqlglot import exp as sg_exp
        comparison_types = (sg_exp.EQ, sg_exp.NEQ, sg_exp.GT, sg_exp.GTE, sg_exp.LT, sg_exp.LTE)

        if isinstance(node, comparison_types):
            if isinstance(node.this, sg_exp.Literal) and node.this.is_string:
                lit_val = node.this.this
                if lit_val in changed_cols:
                    node.set("this", sg_exp.Column(
                        this=sg_exp.Identifier(this=changed_cols[lit_val])
                    ))
        elif isinstance(node, (sg_exp.And, sg_exp.Or)):
            self._replace_where_left_literals(node.this, changed_cols)
            self._replace_where_left_literals(node.expression, changed_cols)
        elif isinstance(node, sg_exp.Paren):
            self._replace_where_left_literals(node.this, changed_cols)
        elif isinstance(node, sg_exp.Not):
            self._replace_where_left_literals(node.this, changed_cols)

    def _replace_having_literals(self, node, changed_cols):
        """
        替换HAVING子句中的字符串字面量为列引用

        Args:
            node: HAVING子句的AST节点
            changed_cols: 原始列名到清洗列名的映射
        """
        from sqlglot import exp as sg_exp
        comparison_types = (sg_exp.EQ, sg_exp.NEQ, sg_exp.GT, sg_exp.GTE, sg_exp.LT, sg_exp.LTE)

        if isinstance(node, comparison_types):
            # HAVING两侧都可能是列引用(HAVING "Player Name" > 100)
            node.this = self._literal_to_column(node.this, changed_cols)
        elif isinstance(node, (sg_exp.And, sg_exp.Or)):
            self._replace_having_literals(node.this, changed_cols)
            self._replace_having_literals(node.expression, changed_cols)

    def _fallback_preprocess(self, sql: str, changed_cols: Dict[str, str]) -> str:
        """
        回退预处理:仅替换SELECT子句中的双引号列名(使用正则提取SELECT子句)

        当AST解析失败时使用,避免全量字符串替换误伤WHERE字符串值.

        Args:
            sql: 原始SQL查询语句
            changed_cols: 原始列名到清洗列名的映射

        Returns:
            str: 预处理后的SQL语句
        """
        # 提取SELECT子句(SELECT ... FROM),仅在此范围内替换
        select_match = re.match(
            r'(SELECT\s+)(.*?)(\s+FROM\s+)',
            sql, re.IGNORECASE | re.DOTALL
        )
        if select_match:
            prefix = select_match.group(1)
            select_clause = select_match.group(2)
            from_suffix = sql[select_match.end(2):]

            for orig_name in sorted(changed_cols.keys(), key=len, reverse=True):
                clean_name = changed_cols[orig_name]
                select_clause = select_clause.replace(f'"{orig_name}"', f'`{clean_name}`')

            return prefix + select_clause + from_suffix

        # 无法提取SELECT子句,不替换(安全优先)
        return sql

    def _optimize_dtypes(self, df) -> pd.DataFrame:
        """
        优化DataFrame数据类型以减少内存占用

        对数值列进行降级(int64->int32/int16/int8, float64->float32),
        对高基数字符串列不做转换(避免转换开销),对低基数字符串列转为category.

        Args:
            df: 原始DataFrame

        Returns:
            pd.DataFrame: 类型优化后的DataFrame
        """
        start_mem = df.memory_usage(deep=True).sum() / 1024 / 1024

        for col in df.columns:
            col_type = df[col].dtype

            if col_type == 'object':
                # 字符串列:仅当低基数(唯一值/总行数 < 0.3)时转为 category
                # 高基数字符串列转 category 反而增加内存开销
                num_unique = df[col].nunique()
                if num_unique > 0 and num_unique / len(df) < 0.3:
                    df[col] = df[col].astype('category')
            elif col_type in ['int64', 'int32']:
                # 整数列降级
                col_min = df[col].min()
                col_max = df[col].max()
                if col_min >= 0:
                    if col_max < 256:
                        df[col] = df[col].astype('uint8')
                    elif col_max < 65536:
                        df[col] = df[col].astype('uint16')
                    elif col_max < 4294967296:
                        df[col] = df[col].astype('uint32')
                else:
                    if col_min > -128 and col_max < 127:
                        df[col] = df[col].astype('int8')
                    elif col_min > -32768 and col_max < 32767:
                        df[col] = df[col].astype('int16')
                    elif col_min > -2147483648 and col_max < 2147483647:
                        df[col] = df[col].astype('int32')
            elif col_type == 'float64':
                # 浮点列降级为 float32(精度足够)
                df[col] = df[col].astype('float32')

        end_mem = df.memory_usage(deep=True).sum() / 1024 / 1024
        reduction = (1 - end_mem / start_mem) * 100 if start_mem > 0 else 0
        logger.debug(f"dtype优化: {start_mem:.1f}MB -> {end_mem:.1f}MB (节省{reduction:.0f}%)")

        return df

    def _validate_sql_support(self, parsed_sql: exp.Expression) -> Dict[str, Any]:
        """
        验证SQL语法支持范围

        Args:
            parsed_sql: 解析后的SQL表达式

        Returns:
            Dict[str, Any]: 验证结果
        """
        try:
            # 检查是否为SELECT语句,UNION,EXCEPT或INTERSECT
            if not isinstance(parsed_sql, (exp.Select, exp.Union, exp.Except, exp.Intersect)):
                return {
                    'valid': False,
                    'error': '只支持SELECT查询语句,不支持INSERT,UPDATE,DELETE等操作'
                }

            # 子查询支持已实现,不再拒绝
            # IN子查询,标量子查询,EXISTS子查询均支持

            # CTE (WITH) 支持
            with_clause = parsed_sql.args.get('with') or parsed_sql.args.get('with_')
            if with_clause:
                return {'valid': True}  # CTE在_execute_query中处理
            # UNION/UNION ALL 支持已实现
            # EXCEPT/INTERSECT 支持已实现
            # 窗口函数支持已实现 (ROW_NUMBER, RANK, DENSE_RANK)

            return {'valid': True}

        except Exception as e:
            return {
                'valid': False,
                'error': f'SQL验证失败: {str(e)}'
            }

    def _replace_cn_columns_in_sql(self, sql: str, worksheets_data: Dict[str, pd.DataFrame]) -> str:
        """
        将SQL中的中文列名替换为英文列名(在sqlglot解析前).

        双行表头的游戏配置表中,第1行是中文描述,第2行是英文字段名.
        策划习惯用中文名查询,但SQL引擎需要英文列名.
        本方法在SQL文本层面做替换,避免给DataFrame添加临时列.

        Args:
            sql: 原始SQL语句
            worksheets_data: 已加载的工作表数据

        Returns:
            str: 替换后的SQL语句
        """
        if not hasattr(self, '_header_descriptions') or not self._header_descriptions:
            return sql

        # 收集所有中文->英文映射(去重)
        cn_to_en = {}
        for sheet_name, desc_map in self._header_descriptions.items():
            for eng_name, cn_desc in desc_map.items():
                if cn_desc and cn_desc != eng_name:
                    cn_to_en[cn_desc] = eng_name

        if not cn_to_en:
            return sql

        # 按中文列名长度降序排列,避免短名称部分匹配长名称
        sorted_names = sorted(cn_to_en.keys(), key=len, reverse=True)

        # 用正则替换:只替换SQL标识符位置(非字符串字面量中的中文)
        # 策略:先把字符串字面量占位保护,替换中文标识符,再恢复字符串
        string_literals = []
        protected_sql = sql

        # 保护单引号字符串
        def protect_string(match):
            """保护字符串字面量不被替换.

            Args:
                match: 正则匹配对象
            """
            string_literals.append(match.group(0))
            return f'__PROTECTED_STR_{len(string_literals) - 1}__'

        protected_sql = re.sub(r"'[^']*'", protect_string, protected_sql)

        # 替换中文列名
        for cn_name in sorted_names:
            en_name = cn_to_en[cn_name]
            protected_sql = re.sub(re.escape(cn_name), en_name, protected_sql)

        # 恢复字符串字面量
        for i, s in enumerate(string_literals):
            protected_sql = protected_sql.replace(f'__PROTECTED_STR_{i}__', s)

        return protected_sql

    def _generate_empty_result_suggestion(self, parsed_sql, df_before_where, worksheets_data):
        """分析WHERE条件类型,生成智能空结果建议"""
        where_clause = parsed_sql.args.get('where')
        if not where_clause:
            return '查询返回0行数据.表可能为空,请检查数据是否已录入.'

        total_rows = len(df_before_where)
        if total_rows == 0:
            return '查询返回0行数据.工作表本身没有数据行.'

        hints = []
        condition = where_clause.this

        # 分析条件树,收集条件类型和涉及的列
        eq_conditions = []  # 等值条件
        range_conditions = []  # 范围条件
        like_conditions = []  # LIKE条件
        in_conditions = []  # IN条件
        between_conditions = []  # BETWEEN条件
        null_conditions = []  # IS NULL条件

        self._collect_condition_types(condition, eq_conditions, range_conditions,
                                       like_conditions, in_conditions, between_conditions, null_conditions)

        # 等值条件:提示列的唯一值
        for col, val in eq_conditions:
            if col in df_before_where.columns:
                unique_vals = df_before_where[col].dropna().unique()
                if len(unique_vals) <= 20:
                    vals_str = ', '.join(str(v) for v in unique_vals[:10])
                    if len(unique_vals) > 10:
                        vals_str += f' ... 共{len(unique_vals)}个'
                    hints.append(f'• 列"{col}"的值为: {vals_str}')
                else:
                    hints.append(f'• 列"{col}"有{len(unique_vals)}个不同值,"{val}"不在其中')

        # 范围条件 + BETWEEN条件:提示列的实际范围(两者提示文本相同)
        range_and_between = [(col, op, val) for col, op, val in range_conditions] + \
                            [(col, f'{low}~{high}', None) for col, low, high in between_conditions]
        range_cols_seen = set()
        for item in range_and_between:
            col = item[0]
            if col in range_cols_seen:
                continue
            range_cols_seen.add(col)
            if col in df_before_where.columns:
                numeric = pd.to_numeric(df_before_where[col], errors='coerce').dropna()
                if len(numeric) > 0:
                    hints.append(f'• 列"{col}"的实际范围: {numeric.min():.2f} ~ {numeric.max():.2f}')
                else:
                    hints.append(f'• 列"{col}"不是数值列,无法用范围条件比较')

        # LIKE条件:提示匹配情况
        for col, pattern in like_conditions:
            if col in df_before_where.columns:
                sample = df_before_where[col].dropna().astype(str).head(5).tolist()
                hints.append(f'• 列"{col}"的样本数据: {", ".join(sample)}')

        # IN条件:提示实际存在的值
        for col, vals in in_conditions:
            if col in df_before_where.columns:
                unique_vals = set(df_before_where[col].dropna().unique())
                matched = unique_vals & set(vals)
                if not matched:
                    hints.append(f'• 列"{col}"中不包含指定的任何值')

        # IS NULL条件
        for col, is_null in null_conditions:
            if col in df_before_where.columns:
                null_count = df_before_where[col].isna().sum()
                if is_null and null_count == 0:
                    hints.append(f'• 列"{col}"没有空值')
                elif not is_null and null_count == total_rows:
                    hints.append(f'• 列"{col}"全部为空值')

        # 多条件提示
        total_conditions = len(eq_conditions) + len(range_conditions) + len(like_conditions) + len(in_conditions) + len(between_conditions) + len(null_conditions)
        if total_conditions > 1:
            hints.append('• 多个AND条件同时满足的行可能不存在,尝试减少条件或改用OR')

        # 通用提示
        hints.append(f'• 源表共{total_rows}行,WHERE过滤后为0行')
        hints.append('• 可用 DESCRIBE 查看表结构,或去掉WHERE先查看全部数据')

        return '查询返回0行数据.分析:\n' + '\n'.join(hints)

    def _collect_condition_types(self, condition, eq, rng, like, in_list, between, null_list):
        """递归收集WHERE条件树中的各类条件"""
        if isinstance(condition, exp.EQ):
            col, col_table = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            if col:
                eq.append((col, val))
        elif isinstance(condition, (exp.GT, exp.GTE, exp.LT, exp.LTE)):
            col, col_table = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            op_map = {exp.GT: '>', exp.GTE: '>=', exp.LT: '<', exp.LTE: '<='}
            if col:
                rng.append((col, op_map.get(type(condition), '?'), val))
        elif isinstance(condition, exp.Like):
            col, col_table = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            if col:
                like.append((col, val))
        elif isinstance(condition, exp.In):
            col, col_table = self._extract_column_name(condition.this)
            vals = []
            if hasattr(condition, 'expressions'):
                for e in condition.expressions:
                    v = self._extract_literal_value(e)
                    if v is not None:
                        vals.append(v)
            if col and vals:
                in_list.append((col, vals))
        elif isinstance(condition, exp.Between):
            col, col_table = self._extract_column_name(condition.this)
            low = self._extract_literal_value(condition.args.get('low'))
            high = self._extract_literal_value(condition.args.get('high'))
            if col:
                between.append((col, low, high))
        elif isinstance(condition, exp.Is):
            col, col_table = self._extract_column_name(condition.this)
            if col:
                null_list.append((col, True))
        elif isinstance(condition, exp.Not):
            inner = condition.this
            if isinstance(inner, exp.Is):
                col, col_table = self._extract_column_name(inner.this)
                if col:
                    null_list.append((col, False))
            else:
                self._collect_condition_types(inner, eq, rng, like, in_list, between, null_list)
        elif isinstance(condition, (exp.And, exp.Or)):
            for child in condition.flatten():
                if child is not condition:
                    self._collect_condition_types(child, eq, rng, like, in_list, between, null_list)

    def _extract_column_name(self, expr):
        """从表达式中提取列名,支持表别名格式(如 r.名称)"""
        if isinstance(expr, exp.Column):
            # 处理表别名格式,如 r.名称 -> 名称, 返回 (列名, 表别名)
            if hasattr(expr, 'table') and expr.table:
                return f"{expr.table}.{expr.name}", expr.table
            else:
                return expr.name, None
        return None, None

    def _resolve_column_name(self, col_name: str, df) -> str:
        """解析列名,支持表别名格式(如 r.名称)"""
        if '.' in col_name:
            # 处理表别名格式,如 r.名称
            table_part, col_part = col_name.split('.', 1)
            # 从 _table_aliases 获取真实的表名
            resolved_table = self._table_aliases.get(table_part, table_part)
            
            # 优先检查用户使用的别名格式是否直接存在
            alias_col = f"{table_part}.{col_part}"
            if alias_col in df.columns:
                return alias_col
            
            # 检查JOIN后pandas添加的后缀格式(table_part_列名)
            pandas_suffix_col = f"{table_part}_{col_part}"
            if pandas_suffix_col in df.columns:
                return pandas_suffix_col
            
            # 尝试其他可能的别名格式
            # 如果用户使用的是原始表名,检查原始表名+列名
            original_col = f"{resolved_table}_{col_part}"
            if original_col in df.columns:
                return original_col
                
            # 最后尝试原始列名
            if col_part in df.columns:
                return col_part
                
        return col_name

    def _extract_literal_value(self, expr):
        """从表达式中提取字面值(委托_parse_literal_value统一处理)"""
        if isinstance(expr, exp.Literal):
            return self._parse_literal_value(expr)
        return None

    def _generate_having_empty_suggestion(self, having_expr, df_before_having) -> str:
        """生成HAVING导致空结果时的智能建议

        Args:
            having_expr: HAVING表达式(完整的having clause)
            df_before_having: HAVING过滤前的聚合结果DataFrame
        """
        hints = ['\nHAVING分析:']
        hints.append(f'• GROUP BY聚合后有{len(df_before_having)}组数据')

        condition = having_expr.this
        col, col_table = self._extract_column_name(condition.left)
        val = self._extract_literal_value(condition.right)

        # 策略0:通过_having_agg_alias_map直接查找聚合表达式对应的SELECT别名
        # 这是最可靠的方式,能正确处理中文列名(如AVG(伤害)->avg_dmg)
        if not col and hasattr(self, '_having_agg_alias_map') and self._having_agg_alias_map:
            left_sql = str(condition.left).strip()
            # 精确匹配:HAVING表达式SQL与map key完全一致
            for map_sql, alias in self._having_agg_alias_map.items():
                if map_sql == left_sql and alias in df_before_having.columns:
                    col = alias
                    break
            # 模糊匹配:去除多余空格后比较(sqlglot有时生成不一致的空格)
            if not col:
                left_normalized = ' '.join(left_sql.split())
                for map_sql, alias in self._having_agg_alias_map.items():
                    map_normalized = ' '.join(map_sql.split())
                    if (left_normalized == map_normalized or left_sql in map_sql or map_sql in left_sql) \
                            and alias in df_before_having.columns:
                        col = alias
                        break

        # HAVING中聚合函数表达式的列名可能是别名,尝试从DataFrame列匹配
        if not col:
            left_str = str(condition.left).lower()
            # 策略1:子串匹配(含中文支持)
            for c in df_before_having.columns:
                c_lower = c.lower()
                if c_lower in left_str or left_str in c_lower:
                    col = c
                    break
            # 策略2:拆分表达式中的标识符(含中文token提取)
            if not col:
                tokens = set(re.findall(r'[a-zA-Z_]+', left_str))
                cn_tokens = set(re.findall(r'[\u4e00-\u9fff]+', left_str))
                if tokens or cn_tokens:
                    for c in df_before_having.columns:
                        c_tokens = set(re.findall(r'[a-zA-Z_]+', c.lower()))
                        c_cn_tokens = set(re.findall(r'[\u4e00-\u9fff]+', c))
                        generic = {'avg', 'sum', 'count', 'min', 'max'}
                        specific = tokens - generic
                        if (specific and specific & c_tokens) or (cn_tokens and cn_tokens & c_cn_tokens):
                            col = c
                            break

        if not col or col not in df_before_having.columns:
            # 无法匹配列名,显示所有聚合列的实际范围
            if len(df_before_having.columns) > 0:
                for c in df_before_having.columns:
                    numeric = pd.to_numeric(df_before_having[c], errors='coerce').dropna()
                    if len(numeric) > 0:
                        hints.append(f'• 列"{c}"范围: {numeric.min()} ~ {numeric.max()}')
            hints.append('• HAVING条件较复杂,建议去掉HAVING先查看聚合结果')
            hints.append('• 可先去掉HAVING查看全部分组结果,再调整过滤条件')
            return '\n'.join(hints)

        numeric = pd.to_numeric(df_before_having[col], errors='coerce').dropna()
        if len(numeric) == 0:
            hints.append(f'• 列"{col}"没有数值数据')
            hints.append('• 可先去掉HAVING查看全部分组结果,再调整过滤条件')
            return '\n'.join(hints)

        # 比较运算符(GT/GTE/LT/LTE)使用分发表
        op_type = type(condition)
        if op_type in self._HAVING_OPS:
            stat_func, op_str, label = self._HAVING_OPS[op_type]
            stat_val = getattr(numeric, stat_func)()
            hints.append(f'• 列"{col}"的{label}值为{stat_val},HAVING要求 {op_str}{val},无满足条件的组')
        elif isinstance(condition, exp.EQ):
            unique_vals = df_before_having[col].dropna().unique()
            if len(unique_vals) <= 10:
                vals_str = ', '.join(str(v) for v in unique_vals)
                hints.append(f'• 列"{col}"的值为: {vals_str},不等于{val}')
            else:
                hints.append(f'• 列"{col}"有{len(unique_vals)}个不同值,不等于{val}')
        else:
            hints.append(f'• HAVING条件较复杂,建议去掉HAVING先查看聚合结果')

        hints.append('• 可先去掉HAVING查看全部分组结果,再调整过滤条件')
        return '\n'.join(hints)

    def _suggest_column_name(self, col_name: str, available_cols: List[str], max_suggestions: int = 3) -> str:
        """
        当列名不存在时,用编辑距离找出最相似的列名作为建议.
        同时检查中文列名描述(双行表头场景).

        Args:
            col_name: 用户输入的列名
            available_cols: 可用的列名列表
            max_suggestions: 最多返回几个建议

        Returns:
            str: 格式化的建议字符串
        """
        if not available_cols:
            return ""

        matches = difflib.get_close_matches(col_name, available_cols, n=max_suggestions, cutoff=0.3)
        if matches:
            return f" 你是否想用: {', '.join(matches)}?"

        # 英文列名匹配不到时,尝试匹配中文列名描述(双行表头)
        if hasattr(self, '_header_descriptions') and self._header_descriptions:
            cn_names = []
            for sheet_name, desc_map in self._header_descriptions.items():
                for eng_name, cn_desc in desc_map.items():
                    if cn_desc:
                        cn_names.append(cn_desc)
            if cn_names:
                cn_matches = difflib.get_close_matches(col_name, cn_names, n=max_suggestions, cutoff=0.3)
                if cn_matches:
                    # 反查中文->英文映射
                    en_lookup = {}
                    for desc_map in self._header_descriptions.values():
                        for eng_name, cn_desc in desc_map.items():
                            if cn_desc:
                                en_lookup[cn_desc] = eng_name
                    mapped = [f"{cn}({en_lookup.get(cn, '?')})" for cn in cn_matches]
                    return f" 中文名称匹配: {', '.join(mapped)}"

        return ""

    def _apply_union_order_by(self, df, order_clause=None) -> pd.DataFrame:
        """
        对 UNION 合并后的 DataFrame 应用 ORDER BY 排序

        Args:
            df: 合并后的 DataFrame
            order_clause: sqlglot Order expression

        Returns:
            pd.DataFrame: 排序后的 DataFrame
        """
        if not order_clause or df.empty:
            return df

        sort_columns = []
        sort_ascending = []

        for ordered in order_clause.find_all(exp.Ordered):
            col_expr = ordered.this
            # 获取列名
            if isinstance(col_expr, exp.Column):
                col_name = col_expr.name
            elif isinstance(col_expr, exp.Identifier):
                col_name = col_expr.this
            else:
                col_name = str(col_expr)

            # 检查是否使用表别名限定符
            if '.' in col_name:
                col_name = col_name.split('.')[-1]

            # 尝试列名匹配(包括中文列名映射)
            if col_name not in df.columns:
                # 搜索中文列名映射
                for cn_name, en_name in self._cn_to_en_map.items():
                    if col_name == cn_name:
                        col_name = en_name
                        break

            if col_name in df.columns:
                sort_columns.append(col_name)
                sort_ascending.append(not ordered.args.get('desc', False))

        if sort_columns:
            df = df.sort_values(sort_columns, ascending=sort_ascending).reset_index(drop=True)

        return df

    def _execute_union(
        self,
        parsed_sql: exp.Expression,
        worksheets_data: Dict[str, pd.DataFrame],
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """
        执行 UNION / UNION ALL 查询

        从 Union 表达式中提取所有 SELECT 语句,分别执行后合并结果.
        UNION 去重,UNION ALL 保留所有行.
        支持 ORDER BY 和 LIMIT 应用于合并后的结果.

        Args:
            parsed_sql: 解析后的 Union 表达式
            worksheets_data: 工作表数据
            limit: 结果限制

        Returns:
            pd.DataFrame: 合并后的查询结果
        """
        # 递归提取所有 SELECT 语句
        def _extract_selects(node):
            if isinstance(node, exp.Select):
                return [node]
            elif isinstance(node, exp.Union):
                selects = []
                # this 可能是 Union(链式)或 Select
                selects.extend(_extract_selects(node.this))
                # expression 是右侧的 Select 或 Union
                selects.extend(_extract_selects(node.expression))
                return selects
            return []

        selects = _extract_selects(parsed_sql)
        if not selects:
            raise ValueError("UNION 查询中未找到有效的 SELECT 语句")

        # 执行每个 SELECT 并收集结果
        result_dfs = []
        for i, select_sql in enumerate(selects):
            # 确保每个 SELECT 有自己的表别名上下文
            df = self._execute_query(select_sql, worksheets_data, limit=None)
            result_dfs.append(df)

        # 合并所有结果(列名对齐)
        if not result_dfs:
            return pd.DataFrame()

        # 以第一个 SELECT 的列名为基准,统一列名
        base_columns = list(result_dfs[0].columns)
        aligned_dfs = []
        for df in result_dfs:
            aligned = df.reindex(columns=base_columns)
            aligned_dfs.append(aligned)

        combined = pd.concat(aligned_dfs, ignore_index=True)

        # UNION(去重) vs UNION ALL(保留重复)
        is_union_all = not parsed_sql.args.get('distinct', True)
        if not is_union_all:
            combined = combined.drop_duplicates().reset_index(drop=True)

        # 应用 ORDER BY(如果有,sqlglot 将其放在外层 Union 上)
        order_clause = parsed_sql.args.get('order')
        if order_clause:
            # 构造一个最小化的 Select 用于 _apply_order_by 的签名
            # _apply_order_by(self, parsed_sql, df, select_aliases) 期望完整的 parsed_sql
            # 但 UNION 的 ORDER BY 是独立的,直接解析排序列
            combined = self._apply_union_order_by(combined, order_clause)

        # 应用 LIMIT(如果有)
        union_limit = self._extract_int_value(parsed_sql.args.get('limit'))
        if union_limit is not None:
            combined = combined.head(union_limit)

        # 应用外部传入的 limit
        if limit is not None:
            combined = combined.head(limit)

        return combined

    def _execute_except_intersect(
        self,
        parsed_sql: exp.Expression,
        worksheets_data: Dict[str, pd.DataFrame],
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """
        执行 EXCEPT / INTERSECT 查询

        从 Except/Intersect 表达式中提取所有 SELECT 语句,分别执行后进行集合操作.
        EXCEPT: 返回第一个查询结果中不存在于第二个查询结果的行(差集)
        INTERSECT: 返回两个查询结果中都存在的行(交集)

        Args:
            parsed_sql: 解析后的 Except/Intersect 表达式
            worksheets_data: 工作表数据
            limit: 结果限制

        Returns:
            pd.DataFrame: 集合操作后的查询结果
        """
        # 确定操作类型
        is_except = isinstance(parsed_sql, exp.Except)
        is_intersect = isinstance(parsed_sql, exp.Intersect)

        # 递归提取所有 SELECT 语句(支持链式 EXCEPT/INTERSECT)
        def _extract_selects(node):
            if isinstance(node, exp.Select):
                return [node]
            elif isinstance(node, (exp.Union, exp.Except, exp.Intersect)):
                selects = []
                # this 可能是 Union/Except/Intersect(链式)或 Select
                selects.extend(_extract_selects(node.this))
                # expression 是右侧的 Select 或 Union/Except/Intersect
                selects.extend(_extract_selects(node.expression))
                return selects
            return []

        selects = _extract_selects(parsed_sql)
        if not selects:
            raise ValueError("EXCEPT/INTERSECT 查询中未找到有效的 SELECT 语句")
        if len(selects) != 2:
            raise ValueError("EXCEPT/INTERSECT 查询必须包含恰好两个 SELECT 语句")

        # 执行两个 SELECT
        df1 = self._execute_query(selects[0], worksheets_data, limit=None)
        df2 = self._execute_query(selects[1], worksheets_data, limit=None)

        # 对齐列名
        base_columns = list(df1.columns)
        df2_aligned = df2.reindex(columns=base_columns)

        if is_except:
            # EXCEPT: 差集 - df1 中不在 df2 中的行
            # 使用 merge(how='left', indicator=True) 实现
            merged = df1.merge(df2_aligned, how='left', indicator=True, on=list(base_columns))
            # 保留只在左侧(df1)的行
            result = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge']).reset_index(drop=True)
        elif is_intersect:
            # INTERSECT: 交集 - df1 和 df2 都有的行
            # 使用 merge(how='inner') 实现
            result = df1.merge(df2_aligned, how='inner', on=list(base_columns)).drop_duplicates().reset_index(drop=True)
        else:
            return pd.DataFrame()

        # 应用 ORDER BY(如果有)
        order_clause = parsed_sql.args.get('order')
        if order_clause:
            result = self._apply_union_order_by(result, order_clause)

        # 应用 LIMIT(如果有)
        op_limit = self._extract_int_value(parsed_sql.args.get('limit'))
        if op_limit is not None:
            result = result.head(op_limit)

        # 应用外部传入的 limit
        if limit is not None:
            result = result.head(limit)

        return result

    def _has_window_function(self, parsed_sql: exp.Expression) -> bool:
        """检查SQL是否包含窗口函数"""
        return bool(parsed_sql.find(exp.Window))

    def _apply_window_functions(self, parsed_sql: exp.Expression, df) -> pd.DataFrame:
        """
        计算窗口函数并将结果添加到DataFrame
        支持: ROW_NUMBER, RANK, DENSE_RANK
        语法: func() OVER ([PARTITION BY col ...] ORDER BY col [ASC|DESC] ...)
        """
        if not self._has_window_function(parsed_sql):
            return df

        df = df.copy()

        # 构建SELECT别名映射(用于将聚合表达式映射到别名,如AVG(damage)->avg_dmg)
        select_alias_map = {}
        for select_expr in parsed_sql.expressions:
            if isinstance(select_expr, exp.Alias):
                alias_name = select_expr.alias
                original = select_expr.this
                # 将原始表达式文本作为key
                expr_key = str(original).strip()
                select_alias_map[expr_key] = alias_name

        for select_expr in parsed_sql.expressions:
            # 跳过 SELECT *
            if isinstance(select_expr, exp.Star):
                continue

            # 提取别名和原始表达式
            if isinstance(select_expr, exp.Alias):
                alias_name = select_expr.alias
                original_expr = select_expr.this
            else:
                alias_name = None
                original_expr = select_expr

            if not isinstance(original_expr, exp.Window):
                continue

            # 确定列名
            col_name = alias_name or f"_window_{len([c for c in df.columns if c.startswith('_window_')])}"

            # 计算窗口函数
            result = self._compute_window_function(original_expr, df, select_alias_map)
            df[col_name] = result

        return df

    def _compute_window_function(self, window_expr: exp.Window, df: pd.DataFrame,
                                  select_alias_map: Optional[Dict] = None) -> pd.Series:
        """计算单个窗口函数,返回结果Series"""
        func_type = type(window_expr.this).__name__

        # 支持的窗口函数类型
        supported_funcs = {'RowNumber', 'Rank', 'DenseRank'}
        if func_type not in supported_funcs:
            raise ValueError(f"不支持的窗口函数: {func_type}.支持的: ROW_NUMBER, RANK, DENSE_RANK")

        if select_alias_map is None:
            select_alias_map = {}

        # 解析 PARTITION BY
        partition_by = window_expr.args.get('partition_by', [])
        partition_cols = []
        for col in partition_by:
            col_name = col.name if hasattr(col, 'name') and col.name else str(col)
            partition_cols.append(col_name)

        # 解析 ORDER BY
        order = window_expr.args.get('order')
        order_cols = []
        ascending = []
        if order:
            for ordered_expr in order.expressions:
                col = ordered_expr.this
                col_name = col.name if hasattr(col, 'name') and col.name else str(col)
                # 如果列名不在DataFrame中,尝试映射
                if col_name not in df.columns:
                    col_name = self._resolve_window_column(col_name, df.columns, select_alias_map)
                order_cols.append(col_name)
                ascending.append(not ordered_expr.args.get('desc', False))

        # 验证列存在
        for col in partition_cols + order_cols:
            if col not in df.columns:
                suggestion = self._suggest_column_name(col, list(df.columns))
                raise ValueError(f"窗口函数中列 '{col}' 不存在.可用列: {list(df.columns)}.{suggestion}")

        # 窗口函数分发表
        _window_dispatch = {
            'RowNumber': self._compute_row_number,
            'Rank': self._compute_rank,
            'DenseRank': self._compute_dense_rank,
        }
        handler = _window_dispatch.get(func_type)
        if handler:
            return handler(df, partition_cols, order_cols, ascending)
        raise ValueError(f"不支持的窗口函数: {func_type}")

    def _resolve_window_column(self, col_name: str, df_columns: list,
                                select_alias_map: Dict[str, str]) -> str:
        """解析窗口函数中的列名(支持聚合表达式->别名映射)"""
        # 1. 直接在SELECT别名映射中查找
        if col_name in select_alias_map:
            alias = select_alias_map[col_name]
            if alias in df_columns:
                return alias

        # 2. 尝试聚合函数名匹配
        agg_funcs = {'AVG', 'SUM', 'COUNT', 'MAX', 'MIN'}
        for func in agg_funcs:
            if col_name.upper().startswith(func):
                match = re.match(rf'{func}\s*\(\s*(.+?)\s*\)', col_name, re.IGNORECASE)
                if match:
                    inner_col = match.group(1).strip()
                    candidates = [
                        f"{func.lower()}_{inner_col}",
                        f"{func.lower()}_{inner_col.lower()}",
                        inner_col,
                    ]
                    for c in candidates:
                        if c in df_columns:
                            return c

        return col_name  # 未找到映射,返回原名

    def _compute_row_number(self, df: pd.DataFrame, partition_cols: list,
                            order_cols: list, ascending: list) -> pd.Series:
        """ROW_NUMBER: 分区内从1开始的连续编号"""
        if not partition_cols and not order_cols:
            # 无PARTITION BY也无ORDER BY: 按原始行顺序编号
            return pd.Series(range(1, len(df) + 1), index=df.index, dtype=int)

        if partition_cols:
            grouped = df.groupby(partition_cols, sort=False)
        else:
            grouped = None  # 全表作为一个分区

        def assign_row_number(group):
            """为分组内的行分配行号.

            Args:
                group: pandas分组对象
            """
            if order_cols:
                sorted_group = group.sort_values(order_cols, ascending=ascending)
                result = pd.Series(range(1, len(sorted_group) + 1), index=sorted_group.index, dtype=int)
                return result.reindex(group.index)
            else:
                return pd.Series(range(1, len(group) + 1), index=group.index, dtype=int)

        if grouped is not None:
            result = grouped.apply(assign_row_number, include_groups=False)
            # groupby.apply可能返回MultiIndex Series,需要展平
            if isinstance(result.index, pd.MultiIndex):
                result = result.droplevel(result.index.names[:-1])
        else:
            result = assign_row_number(df)

        return result

    def _compute_rank(self, df: pd.DataFrame, partition_cols: list,
                      order_cols: list, ascending: list) -> pd.Series:
        """RANK: 相同值相同排名,下一个排名跳过(1,2,2,4)"""
        if not order_cols:
            raise ValueError("RANK() 窗口函数需要 ORDER BY 子句")

        if partition_cols:
            grouped = df.groupby(partition_cols, sort=False)
        else:
            grouped = None

        def assign_rank(group):
            """为分组内的行分配排名.

            Args:
                group: pandas分组对象
            """
            sorted_group = group.sort_values(order_cols, ascending=ascending)
            # 使用pandas rank(method='first')模拟RANK行为
            # RANK: 相同值取相同排名,下一个排名跳过
            rank_series = sorted_group[order_cols[0]].rank(method='min', ascending=ascending[0])
            rank_series = rank_series.astype(int)
            return rank_series.reindex(group.index)

        if grouped is not None:
            result = grouped.apply(assign_rank, include_groups=False)
            if isinstance(result.index, pd.MultiIndex):
                result = result.droplevel(result.index.names[:-1])
        else:
            result = assign_rank(df)

        return result

    def _compute_dense_rank(self, df: pd.DataFrame, partition_cols: list,
                            order_cols: list, ascending: list) -> pd.Series:
        """DENSE_RANK: 相同值相同排名,下一个排名不跳过(1,2,2,3)"""
        if not order_cols:
            raise ValueError("DENSE_RANK() 窗口函数需要 ORDER BY 子句")

        if partition_cols:
            grouped = df.groupby(partition_cols, sort=False)
        else:
            grouped = None

        def assign_dense_rank(group):
            """为分组内的行分配密集排名.

            Args:
                group: pandas分组对象
            """
            sorted_group = group.sort_values(order_cols, ascending=ascending)
            # DENSE_RANK: 相同值取相同排名,下一个排名连续
            rank_series = sorted_group[order_cols[0]].rank(method='dense', ascending=ascending[0])
            rank_series = rank_series.astype(int)
            return rank_series.reindex(group.index)

        if grouped is not None:
            result = grouped.apply(assign_dense_rank, include_groups=False)
            if isinstance(result.index, pd.MultiIndex):
                result = result.droplevel(result.index.names[:-1])
        else:
            result = assign_dense_rank(df)

        return result

    def _execute_query(
        self,
        parsed_sql: exp.Expression,
        worksheets_data: Dict[str, pd.DataFrame],
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """
        执行解析后的SQL查询

        Args:
            parsed_sql: 解析后的SQL表达式
            worksheets_data: 工作表数据
            limit: 结果限制

        Returns:
            pd.DataFrame: 查询结果
        """
        # 处理CTE (WITH ... AS ...)
        # 兼容sqlglot不同版本:arg key可能是'with'或'with_'
        _with_key = 'with' if parsed_sql.args.get('with') else 'with_'
        with_clause = parsed_sql.args.get(_with_key)
        if with_clause:
            # 复制worksheets_data避免修改原始数据,逐步添加CTE结果
            cte_data = dict(worksheets_data)
            for cte_expr in with_clause.expressions:
                cte_name = cte_expr.alias
                cte_query = cte_expr.this  # inner Select
                try:
                    # 每个CTE在已有的cte_data上执行(支持CTE引用前面的CTE)
                    cte_result = self._execute_query(cte_query, cte_data, limit=None)
                    cte_data[cte_name] = cte_result
                except Exception as e:
                    raise ValueError(f"CTE '{cte_name}' 执行失败: {e}")
            # 从parsed_sql中移除with子句,让后续逻辑正常处理
            parsed_sql = parsed_sql.copy()
            parsed_sql.set(_with_key, None)

        # 获取FROM子句中的表名(及可选的子查询)
        from_table, from_subquery = self._get_from_table(parsed_sql)

        # 查找表名时也搜索CTE定义的临时表
        effective_data = cte_data if with_clause else worksheets_data

        # 如果FROM是子查询,先执行子查询并将结果注入effective_data
        if from_subquery is not None:
            try:
                sub_result = self._execute_subquery(from_subquery, effective_data)
                effective_data[from_table] = sub_result
            except Exception as e:
                raise StructuredSQLError(
                    "from_subquery_error",
                    f"FROM子查询执行失败: {e}",
                    hint="请检查FROM子查询中的SQL语法和表名是否正确.",
                    context={"subquery_alias": from_table, "error": str(e)}
                )

        if from_table not in effective_data:
            raise StructuredSQLError(
                "table_not_found",
                f"表 '{from_table}' 不存在.可用表: {list(effective_data.keys())}",
                hint="请检查表名拼写,或用excel_list_sheets查看可用工作表名.",
                context={"table_requested": from_table, "available_tables": list(effective_data.keys())}
            )

        base_df = effective_data[from_table].copy()

        # 构建表别名映射
        self._table_aliases = {}
        self._table_aliases[from_table] = from_table
        # 检查FROM子句是否有别名 (FROM 技能表 a)
        from_clause = parsed_sql.args.get('from')
        if from_clause:
            # 优先使用 Table.alias 属性
            from_table_expr = from_clause.this
            if hasattr(from_table_expr, 'alias') and from_table_expr.alias:
                from_alias = from_table_expr.alias
                if isinstance(from_alias, str) and from_alias != from_table:
                    self._table_aliases[from_alias] = from_table
                    self._table_aliases[from_table] = from_table
            # 备用:遍历 TableAlias 节点
            found_from_alias = False
            for alias in from_clause.find_all(exp.TableAlias):
                parent_table = from_clause.this.name if hasattr(from_clause.this, 'name') else str(from_clause.this)
                self._table_aliases[alias.alias] = parent_table
                self._table_aliases[parent_table] = parent_table
                found_from_alias = True
            if not found_from_alias:
                for alias in from_clause.find_all(exp.Alias):
                    parent_table = from_clause.this.name if hasattr(from_clause.this, 'name') else str(from_clause.this)
                    self._table_aliases[alias.alias] = parent_table
                    self._table_aliases[parent_table] = parent_table

        # 应用JOIN子句
        joins = parsed_sql.args.get('joins')
        if joins:
            base_df = self._apply_join_clause(joins, base_df, effective_data, from_table)

        # 应用WHERE条件
        # 保存WHERE前的DataFrame,用于空结果智能建议
        base_df_before_where = base_df.copy()
        self._df_before_where = base_df_before_where
        # 保存当前工作表数据供子查询使用
        self._current_worksheets = effective_data
        base_df = self._apply_where_clause(parsed_sql, base_df)

        # 检查是否有聚合函数
        has_aggregate = self._check_has_aggregate_function(parsed_sql)

        # 应用GROUP BY和聚合
        if parsed_sql.args.get('group') or has_aggregate:
            # 有GROUP BY或有聚合函数时,应用分组聚合
            base_df = self._apply_group_by_aggregation(parsed_sql, base_df)

            # 应用HAVING条件
            has_having = parsed_sql.args.get('having') is not None
            if has_having:
                # 保存HAVING前的DataFrame,用于HAVING空结果建议
                self._df_before_having = base_df.copy()
                base_df = self._apply_having_clause(parsed_sql, base_df)
        else:
            has_having = False

        # 应用窗口函数(ROW_NUMBER, RANK, DENSE_RANK)
        # 窗口函数在GROUP BY/HAVING之后,ORDER BY/SELECT之前计算
        base_df = self._apply_window_functions(parsed_sql, base_df)

        if parsed_sql.args.get('group') or has_aggregate:
            # ORDER BY(聚合查询:在GROUP BY之后)
            if parsed_sql.args.get('order'):
                base_df = self._apply_order_by(parsed_sql, base_df)
        else:
            # 非聚合查询:提取SELECT别名,然后ORDER BY(支持引用别名和原始列),最后SELECT
            select_aliases = self._extract_select_aliases(parsed_sql)
            if parsed_sql.args.get('order'):
                base_df = self._apply_order_by(parsed_sql, base_df, select_aliases=select_aliases)

            # 应用SELECT表达式(裁剪列,计算字段,别名)
            base_df = self._apply_select_expressions(parsed_sql, base_df)

        # 应用OFFSET(在LIMIT之前)
        offset_value = self._extract_int_value(parsed_sql.args.get('offset'))
        if offset_value is not None:
            base_df = base_df.iloc[offset_value:]

        # 应用LIMIT
        limit_value = self._extract_int_value(parsed_sql.args.get('limit'))
        if limit_value is not None:
            base_df = base_df.head(limit_value)
        elif limit:
            base_df = base_df.head(limit)

        # 应用SELECT DISTINCT去重
        if parsed_sql.args.get('distinct'):
            base_df = base_df.drop_duplicates()

        return base_df

    @staticmethod
    def _extract_int_value(clause) -> Optional[int]:
        """从SQL子句中提取整数值(LIMIT/OFFSET等)"""
        if clause is None:
            return None
        if hasattr(clause, 'expression'):
            return int(clause.expression.this)
        return int(clause.this)

    def _check_has_aggregate_function(self, parsed_sql: exp.Expression) -> bool:
        """检查SQL查询是否包含聚合函数"""
        for select_expr in parsed_sql.expressions:
            if self._is_aggregate_function(select_expr):
                return True
        return False

    def _apply_select_expressions(self, parsed_sql: exp.Expression, df) -> pd.DataFrame:
        """
        应用SELECT表达式(非聚合查询)
        处理计算字段,别名等

        Args:
            parsed_sql: 解析后的SQL表达式
            df: 数据DataFrame

        Returns:
            pd.DataFrame: 处理后的DataFrame
        """
        result_data = {}
        ordered_columns = []
        used_aliases = set()  # 跟踪已使用的别名,避免重复

        # 按SELECT表达式顺序处理列
        for i, select_expr in enumerate(parsed_sql.expressions):
            if isinstance(select_expr, exp.Star):
                # SELECT *: 返回所有列
                for col in df.columns:
                    if col not in result_data:
                        result_data[col] = df[col]
                        ordered_columns.append(col)
                continue

            # 处理别名
            alias_name, original_expr = self._extract_select_alias(select_expr, i)

            # 处理重复别名:当多个SELECT列解析为相同名称时,添加表前缀
            if alias_name in used_aliases and isinstance(original_expr, exp.Column):
                table_part = original_expr.table if hasattr(original_expr, 'table') and original_expr.table else None
                if table_part:
                    qualified_alias = f"{table_part}.{alias_name}"
                    if qualified_alias not in used_aliases:
                        alias_name = qualified_alias
                    else:
                        alias_name = f"{table_part}_{alias_name}_{i}"
                else:
                    alias_name = f"{alias_name}_{i}"
            used_aliases.add(alias_name)

            # 解包Paren: (伤害 * 1.2) → 伤害 * 1.2
            while isinstance(original_expr, exp.Paren):
                original_expr = original_expr.this

            # 计算表达式值
            try:
                if isinstance(original_expr, exp.Column):
                    # 普通列引用(支持表限定符 a.column)
                    column_name = original_expr.name
                    table_part = original_expr.table if hasattr(original_expr, 'table') and original_expr.table else None
                    qualified = f"{table_part}.{column_name}" if table_part else None

                    # 先尝试直接使用qualified列名
                    if qualified and qualified in df.columns:
                        result_data[alias_name] = df[qualified]
                    # 如果qualified不存在,尝试映射
                    elif table_part:
                        # 使用_expression_to_column_reference进行映射
                        try:
                            mapped_column = self._expression_to_column_reference(original_expr, df)
                            # 去掉反引号
                            mapped_column = mapped_column.strip('`')
                            result_data[alias_name] = df[mapped_column]
                        except Exception:
                            # 修复:JOIN表别名映射失败时的回退逻辑
                            # 尝试查找可能的pandas merge后缀格式 (_x/_y)
                            possible_columns = [
                                f"{table_part}.{column_name}",  # 用户原始别名格式
                                f"{table_part}_{column_name}",  # table_part_列名格式
                                f"{column_name}_x",             # _x后缀格式
                                f"{column_name}_y",             # _y后缀格式
                                f"{table_part}_x",              # table_part_x格式
                                f"{table_part}_y",              # table_part_y格式
                                column_name                     # 无表前缀的原始列名
                            ]
                            
                            # 去重并检查存在的列
                            for possible_col in possible_columns:
                                if possible_col in df.columns:
                                    result_data[alias_name] = df[possible_col]
                                    break
                            else:
                                # 所有可能的映射都失败,尝试直接使用列名
                                if column_name in df.columns:
                                    result_data[alias_name] = df[column_name]
                                else:
                                    suggestion = self._suggest_column_name(column_name, list(df.columns))
                                    raise StructuredSQLError(
                                        "column_not_found",
                                        f"列 '{qualified or column_name}' 不存在.可用列: {list(df.columns)}.{suggestion}",
                                        hint="请检查列名拼写,或用excel_get_headers查看所有可用列名.",
                                        context={"column_requested": qualified or column_name, "available_columns": list(df.columns)}
                                    )
                    elif column_name in df.columns:
                        result_data[alias_name] = df[column_name]
                    else:
                        suggestion = self._suggest_column_name(column_name, list(df.columns))
                        raise StructuredSQLError(
                            "column_not_found",
                            f"列 '{qualified or column_name}' 不存在.可用列: {list(df.columns)}.{suggestion}",
                            hint="请检查列名拼写,或用excel_get_headers查看所有可用列名.",
                            context={"column_requested": qualified or column_name, "available_columns": list(df.columns)}
                        )

                elif isinstance(original_expr, exp.Case):
                    # CASE WHEN表达式
                    result_data[alias_name] = self._evaluate_case_expression(original_expr, df)

                elif isinstance(original_expr, exp.Coalesce):
                    # COALESCE/IFNULL表达式(向量化)
                    result_data[alias_name] = self._evaluate_coalesce_vectorized(original_expr, df)

                elif self._is_string_function(original_expr):
                    # 字符串函数: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT
                    result_data[alias_name] = self._evaluate_string_function(original_expr, df)

                elif isinstance(original_expr, exp.Window):
                    # 窗口函数: ROW_NUMBER, RANK, DENSE_RANK(已由_apply_window_functions预计算)
                    if alias_name in df.columns:
                        result_data[alias_name] = df[alias_name]
                    else:
                        raise ValueError(f"窗口函数结果列 '{alias_name}' 未找到")

                elif self._is_mathematical_expression(original_expr):
                    # 数学表达式
                    result_data[alias_name] = self._evaluate_math_expression(original_expr, df)

                elif isinstance(original_expr, exp.Literal):
                    # SELECT中的字面量值(如 SELECT 1, SELECT 'hello')
                    val = self._parse_literal_value(original_expr)
                    result_data[alias_name] = pd.Series([val] * len(df), index=df.index)

                elif isinstance(original_expr, exp.Subquery):
                    # 标量子查询(如 SELECT (SELECT MAX(col) FROM t))
                    try:
                        sub_result = self._execute_subquery(original_expr, self._current_worksheets)
                        if len(sub_result) > 0 and len(sub_result.columns) > 0:
                            scalar_val = sub_result.iloc[0, 0]
                            result_data[alias_name] = pd.Series([scalar_val] * len(df), index=df.index)
                        else:
                            result_data[alias_name] = pd.Series([None] * len(df), index=df.index)
                    except Exception as sub_e:
                        raise ValueError(f"标量子查询执行失败: {sub_e}")

                else:
                    # 其他表达式,尝试作为列处理
                    if hasattr(original_expr, 'name') and original_expr.name in df.columns:
                        result_data[alias_name] = df[original_expr.name]
                    else:
                        raise ValueError(f"不支持的表达式: {original_expr}")

                ordered_columns.append(alias_name)

            except Exception as e:
                # 表达式处理失败,尝试返回原始值
                if hasattr(original_expr, 'name') and original_expr.name in df.columns:
                    result_data[alias_name] = df[original_expr.name]
                    ordered_columns.append(alias_name)
                else:
                    raise ValueError(f"处理SELECT表达式失败: {e}")

        # 构建结果DataFrame,保持SELECT顺序
        if result_data:
            result_df = pd.DataFrame(result_data)
            # 按照SQL SELECT顺序重新排列列
            result_df = result_df[ordered_columns]
            return result_df
        else:
            return df

    def _extract_select_alias(self, select_expr, index: int) -> tuple:
        """
        从SELECT表达式提取别名和原始表达式.
        复用于 _apply_select_expressions 和 _apply_group_by_aggregation.

        Args:
            select_expr: SELECT表达式节点
            index: 表达式索引(用于生成默认别名 col_N)

        Returns:
            (alias_name, original_expr) 元组
        """
        if isinstance(select_expr, exp.Alias):
            return select_expr.alias, select_expr.this
        # 无别名:优先用列名,否则用聚合别名,最后 col_N
        if isinstance(select_expr, exp.Column):
            # 保留表别名前缀(如 r.名称 -> 别名为 "r.名称")
            table_part = select_expr.table if hasattr(select_expr, 'table') and select_expr.table else None
            if table_part:
                return f"{table_part}.{select_expr.name}", select_expr
            return select_expr.name, select_expr
        if self._is_aggregate_function(select_expr):
            return self._generate_aggregate_alias(select_expr), select_expr
        if hasattr(select_expr, 'name') and select_expr.name:
            return select_expr.name, select_expr
        return f"col_{index}", select_expr

    def _is_mathematical_expression(self, expr) -> bool:
        """检查是否为数学表达式"""
        return isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div, exp.Mod))

    # 数学运算符分发表:二元运算符统一处理
    _MATH_BINARY_OPS = {
        exp.Add: operator.add,
        exp.Sub: operator.sub,
        exp.Mul: operator.mul,
        exp.Div: operator.truediv,
        exp.Mod: operator.mod,
    }

    # 比较运算符分发表:逐行条件评估统一处理
    _COMPARISON_OPS = {
        exp.EQ: lambda l, r: l == r,
        exp.NEQ: lambda l, r: l != r,
        exp.GT: lambda l, r: _safe_float_comparison(l, r, '>'),
        exp.GTE: lambda l, r: _safe_float_comparison(l, r, '>='),
        exp.LT: lambda l, r: _safe_float_comparison(l, r, '<'),
        exp.LTE: lambda l, r: _safe_float_comparison(l, r, '<='),
    }

    # 复杂表达式类型集合:WHERE子句逐行过滤触发条件
    _COMPLEX_EXPR_TYPES = frozenset({
        exp.Coalesce, exp.Case, exp.Exists,
        exp.Upper, exp.Lower, exp.Trim, exp.Length,
        exp.Concat, exp.Replace, exp.Substring, exp.Left, exp.Right,
        exp.Add, exp.Sub, exp.Mul, exp.Div, exp.Mod,
    })

    # HAVING空结果建议分发表:(stat_func, op_str, label)
    _HAVING_OPS = {
        exp.GT:  ('max', '>',  '最大'),
        exp.GTE: ('max', '>=', '最大'),
        exp.LT:  ('min', '<',  '最小'),
        exp.LTE: ('min', '<=', '最小'),
    }

    # JOIN类型分发表:(side, kind) -> how
    _JOIN_KIND_MAP = {
        ('LEFT', None): 'left',
        (None, 'LEFT'): 'left',
        ('RIGHT', None): 'right',
        (None, 'RIGHT'): 'right',
        ('FULL', None): 'outer',
        (None, 'FULL'): 'outer',
        ('INNER', None): 'inner',
        (None, 'INNER'): 'inner',
        (None, 'CROSS'): 'cross',
    }

    # Pandas条件运算符分发表:SQL条件->pandas query字符串
    _PANDAS_OPS = {
        exp.EQ: '==',
        exp.NEQ: '!=',
        exp.GT: '>',
        exp.GTE: '>=',
        exp.LT: '<',
        exp.LTE: '<=',
    }

    def _evaluate_math_expression(self, expr, df: pd.DataFrame):
        """计算数学表达式"""
        op_type = type(expr)
        if op_type in self._MATH_BINARY_OPS:
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return self._MATH_BINARY_OPS[op_type](left, right)
        elif isinstance(expr, exp.Column):
            return df[expr.name]
        elif isinstance(expr, exp.Literal):
            return self._expression_to_value(expr, df)
        elif isinstance(expr, exp.Coalesce):
            # COALESCE在数学表达式中(向量化)
            return self._evaluate_coalesce_vectorized(expr, df)
        else:
            raise ValueError(f"不支持的数学表达式部分: {expr}")

    def _is_string_function(self, expr) -> bool:
        """检查是否为字符串函数"""
        return isinstance(expr, (exp.Upper, exp.Lower, exp.Trim, exp.Length,
                                exp.Concat, exp.Replace, exp.Substring, exp.Left, exp.Right))

    # 简单字符串函数分发表:一元操作,统一模式 val_series.astype(str).str.<op>()
    _SIMPLE_STR_OPS = {
        exp.Upper: 'upper',
        exp.Lower: 'lower',
        exp.Trim: 'strip',
        exp.Length: 'len',
    }

    def _evaluate_string_function(self, expr, df) -> pd.Series:
        """计算字符串函数,返回pd.Series"""
        func_type = type(expr)

        # 简单字符串函数:分发表处理
        if func_type in self._SIMPLE_STR_OPS:
            val_series = self._expr_to_series(expr.this, df)
            return getattr(val_series.astype(str).str, self._SIMPLE_STR_OPS[func_type])()

        func_name = func_type.__name__.lower()

        if func_name == 'concat':
            # CONCAT(a, b, ...) -- expressions列表包含所有参数
            parts = [self._expr_to_series(arg, df).astype(str) for arg in expr.expressions]
            if parts:
                result = parts[0]
                for p in parts[1:]:
                    result = result + p
                return result
            return pd.Series([''] * len(df), index=df.index)

        if func_name == 'replace':
            # REPLACE(str, old, new) -- sqlglot: this=string, expression=old, replacement=new
            val_series = self._expr_to_series(expr.this, df).astype(str)
            old_val = self._get_arg(expr, 'expression', '', str)
            new_val = self._get_arg(expr, 'replacement', '', str)
            return val_series.str.replace(old_val, new_val, regex=False)

        if func_name in ('substring', 'left', 'right'):
            val_series = self._expr_to_series(expr.this, df).astype(str)
            if func_name == 'substring':
                start = self._get_arg(expr, 'start', 1, int) - 1
                length = self._get_arg(expr, 'length', len(val_series.iloc[0]), int)
                return val_series.str.slice(start, start + length)
            if func_name == 'left':
                n = self._get_arg(expr, 'expression', 1, int)
                return val_series.str.slice(0, n)
            # right
            n = self._get_arg(expr, 'expression', 1, int)
            return val_series.str.slice(-n)

        raise ValueError(f"不支持的字符串函数: {func_name}")

    # 逐行字符串函数分发表:一元操作,统一模式 op(val)
    _ROW_STR_OPS = {
        'upper': lambda v: v.upper(),
        'lower': lambda v: v.lower(),
        'trim': lambda v: v.strip(),
        'length': lambda v: len(v),
    }

    def _evaluate_string_function_for_row(self, expr, row: pd.Series) -> Any:
        """逐行评估字符串函数"""
        func_name = type(expr).__name__.lower()
        val = self._get_row_value(expr.this, row)
        if val is None:
            return None
        val = str(val)

        # 简单函数:分发表处理
        if func_name in self._ROW_STR_OPS:
            return self._ROW_STR_OPS[func_name](val)

        # 复杂函数:各自独立处理
        if func_name == 'concat':
            parts = [str(self._get_row_value(arg, row) or '') for arg in expr.expressions]
            return ''.join(parts)
        if func_name == 'replace':
            old_val = self._get_arg(expr, 'expression', '', str)
            new_val = self._get_arg(expr, 'replacement', '', str)
            return val.replace(old_val, new_val)
        if func_name == 'substring':
            start = self._get_arg(expr, 'start', 1, int) - 1
            length = self._get_arg(expr, 'length', len(val), int)
            return val[start:start + length]
        if func_name == 'left':
            n = self._get_arg(expr, 'expression', 1, int)
            return val[:n]
        if func_name == 'right':
            n = self._get_arg(expr, 'expression', 1, int)
            return val[-n:] if n > 0 else ''
        return val

    def _expr_to_series(self, expr, df) -> pd.Series:
        """将表达式转换为pd.Series(支持列引用,字面量,数学表达式,字符串函数)"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name in df.columns:
                return df[col_name]
            # 表限定符
            table_part = expr.table if hasattr(expr, 'table') and expr.table else None
            qualified = f"{table_part}.{col_name}" if table_part else None
            if qualified and qualified in df.columns:
                return df[qualified]
            raise ValueError(f"列 '{qualified or col_name}' 不存在")
        elif isinstance(expr, exp.Literal):
            val = expr.this
            return pd.Series([val] * len(df), index=df.index)
        elif isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div)):
            return self._evaluate_math_expression(expr, df)
        elif isinstance(expr, exp.Coalesce):
            return self._evaluate_coalesce_vectorized(expr, df)
        elif self._is_string_function(expr):
            return self._evaluate_string_function(expr, df)
        elif isinstance(expr, exp.Case):
            return self._evaluate_case_expression(expr, df)
        else:
            raise ValueError(f"不支持的表达式类型: {type(expr).__name__}")

    def _literal_value(self, expr) -> Any:
        """提取字面量值"""
        if isinstance(expr, exp.Literal):
            return expr.this
        elif isinstance(expr, exp.Column):
            return expr.name
        return str(expr)

    def _get_arg(self, expr, arg_name, default=None, type_fn=None):
        """从表达式参数中提取值,支持类型转换

        消除字符串函数中重复的 _literal_value(expr.args.get(...)) 模式.
        例如: int(self._literal_value(expr.args.get('n'))) if expr.args.get('n') else 1
        简化为: self._get_arg(expr, 'n', 1, int)
        """
        arg = expr.args.get(arg_name)
        if arg is None:
            return default
        val = self._literal_value(arg)
        if type_fn and val is not None:
            return type_fn(val)
        return val

    def _get_from_table(self, parsed_sql: exp.Expression) -> Tuple[str, Optional[exp.Expression]]:
        """获取FROM子句中的表名.

        Returns:
            (table_name, subquery_expr): 
                - 普通表: (表名, None)
                - FROM子查询: (别名, Subquery/Select表达式)
        """
        from_clause = parsed_sql.args.get('from')
        if not from_clause:
            # 尝试使用 from_ 键(sqlglot的另一种存储方式)
            from_clause = parsed_sql.args.get('from_')
        if from_clause:
            # 检查FROM子句是否是子查询(FROM (SELECT ...) AS alias)
            if hasattr(from_clause, 'this') and isinstance(from_clause.this, (exp.Subquery, exp.Select)):
                subquery_node = from_clause.this
                # 获取别名
                alias = getattr(subquery_node, 'alias', None)
                if not alias and isinstance(subquery_node, exp.Subquery):
                    # sqlglot存储别名在alias属性
                    alias = subquery_node.alias
                if not alias:
                    alias = "_subquery"
                return (alias, subquery_node)
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'name'):
                return (from_clause.this.name, None)
            # 兼容 Table 对象
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'this'):
                return (from_clause.this.this, None)

        # 如果没有明确的FROM子句,返回第一个表名
        raise ValueError("无法确定FROM子句中的表名")

    def _apply_join_clause(self, joins, left_df, worksheets_data=None, left_table=None) -> pd.DataFrame:
        """
        应用JOIN子句,支持INNER/LEFT/RIGHT/FULL/CROSS JOIN
        性能优化:使用索引优化和智能JOIN策略

        Args:
            joins: sqlglot joins列表
            left_df: 左表DataFrame
            worksheets_data: 所有工作表数据
            left_table: 左表名

        Returns:
            pd.DataFrame: JOIN后的DataFrame
        """
        if isinstance(joins, exp.Join):
            joins = [joins]

        # 性能优化:检查数据大小,决定是否使用索引优化
        left_size = len(left_df)
        total_memory_mb = left_size * len(left_df.columns) * 8 / (1024 * 1024)  # 估算内存使用
        
        # 初始化JOIN列映射(用于别名解析)
        self._join_column_mapping = {}
        
        result_df = left_df

        for join in joins:
            # 解析JOIN类型
            join_side = str(join.side).upper() if join.side else None
            join_kind_name = str(join.kind).upper() if join.kind else None
            join_kind = self._JOIN_KIND_MAP.get((join_side, join_kind_name), 'inner')

            # 解析右表
            right_table_expr = join.this
            if hasattr(right_table_expr, 'this'):
                right_table = right_table_expr.this if isinstance(right_table_expr.this, str) else (
                    right_table_expr.this.name if hasattr(right_table_expr.this, 'name') else str(right_table_expr.this)
                )
            else:
                right_table = right_table_expr.name if hasattr(right_table_expr, 'name') else str(right_table_expr)

            # 检查右表是否有别名
            right_alias = right_table
            # 优先使用 Table.alias 属性(sqlglot 中 Table 的 alias 返回字符串)
            if hasattr(right_table_expr, 'alias') and right_table_expr.alias:
                table_alias = right_table_expr.alias
                if isinstance(table_alias, str) and table_alias != right_table:
                    right_alias = table_alias
                elif hasattr(table_alias, 'alias'):
                    right_alias = table_alias.alias
            # 备用:遍历 TableAlias 节点
            if right_alias == right_table:
                for alias in join.find_all(exp.TableAlias):
                    parent = alias.this
                    parent_name = parent.name if hasattr(parent, 'name') else str(parent)
                    if parent_name == right_table or str(parent) == right_table:
                        right_alias = alias.alias
                        break

            # 记录别名映射
            self._table_aliases[right_alias] = right_table
            self._table_aliases[right_table] = right_table

            if right_table not in worksheets_data:
                available = list(worksheets_data.keys())
                raise StructuredSQLError(
                    "table_not_found",
                    f"JOIN表 '{right_table}' 不存在.可用表: {available}",
                    hint="请检查JOIN的表名,或用excel_list_sheets查看可用工作表名.",
                    context={"table_requested": right_table, "available_tables": available}
                )

            right_df = worksheets_data[right_table].copy()

            # 解析ON条件(CROSS JOIN不需要ON)
            on_clause = join.args.get('on')
            left_on_col = None
            right_on_col = None
            actual_right_on = None
            
            # 性能优化:为大数据集创建索引(如果JOIN列存在且数据量大)
            if on_clause and total_memory_mb > 10:  # 大于10MB的数据集使用索引优化
                # 提前解析ON条件来决定是否需要索引
                left_on_col, right_on_col, _non_equi = self._parse_join_on_condition(on_clause, left_table, right_table, right_alias)
                
                # 非等值连接不需要索引优化
                if not _non_equi and left_on_col:
                    if left_on_col in result_df.columns:
                        result_df = result_df.set_index(left_on_col, inplace=False)
                    if right_on_col and right_on_col in right_df.columns:
                        right_df = right_df.set_index(right_on_col, inplace=False)

            if join_kind == 'cross':
                # CROSS JOIN: 笛卡尔积,不需要ON条件
                pass
            elif not on_clause:
                raise StructuredSQLError(
                    "join_error",
                    "JOIN缺少ON条件",
                    hint="JOIN必须包含ON条件,例如:... JOIN 表2 ON 表1.id = 表2.id."
                )
            else:
                left_on_col, right_on_col, non_equi_cond = self._parse_join_on_condition(on_clause, left_table, right_table, right_alias)
                # 等值连接:验证列存在
                if not non_equi_cond and left_on_col and left_on_col not in result_df.columns:
                    raise StructuredSQLError(
                        "column_not_found",
                        f"左表 '{left_table}' 没有列 '{left_on_col}'.可用列: {list(result_df.columns)}",
                        hint="请检查ON条件中左表的列名拼写.",
                        context={"table": left_table, "column_requested": left_on_col, "available_columns": list(result_df.columns)}
                    )
                if right_on_col and right_on_col not in right_df.columns:
                    raise StructuredSQLError(
                        "column_not_found",
                        f"右表 '{right_table}' 没有列 '{right_on_col}'.可用列: {list(right_df.columns)}",
                        hint="请检查ON条件中右表的列名拼写.",
                        context={"table": right_table, "column_requested": right_on_col, "available_columns": list(right_df.columns)}
                    )

            # 执行JOIN
            # 为右表列添加别名前缀避免冲突
            right_df_renamed = right_df.copy()
            col_mapping = {}
            for col in right_df_renamed.columns:
                if col in result_df.columns and (left_on_col is None or col != left_on_col):
                    new_col = f"{right_alias}.{col}"
                    col_mapping[col] = new_col
                elif left_on_col and col == left_on_col:
                    # ON列(无论左右列名是否相同):右表列重命名避免_x/_y后缀
                    new_col = f"{right_alias}.{col}"
                    col_mapping[col] = new_col
            right_df_renamed = right_df_renamed.rename(columns=col_mapping)
            
            # 更新JOIN列映射,用于别名解析
            if right_alias not in self._join_column_mapping:
                self._join_column_mapping[right_alias] = {}
            for orig_col, new_col in col_mapping.items():
                self._join_column_mapping[right_alias][orig_col] = new_col

            # 调整右表ON列名(如果被重命名了)
            if right_on_col:
                actual_right_on = col_mapping.get(right_on_col, right_on_col)

            # 合并双行表头描述
            if right_table in self._header_descriptions:
                for orig_col, new_col in col_mapping.items():
                    if orig_col in self._header_descriptions[right_table]:
                        self._header_descriptions[right_table][new_col] = self._header_descriptions[right_table][orig_col]

            if join_kind == 'cross':
                # CROSS JOIN: 笛卡尔积(无需ON列)
                # 先临时移除冲突列名,合并后再恢复
                temp_col_mapping = {}
                right_df_for_cross = right_df_renamed.copy()
                
                for col in right_df_for_cross.columns:
                    if col in result_df.columns:
                        temp_col = f"{right_alias}_temp_{col}"
                        temp_col_mapping[col] = temp_col
                        right_df_for_cross = right_df_for_cross.rename(columns={col: temp_col})
                
                result_df = result_df.merge(right_df_for_cross, how='cross')
                
                # 恢复原始列名
                for old_col, new_col in temp_col_mapping.items():
                    result_df = result_df.rename(columns={new_col: old_col})
            elif non_equi_cond is not None:
                # 非等值连接: cross join + row filter
                result_df = result_df.merge(right_df_renamed, how='cross')
                result_df = self._apply_row_filter(non_equi_cond, result_df)
            else:
                result_df = result_df.merge(
                    right_df_renamed,
                    left_on=left_on_col,
                    right_on=actual_right_on,
                    how=join_kind
                )

            # 合并后删除重复的ON列(右表侧)
            if actual_right_on and actual_right_on in result_df.columns and actual_right_on != left_on_col:
                result_df = result_df.drop(columns=[actual_right_on])

        return result_df

    def _parse_join_on_condition(self, on_clause, left_table: str, right_table: str, right_alias: str):
        """
        解析JOIN ON条件

        等值连接返回 (left_col, right_col, None)
        非等值连接返回 (None, None, on_clause)
        """
        # 非等值连接: 返回条件用于cross+filter
        if isinstance(on_clause, (exp.GT, exp.GTE, exp.LT, exp.LTE, exp.NEQ)):
            return (None, None, on_clause)

        if isinstance(on_clause, exp.EQ):
            left_expr = on_clause.left
            right_expr = on_clause.right
        elif isinstance(on_clause, exp.And):
            # 多条件JOIN暂不支持,取第一个等值条件
            for child in on_clause.find_all(exp.EQ):
                left_expr = child.left
                right_expr = child.right
                break
            else:
                raise ValueError("JOIN ON多条件暂不支持,请使用单个等值连接条件")
        else:
            raise ValueError(f"JOIN ON条件格式不支持,请使用等值连接: ON a.id = b.id")

        def resolve_column(col_expr) -> str:
            """解析列表达式,返回列名和表名.

            Args:
                col_expr: SQLGlot列表达式

            Returns:
                tuple: (列名, 表名或None)
            """
            if isinstance(col_expr, exp.Column):
                col_name = col_expr.name
                # 检查是否有表限定符
                table_part = col_expr.table if hasattr(col_expr, 'table') and col_expr.table else None
                return col_name, table_part
            return str(col_expr), None

        left_col, left_tbl = resolve_column(left_expr)
        right_col, right_tbl = resolve_column(right_expr)

        # 判断哪个属于左表,哪个属于右表
        if left_tbl:
            resolved_left_tbl = self._table_aliases.get(left_tbl, left_tbl)
            if resolved_left_tbl == right_table or left_tbl == right_alias:
                # 左表达式实际指向右表,交换
                return right_col, left_col, None
        if right_tbl:
            resolved_right_tbl = self._table_aliases.get(right_tbl, right_tbl)
            if resolved_right_tbl == left_table:
                # 右表达式实际指向左表,交换
                return right_col, left_col, None

        # 无表限定符:根据列是否存在于左右表来判断
        # 默认左=左表, 右=右表
        return left_col, right_col, None

    def _apply_where_clause(self, parsed_sql: exp.Expression, df) -> pd.DataFrame:
        """应用WHERE条件"""
        where_clause = parsed_sql.args.get('where')
        if not where_clause:
            return df

        # 如果WHERE包含复杂表达式(pandas query不支持的类型),直接使用逐行过滤
        where_expr = where_clause.this
        has_complex = any(where_expr.find(t) is not None for t in self._COMPLEX_EXPR_TYPES)

        if has_complex:
            return self._apply_row_filter(where_expr, df)

        # 将SQLGlot表达式转换为pandas查询条件
        condition_str = self._sql_condition_to_pandas(where_expr, df)

        if condition_str:
            try:
                return df.query(condition_str)
            except Exception as e:
                # 如果查询失败,尝试逐行过滤
                return self._apply_row_filter(where_clause.this, df)

        logger.warning("WHERE条件转换为pandas表达式失败,回退到逐行过滤: %s", where_expr)
        return self._apply_row_filter(where_expr, df)

    @staticmethod
    def _like_to_regex(value_str: str) -> str:
        """将SQL LIKE模式转换为pandas regex模式(%->.*  _->.)"""
        pattern = str(value_str).strip("'\"")
        return pattern.replace('%', '.*').replace('_', '.')

    @staticmethod
    def _parse_literal_value(expr: exp.Literal) -> Any:
        """将SQL Literal解析为Python值(字符串->str,数字->int/float)"""
        if expr.is_string:
            return expr.this
        try:
            return float(expr.this) if '.' in str(expr.this) else int(expr.this)
        except (ValueError, TypeError):
            return expr.this

    def _in_to_pandas(self, in_expr: exp.In, df, negate: bool = False) -> str:
        """将IN/NOT IN条件转换为pandas表达式(支持子查询和值列表)

        Args:
            in_expr: sqlglot In表达式
            df: 当前DataFrame
            negate: True->NOT IN(~isin),False->IN(isin)
        """
        left = self._expression_to_column_reference(in_expr.this, df)
        prefix = "~" if negate else ""

        # 子查询模式 (IN (SELECT ...))
        subquery = in_expr.args.get('query')
        if subquery and isinstance(subquery, exp.Subquery):
            try:
                sub_result = self._execute_subquery(subquery, self._current_worksheets)
                if len(sub_result.columns) > 0:
                    # 确保子查询结果排除表头行（iloc[1:, 0]而不是iloc[:, 0]）
                    sub_values = sub_result.iloc[1:, 0].dropna().tolist()
                    values_str = ', '.join(repr(v) for v in sub_values)
                    return f"{prefix}{left}.isin([{values_str}])"
                return f"{prefix}{left}.isin([])"
            except Exception as e:
                op = "NOT IN" if negate else "IN"
                raise ValueError(f"{op}子查询执行失败: {e}")

        # 值列表模式
        values = [self._expression_to_value(v, df) for v in in_expr.expressions]
        values_str = ', '.join(str(v) for v in values)
        return f"{prefix}{left}.isin([{values_str}])"

    def _sql_condition_to_pandas(self, condition: exp.Expression, df) -> str:
        """将SQL条件转换为pandas查询字符串"""
        op_type = type(condition)
        if op_type in self._PANDAS_OPS:
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} {self._PANDAS_OPS[op_type]} {right}"

        elif isinstance(condition, exp.And):
            left = self._sql_condition_to_pandas(condition.left, df)
            right = self._sql_condition_to_pandas(condition.right, df)
            if left is None or right is None:
                raise ValueError("AND条件包含不支持的子查询类型,需使用逐行过滤")
            return f"({left}) & ({right})"

        elif isinstance(condition, exp.Or):
            left = self._sql_condition_to_pandas(condition.left, df)
            right = self._sql_condition_to_pandas(condition.right, df)
            if left is None or right is None:
                raise ValueError("OR条件包含不支持的子查询类型,需使用逐行过滤")
            return f"({left}) | ({right})"

        elif isinstance(condition, exp.Paren):
            return self._sql_condition_to_pandas(condition.this, df)

        elif isinstance(condition, exp.Not):
            inner = condition.this
            if isinstance(inner, exp.Like):
                left = self._expression_to_column_reference(inner.this, df)
                right = self._expression_to_value(inner.expression, df)
                regex = self._like_to_regex(right)
                return f"~{left}.str.match('{regex}', case=False, na=False)"
            if isinstance(inner, exp.In):
                return self._in_to_pandas(inner, df, negate=True)
            # 其他NOT表达式(IS NOT NULL等)
            pandas_expr = self._sql_condition_to_pandas(inner, df)
            return f"~({pandas_expr})"

        elif isinstance(condition, exp.Like):
            left = self._expression_to_column_reference(condition.this, df)
            right = self._expression_to_value(condition.expression, df)
            regex = self._like_to_regex(right)
            return f"{left}.str.match('{regex}', case=False, na=False)"

        elif isinstance(condition, exp.In):
            return self._in_to_pandas(condition, df, negate=False)

        # EXISTS (子查询) -- 需要逐行评估,返回None触发行过滤回退
        elif isinstance(condition, exp.Exists):
            return None

        # IS NULL (sqlglot解析为 exp.Is)
        elif isinstance(condition, exp.Is):
            left = self._expression_to_column_reference(condition.this, df)
            return f"{left}.isna()"

        # BETWEEN x AND y
        elif isinstance(condition, exp.Between):
            left = self._expression_to_column_reference(condition.this, df)
            low = self._expression_to_value(condition.args['low'], df)
            high = self._expression_to_value(condition.args['high'], df)
            return f"({left} >= {low}) & ({left} <= {high})"

        else:
            raise ValueError(f"不支持的条件类型: {type(condition)}")

    def _expression_to_column_reference(self, expr: exp.Expression, df) -> str:
        """将表达式转换为列引用(支持表限定符 a.column)"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            # 处理表限定符 (a.column_name -> 查找 "a.column_name" 或 "column_name")
            table_part = expr.table if hasattr(expr, 'table') and expr.table else None
            qualified = f"{table_part}.{col_name}" if table_part else None

            if qualified and qualified in df.columns:
                return f"`{qualified}`"
            if col_name in df.columns:
                return f"`{col_name}`"
            
            # JOIN别名映射支持
            if table_part:
                # 1. 检查用户使用的别名格式是否直接存在
                alias_col = f"{table_part}.{col_name}"
                if alias_col in df.columns:
                    return f"`{alias_col}`"
                
                # 2. 检查JOIN后pandas添加的后缀格式(table_part_列名)
                pandas_suffix_col = f"{table_part}_{col_name}"
                if pandas_suffix_col in df.columns:
                    return f"`{pandas_suffix_col}`"
                
                # 2.1 检查pandas merge后的_x/_y后缀映射
                # 2.1.1 检查列名是否直接是_x/_y后缀格式(pandas自动处理)
                for col in df.columns:
                    if col.endswith('_x') and col[:-2] == col_name:
                        # 检查是否是JOIN冲突导致的_x后缀
                        if table_part:
                            # 如果table_part存在,检查是否有对应的x_col
                            x_col = f"{table_part}_x"
                            if x_col in df.columns:
                                return f"`{x_col}`"
                            # 如果没有,检查是否是冲突导致的直接_x后缀
                            return f"`{col}`"
                    elif col.endswith('_y') and col[:-2] == col_name:
                        if table_part:
                            # 如果table_part存在,检查是否有对应的y_col
                            y_col = f"{table_part}_y"
                            if y_col in df.columns:
                                return f"`{y_col}`"
                            # 如果没有,检查是否是冲突导致的直接_y后缀
                            return f"`{col}`"
                
                # 2.1.2 检查表名+_x/_y后缀格式(更智能的匹配)
                # 处理table_part_x和table_part_y格式
                table_part_x = f"{table_part}_x"
                table_part_y = f"{table_part}_y"
                if table_part_x in df.columns:
                    return f"`{table_part_x}`"
                if table_part_y in df.columns:
                    return f"`{table_part_y}`"
                
                # 3. 如果有JOIN映射,检查映射后的列名
                if hasattr(self, '_join_column_mapping'):
                    mapped_col = self._join_column_mapping.get(table_part, {}).get(col_name)
                    if mapped_col and mapped_col in df.columns:
                        return f"`{mapped_col}`"
                
                # 4. 尝试其他可能的别名格式
                if hasattr(self, '_table_aliases'):
                    resolved_table = self._table_aliases.get(table_part, table_part)
                    # 检查原始表名+列名
                    original_col = f"{resolved_table}_{col_name}"
                    if original_col in df.columns:
                        return f"`{original_col}`"
                    # 如果用户使用的是原始表名,检查原始表名+列名
                    if table_part == resolved_table and f"{table_part}_{col_name}" in df.columns:
                        return f"`{table_part}_{col_name}`"
                
                # 5. 最后尝试原始列名
                if col_name in df.columns:
                    return f"`{col_name}`"
            
            suggestion = self._suggest_column_name(col_name, list(df.columns))
            raise ValueError(f"列 '{qualified or col_name}' 不存在.可用列: {list(df.columns)}.{suggestion}")

        elif isinstance(expr, exp.Literal):
            return str(expr.this)

        elif isinstance(expr, exp.AggFunc):
            # 对于HAVING子句中的聚合函数,需要查找对应的列
            func_name = type(expr).__name__.lower()

            # 优先:通过SELECT别名映射(HAVING COUNT(*) -> SELECT COUNT(*) as cnt)
            if hasattr(self, '_having_agg_alias_map') and self._having_agg_alias_map:
                agg_sql = expr.sql()
                for map_sql, alias in self._having_agg_alias_map.items():
                    if agg_sql == map_sql:
                        if alias in df.columns:
                            return f"`{alias}`"
                        break

            # 模糊匹配:查找列名包含函数名的列(count->count_star/avg_dmg等)
            for col in df.columns:
                if func_name in col.lower():
                    return f"`{col}`"
            
            # 单数值列兜底(常见于全表聚合场景)
            numeric_cols = [
                col for col in df.columns
                if pd.to_numeric(df[col], errors='coerce').notna().sum() > 0
            ]
            if len(numeric_cols) == 1:
                return f"`{numeric_cols[0]}`"
            if df.columns.size > 0:
                return f"`{df.columns[0]}`"

            raise ValueError(f"无法找到聚合函数 {func_name} 对应的列.可用列: {list(df.columns)}")

        else:
            raise ValueError(f"不支持的表达式类型: {type(expr)}")

    def _expression_to_value(self, expr: exp.Expression, df) -> Union[str, int, float]:
        """将表达式转换为值"""
        if isinstance(expr, exp.Literal):
            # 委托_parse_literal_value统一处理Literal->Python值转换
            parsed = self._parse_literal_value(expr)
            if isinstance(parsed, str):
                return f"'{parsed}'"
            return parsed

        elif isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name not in df.columns:
                suggestion = self._suggest_column_name(col_name, list(df.columns))
                raise ValueError(f"列 '{col_name}' 不存在.可用列: {list(df.columns)}.{suggestion}")
            return f"`{col_name}`"

        elif isinstance(expr, exp.AggFunc):
            # 聚合函数作为值的处理(HAVING子句中)
            # 使用与_expression_to_column_reference相同的逻辑
            return self._expression_to_column_reference(expr, df)

        elif isinstance(expr, exp.Subquery):
            """标量子查询: WHERE col > (SELECT AVG(...) FROM ...)
            
            子查询结果应返回单行单列的标量值.
            修复: 直接从 DataFrame 提取标量值,不再假设有标题行.
            """
            try:
                sub_result = self._execute_subquery(expr, self._current_worksheets)
                if len(sub_result) > 0 and len(sub_result.columns) > 0:
                    scalar_val = sub_result.iloc[0, 0]
                    if isinstance(scalar_val, (int, float, np.integer, np.floating)):
                        return float(scalar_val)
                    if scalar_val is not None:
                        return f"'{scalar_val}'"
                return "0"
            except Exception as e:
                raise ValueError(f"标量子查询执行失败: {e}")

        else:
            raise ValueError(f"不支持的表达式类型: {type(expr)}")

    def _apply_row_filter(self, condition: exp.Expression, df) -> pd.DataFrame:
        """逐行应用过滤条件(备用方案),使用apply替代iterrows提升性能"""
        mask = df.apply(lambda row: self._evaluate_condition_for_row(condition, row), axis=1)
        return df[mask]

    def _evaluate_condition_for_row(self, condition: exp.Expression, row: pd.Series) -> bool:
        """为单行评估条件"""
        try:
            op_type = type(condition)
            if op_type in self._COMPARISON_OPS:
                left_val = self._get_row_value(condition.left, row)
                right_val = self._get_row_value(condition.right, row)
                try:
                    return self._COMPARISON_OPS[op_type](left_val, right_val)
                except (TypeError, ValueError):
                    return False

            elif isinstance(condition, exp.And):
                return (self._evaluate_condition_for_row(condition.left, row) and
                       self._evaluate_condition_for_row(condition.right, row))

            elif isinstance(condition, exp.Or):
                return (self._evaluate_condition_for_row(condition.left, row) or
                       self._evaluate_condition_for_row(condition.right, row))

            elif isinstance(condition, exp.Is):
                val = self._get_row_value(condition.this, row)
                return pd.isna(val) or val is None

            elif isinstance(condition, exp.Not):
                inner = condition.this
                if isinstance(inner, exp.Is):
                    val = self._get_row_value(inner.this, row)
                    return not (pd.isna(val) or val is None)
                return not self._evaluate_condition_for_row(inner, row)

            elif isinstance(condition, exp.Like):
                val = str(self._get_row_value(condition.this, row) or '')
                pattern = str(self._get_row_value(condition.expression, row) or '')
                regex = self._like_to_regex(pattern)
                return bool(re.match(regex, val, re.IGNORECASE))

            elif isinstance(condition, exp.In):
                val = self._get_row_value(condition.this, row)
                values = [self._get_row_value(e, row) for e in condition.expressions]
                return val in values

            elif isinstance(condition, exp.Between):
                val = self._get_row_value(condition.this, row)
                low = self._get_row_value(condition.args['low'], row)
                high = self._get_row_value(condition.args['high'], row)
                try:
                    return float(low) <= float(val) <= float(high)
                except (TypeError, ValueError):
                    return False

            elif isinstance(condition, exp.Exists):
                return self._evaluate_exists_for_row(condition, row)

            # 其他条件类型...

            return True

        except Exception:
            return False

    def _evaluate_exists_for_row(self, condition: exp.Exists, row: pd.Series) -> bool:
        """评估EXISTS子查询(支持关联子查询)"""
        inner_expr = condition.this
        if isinstance(inner_expr, exp.Subquery):
            inner_select = inner_expr.this
        elif isinstance(inner_expr, exp.Select):
            inner_select = inner_expr
        else:
            return False

        if not (hasattr(self, '_current_worksheets') and self._current_worksheets):
            return False

        inner_from, _ = self._get_from_table(inner_select)
        has_correlation = False
        for col in inner_select.find_all(exp.Column):
            col_name = col.name
            table_part = col.table if hasattr(col, 'table') and col.table else None
            if table_part:
                resolved = self._table_aliases.get(table_part, table_part)
                if resolved != inner_from:
                    has_correlation = True
                    break
            else:
                if inner_from in self._current_worksheets:
                    for tbl_name, tbl_df in self._current_worksheets.items():
                        if tbl_name != inner_from and col_name in tbl_df.columns:
                            has_correlation = True
                            break

        if has_correlation:
            return self._evaluate_correlated_exists(inner_select, inner_from, row)
        else:
            sub_result = self._execute_subquery(inner_expr, self._current_worksheets)
            return len(sub_result) > 0

    def _evaluate_correlated_exists(self, inner_select, inner_from: str, row: pd.Series) -> bool:
        """评估关联EXISTS子查询:替换外部引用为当前行值后执行"""
        inner_sql = str(inner_select)
        inner_from_cols = set()
        if inner_from in self._current_worksheets:
            inner_from_cols = set(self._current_worksheets[inner_from].columns)

        for col in inner_select.find_all(exp.Column):
            col_name = col.name
            table_part = col.table if hasattr(col, 'table') and col.table else None
            should_substitute = False

            if table_part:
                resolved = self._table_aliases.get(table_part, table_part)
                for tbl_name in self._current_worksheets:
                    if tbl_name != inner_from and (resolved == tbl_name or table_part == tbl_name):
                        should_substitute = True
                        break
            else:
                for tbl_name, tbl_df in self._current_worksheets.items():
                    if tbl_name != inner_from and col_name in tbl_df.columns and col_name not in inner_from_cols:
                        should_substitute = True
                        break

            if should_substitute:
                val = row.get(col_name)
                if val is not None:
                    if table_part:
                        inner_sql = inner_sql.replace(f"{table_part}.{col_name}", repr(val), 1)
                    else:
                        pattern = r'\b' + re.escape(col_name) + r'\b'
                        inner_sql = re.sub(pattern, repr(val), inner_sql, count=1)

        try:
            parsed_inner = sqlglot.parse_one(inner_sql)
            # 保存外部查询状态,避免_execute_query内部重置覆盖
            saved_aliases = dict(self._table_aliases)
            saved_worksheets = dict(self._current_worksheets)
            try:
                sub_result = self._execute_query(parsed_inner, saved_worksheets)
                return len(sub_result) > 0
            finally:
                self._table_aliases = saved_aliases
                self._current_worksheets = saved_worksheets
        except Exception:
            return False

    def _extract_column_references(self, expr: exp.Expression) -> List[str]:
        """从表达式中提取所有列引用

        Args:
            expr: SQL表达式对象

        Returns:
            列名列表

        修复说明(REQ-061):
        - 支持从复杂表达式(CASE WHEN/COALESCE/函数调用)中提取列引用
        - 递归遍历表达式树,提取所有Column类型的子节点
        """
        columns = []
        if isinstance(expr, exp.Column):
            columns.append(expr.name)
        elif hasattr(expr, 'expressions'):
            for sub_expr in expr.expressions:
                columns.extend(self._extract_column_references(sub_expr))
        elif hasattr(expr, 'this'):
            columns.extend(self._extract_column_references(expr.this))
        if hasattr(expr, 'args'):
            for arg in expr.args.values():
                if isinstance(arg, exp.Expression):
                    columns.extend(self._extract_column_references(arg))
        return columns

    def _get_row_value(self, expr: exp.Expression, row: pd.Series) -> Any:
        """获取行中表达式的值"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            return row.get(col_name)

        elif isinstance(expr, exp.Literal):
            return self._parse_literal_value(expr)

        elif isinstance(expr, exp.Coalesce):
            return self._evaluate_coalesce_for_row(expr, row)

        elif isinstance(expr, exp.Case):
            return self._evaluate_case_expression(expr, None, row=row)

        elif isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div)):
            return self._evaluate_math_for_row(expr, row)

        elif self._is_string_function(expr):
            return self._evaluate_string_function_for_row(expr, row)

        else:
            return None

    def _apply_group_by_aggregation(self, parsed_sql: exp.Expression, df) -> pd.DataFrame:
        """应用GROUP BY和聚合函数
        
        Args:
            parsed_sql: SQL解析后的表达式对象
            df: 要处理的DataFrame数据
        
        Returns:
            应用GROUP BY和聚合函数后的DataFrame
        
        性能优化:
        - 大数据集使用向量化操作代替逐行计算
        - 减少不必要的DataFrame复制
        - 优化groupby操作使用observed=True避免稀疏分组
        
        修复说明(REQ-061):
        - 修复多列GROUP BY逻辑，确保当GROUP BY包含多列且SELECT包含计算表达式时，
          正确按所有GROUP BY列分组，而非仅按计算列分组
        - 确保所有GROUP BY列都包含在最终结果中
        """
        group_by_columns = []
        group_clause = parsed_sql.args.get('group')
        if group_clause:
            for group_expr in group_clause.expressions:
                if isinstance(group_expr, exp.Column):
                    group_by_columns.append(group_expr.name)
                else:
                    # 修复(REQ-061): 从复杂表达式中提取列引用
                    # 支持GROUP BY包含表达式(如CASE WHEN/COALESCE)的情况
                    column_refs = self._extract_column_references(group_expr)
                    for col_name in column_refs:
                        if col_name not in group_by_columns:
                            group_by_columns.append(col_name)

        # 检查是否有聚合函数
        aggregations = {}
        select_exprs = {}

        for i, select_expr in enumerate(parsed_sql.expressions):
            alias_name, original_expr = self._extract_select_alias(select_expr, i)
            select_exprs[alias_name] = original_expr

        # 检查聚合函数
        for alias_name, expr in select_exprs.items():
            if self._is_aggregate_function(expr):
                aggregations[alias_name] = expr
            elif hasattr(expr, 'name') and expr.name not in group_by_columns:
                # 如果是非聚合列且不在GROUP BY中,需要添加到GROUP BY
                # 跳过窗口函数表达式(窗口函数在GROUP BY之后处理)
                if isinstance(expr, exp.Window):
                    continue
                if isinstance(expr, (exp.Column, exp.Identifier)):
                    group_by_columns.append(expr.name)
                else:
                    # 修复(REQ-061): 从复杂表达式中提取列引用
                    # 支持SELECT包含表达式(如CASE WHEN/COALESCE)的情况
                    # 跳过窗口函数表达式
                    if isinstance(expr, exp.Window):
                        continue
                    column_refs = self._extract_column_references(expr)
                    for col_name in column_refs:
                        if col_name not in group_by_columns:
                            group_by_columns.append(col_name)

        # 保存GROUP BY列到实例变量,供_build_total_row使用
        self._group_by_columns = group_by_columns

        if not aggregations:
            # 没有聚合函数,只应用GROUP BY去重
            if group_by_columns:
                # 性能优化:使用drop_duplicates的subset参数避免全列比较
                return df[group_by_columns].drop_duplicates(subset=group_by_columns).reset_index(drop=True)
            else:
                return df

        # 预计算CASE WHEN/COALESCE/标量子查询表达式,添加到df副本,使grouped可访问
        df = df.copy()
        for alias_name, expr in select_exprs.items():
            if isinstance(expr, exp.Case) and alias_name not in df.columns:
                df[alias_name] = self._evaluate_case_expression(expr, df)
            elif isinstance(expr, exp.Coalesce) and alias_name not in df.columns:
                df[alias_name] = self._evaluate_coalesce_vectorized(expr, df)
            elif isinstance(expr, exp.Subquery) and alias_name not in df.columns:
                try:
                    sub_result = self._execute_subquery(expr, self._current_worksheets)
                    if len(sub_result) > 0 and len(sub_result.columns) > 0:
                        scalar_val = sub_result.iloc[0, 0]
                        df[alias_name] = pd.Series([scalar_val] * len(df), index=df.index)
                    else:
                        df[alias_name] = pd.Series([None] * len(df), index=df.index)
                except Exception:
                    df[alias_name] = pd.Series([None] * len(df), index=df.index)

        # 应用聚合
        # 性能优化:使用observed=True减少分组计算开销
        if group_by_columns:
            # 确保group_by_columns中的列都存在
            valid_group_cols = [c for c in group_by_columns if c in df.columns]
            if valid_group_cols:
                grouped = df.groupby(valid_group_cols, observed=True)
            else:
                grouped = df.groupby(lambda x: 0)
        else:
            # 全表聚合
            grouped = df.groupby(lambda x: 0)  # 将所有行分组为一组

        # 按照SQL SELECT表达式的顺序构建结果
        result_data = {}
        ordered_columns = []

        # 按SELECT表达式顺序处理列
        for i, select_expr in enumerate(parsed_sql.expressions):
            alias_name, original_expr = self._extract_select_alias(select_expr, i)

            # 跳过SELECT *，它会在循环后单独处理
            if isinstance(original_expr, exp.Star):
                continue

            ordered_columns.append(alias_name)

            # 处理聚合函数
            is_agg = self._is_aggregate_function(select_expr if not isinstance(select_expr, exp.Alias) else select_expr.this)
            
            if is_agg:
                agg_expr = original_expr if isinstance(select_expr, exp.Alias) else select_expr
                agg_result = self._apply_aggregation_function(agg_expr, grouped, df)
                # 修复:确保聚合结果是正确格式
                if isinstance(agg_result, (int, float, np.integer, np.floating)):
                    result_data[alias_name] = pd.Series([agg_result])
                elif isinstance(agg_result, pd.Series):
                    result_data[alias_name] = agg_result.reset_index(drop=True)
                else:
                    # 尝试转换为Series
                    try:
                        result_data[alias_name] = pd.Series([agg_result])
                    except:
                        result_data[alias_name] = pd.Series([None])
            # 处理CASE WHEN表达式(已预计算到df,直接从grouped取first)
            elif isinstance(original_expr, exp.Case):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理COALESCE表达式(已预计算到df,直接从grouped取first)
            elif isinstance(original_expr, exp.Coalesce):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理标量子查询(已预计算到df,直接从grouped取first)
            elif isinstance(original_expr, exp.Subquery):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理普通列(GROUP BY列)
            elif hasattr(original_expr, 'name'):
                col_name = original_expr.name
                if col_name in group_by_columns:
                    result_data[alias_name] = grouped[col_name].first().reset_index(drop=True)
            
        # 处理SELECT *的情况：添加所有GROUP BY列
        has_star = any(isinstance(self._extract_select_alias(expr, 0)[1], exp.Star) 
                       for expr in parsed_sql.expressions)
        if has_star:
            for col in group_by_columns:
                if col not in result_data:
                    result_data[col] = grouped[col].first().reset_index(drop=True)
                    if col not in ordered_columns:
                        ordered_columns.append(col)

        # 确保所有GROUP BY列都在结果中(修复多列GROUP BY逻辑)
        # 如果SELECT中不包含某些GROUP BY列,需要确保它们被包含
        for col in group_by_columns:
            if col not in result_data:
                # 为缺失的GROUP BY列添加到结果
                result_data[col] = grouped[col].first().reset_index(drop=True)
                # 如果不在有序列列表中,添加到末尾
                if col not in ordered_columns:
                    ordered_columns.append(col)

        # 组合结果,保持列顺序
        try:
            result_df = pd.DataFrame(result_data, columns=ordered_columns).reset_index(drop=True)
        except Exception as e:
            # 如果创建DataFrame失败,尝试逐列构建
            print(f"警告:构建结果DataFrame时出错: {e}")
            result_data = {k: v.reset_index(drop=True) for k, v in result_data.items()}
            result_df = pd.DataFrame(result_data, columns=ordered_columns).reset_index(drop=True)

        return result_df

    def _is_aggregate_function(self, expr: exp.Expression) -> bool:
        """检查是否为聚合函数"""
        if isinstance(expr, exp.AggFunc):
            return True
        elif isinstance(expr, exp.Alias):
            return self._is_aggregate_function(expr.this)
        return False

    def _execute_subquery(
        self,
        subquery_expr,
        worksheets_data: Dict[str, pd.DataFrame]
    ) -> pd.DataFrame:
        """
        执行子查询,返回结果DataFrame

        Args:
            subquery_expr: sqlglot Subquery或Select表达式
            worksheets_data: 当前可用的所有工作表数据

        Returns:
            pd.DataFrame: 子查询结果
        """
        # sqlglot可能将子查询直接存储为Select(而非Subquery包装)
        if isinstance(subquery_expr, exp.Subquery):
            inner_select = subquery_expr.this
        elif isinstance(subquery_expr, exp.Select):
            inner_select = subquery_expr
        else:
            raise ValueError(f"不支持子查询类型: {type(subquery_expr)}")

        # 获取子查询的FROM表
        from_table, from_subquery = self._get_from_table(inner_select)
        if from_subquery is not None:
            raise ValueError("不支持嵌套FROM子查询(FROM子查询中不能再包含FROM子查询)")
        if from_table not in worksheets_data:
            raise ValueError(f"子查询中表 '{from_table}' 不存在.可用表: {list(worksheets_data.keys())}")

        # 复用现有查询执行逻辑
        try:
            result = self._execute_query(inner_select, worksheets_data)
            return result
        except Exception as e:
            raise ValueError(f"子查询执行失败: {e}")

    def _evaluate_case_expression(self, case_expr: exp.Case, df, row=None) -> Any:
        """
        评估CASE WHEN表达式

        支持格式:
        - CASE WHEN cond1 THEN val1 WHEN cond2 THEN val2 ELSE default END
        - 单行评估模式(row参数)或向量化模式(无row参数)

        Args:
            case_expr: sqlglot Case表达式
            df: DataFrame(向量化模式)
            row: 可选,单行数据(逐行模式)

        Returns:
            向量化模式返回pd.Series,逐行模式返回单个值
        """
        ifs = case_expr.args.get('ifs', [])
        default_value = case_expr.args.get('default')

        if row is not None:
            # 逐行评估模式
            for if_clause in ifs:
                condition = if_clause.this
                if self._evaluate_condition_for_row(condition, row):
                    return self._get_expression_value(if_clause.args.get('true'), row)
            # 没有匹配的WHEN,返回ELSE默认值
            if default_value is not None:
                return self._get_expression_value(default_value, row)
            return None
        else:
            # 向量化模式 - 复用逐行评估
            return pd.Series(
                [self._evaluate_case_expression(case_expr, df, df.iloc[i]) for i in range(len(df))],
                index=df.index
            )

    def _get_expression_value(self, expr: exp.Expression, row: pd.Series) -> Any:
        """获取表达式在指定行的值(委托给_get_row_value,两者功能完全重叠)"""
        return self._get_row_value(expr, row)

    def _evaluate_math_for_row(self, expr: exp.Expression, row: pd.Series) -> Any:
        """逐行评估数学表达式,复用类级别分发表"""
        op_type = type(expr)
        if op_type not in self._MATH_BINARY_OPS:
            return None
        left = self._get_expression_value(expr.left, row)
        right = self._get_expression_value(expr.right, row)
        try:
            left_n = float(left) if left is not None else None
            right_n = float(right) if right is not None else None
            if left_n is None or right_n is None:
                return None
            # 除零保护
            if op_type == exp.Div and right_n == 0:
                return None
            return self._MATH_BINARY_OPS[op_type](left_n, right_n)
        except (TypeError, ValueError):
            return None

    def _evaluate_coalesce_for_row(self, coalesce_expr: exp.Coalesce, row: pd.Series) -> Any:
        """逐行评估COALESCE/IFNULL表达式,空字符串转为0"""
        # COALESCE结构: this=第一个参数, expressions=[后续参数]
        values = [coalesce_expr.this] + list(coalesce_expr.expressions)
        for val_expr in values:
            val = self._get_expression_value(val_expr, row)
            # 跳过None/NaN,继续查找下一个参数
            if val is not None and not (isinstance(val, float) and np.isnan(val)):
                if val == '':
                    return 0  # 空字符串转为0
                return val
        return 0  # 所有参数都无效(None/NaN)时返回0

    def _evaluate_coalesce_vectorized(self, coalesce_expr: exp.Coalesce, df) -> pd.Series:
        """向量化评估COALESCE/IFNULL表达式(用于DataFrame),空字符串转为0

        使用 pandas combine_first 实现真正的向量化操作,
        替代逐行 _evaluate_coalesce_for_row 循环.
        仅当所有参数为列引用或字面量时可向量化,否则回退逐行.
        """
        values = [coalesce_expr.this] + list(coalesce_expr.expressions)
        result = None
        fallback = False

        for val_expr in values:
            if isinstance(val_expr, exp.Column) and val_expr.name in df.columns:
                series = df[val_expr.name].astype(object)
                # 空字符串转为0(在NaN处理之前)
                series = series.replace('', 0)
                # None/NaN保持不变,用于combine_first处理
            elif isinstance(val_expr, exp.Literal):
                v = self._parse_literal_value(val_expr)
                series = pd.Series([v] * len(df), index=df.index, dtype=object)
            else:
                fallback = True
                break

            if result is None:
                result = series
            else:
                result = result.combine_first(series)

        if fallback:
            results = [self._evaluate_coalesce_for_row(coalesce_expr, df.iloc[i]) for i in range(len(df))]
            return pd.Series(results, index=df.index)

        # 所有参数都无效时返回0
        return result.fillna(0)


    def _generate_aggregate_alias(self, expr: exp.Expression) -> str:
        """为无别名的聚合函数生成有意义的列名

        例: COUNT(*) -> count_star, AVG(damage) -> avg_damage, SUM(hp) -> sum_hp
        """
        func_name = type(expr).__name__.lower()  # count, sum, avg, max, min

        # 提取参数
        is_distinct = isinstance(expr.this, exp.Distinct)
        if is_distinct:
            target = expr.this.expressions[0]
        else:
            target = expr.this

        if isinstance(target, exp.Star):
            arg_name = "star"
        elif isinstance(target, exp.Column):
            arg_name = target.name
        elif hasattr(target, 'name') and target.name:
            arg_name = target.name
        else:
            arg_name = "expr"

        distinct_prefix = "distinct_" if is_distinct else ""
        return f"{func_name}_{distinct_prefix}{arg_name}"

    @staticmethod
    def _extract_agg_column(expr_node, context: str = "表达式") -> str:
        """从聚合函数参数节点提取列名(消除重复的Column/hasattr提取逻辑)"""
        if isinstance(expr_node, exp.Column):
            return expr_node.name
        if hasattr(expr_node, 'name') and expr_node.name:
            return expr_node.name
        raise ValueError(f"{context}参数格式错误: {expr_node}")

    # 聚合函数分发表:sum/avg/max/min 统一为 pd.to_numeric -> agg
    _AGG_OPS = {
        'sum': lambda g, col: g[col].agg(lambda x: pd.to_numeric(x, errors='coerce').sum()),
        'avg': lambda g, col: g[col].agg(lambda x: pd.to_numeric(x, errors='coerce').mean()),
        'max': lambda g, col: g[col].agg(lambda x: pd.to_numeric(x, errors='coerce').max()),
        'min': lambda g, col: g[col].agg(lambda x: pd.to_numeric(x, errors='coerce').min()),
    }

    def _apply_aggregation_function(self, expr: exp.Expression, grouped, df) -> pd.Series:
        """应用聚合函数"""
        if isinstance(expr, exp.Alias):
            return self._apply_aggregation_function(expr.this, grouped, df)

        if not isinstance(expr, exp.AggFunc):
            raise ValueError(f"不是聚合函数: {type(expr)}")

        func_name = type(expr).__name__.lower()

        # COUNT 特殊处理
        if func_name == 'count':
            if isinstance(expr.this, exp.Star):
                return grouped.size()
            if isinstance(expr.this, exp.Distinct):
                col_name = self._extract_agg_column(expr.this.expressions[0], "COUNT(DISTINCT)")
                return grouped[col_name].nunique()
            col_name = self._extract_agg_column(expr.this, "COUNT")
            return grouped[col_name].count()

        # 其他聚合函数不支持 *
        if isinstance(expr.this, exp.Star):
            raise ValueError(f"函数 {func_name} 不支持 * 参数")

        # 分发表处理 sum/avg/max/min
        if func_name in self._AGG_OPS:
            # 检查是否为表达式(如 攻击力+防御力)
            if self._is_expression(expr.this):
                # 对于表达式,先计算表达式列,再聚合
                expr_col = self._evaluate_expression(expr.this, df if df is not None else grouped)
                return self._AGG_OPS[func_name](grouped, expr_col)
            else:
                # 单列处理
                col_name = self._extract_agg_column(expr.this, func_name.upper())
                return self._AGG_OPS[func_name](grouped, col_name)

        raise ValueError(f"不支持的聚合函数: {func_name}")

    def _is_expression(self, node) -> bool:
        """检查是否为表达式(非单列)"""
        return not isinstance(node, exp.Column)

    def _evaluate_expression(self, expr_node, df) -> str:
        """计算表达式,返回临时列名"""
        if isinstance(expr_node, exp.Column):
            return expr_node.name
        
        # 处理字面量(数字)
        if isinstance(expr_node, exp.Literal):
            temp_col = f"_temp_literal_{id(expr_node)}"
            df[temp_col] = float(expr_node.this) if expr_node.this is not None else 0
            return temp_col
        
        # 生成临时列名
        temp_col = f"_temp_expr_{id(expr_node)}"
        
        # 处理加法表达式
        if isinstance(expr_node, exp.Add):
            left_col = self._evaluate_expression(expr_node.this, df)
            right_col = self._evaluate_expression(expr_node.expression, df)
            df[temp_col] = pd.to_numeric(df[left_col], errors='coerce') + pd.to_numeric(df[right_col], errors='coerce')
        
        # 处理减法表达式
        elif isinstance(expr_node, exp.Sub):
            left_col = self._evaluate_expression(expr_node.this, df)
            right_col = self._evaluate_expression(expr_node.expression, df)
            df[temp_col] = pd.to_numeric(df[left_col], errors='coerce') - pd.to_numeric(df[right_col], errors='coerce')
        
        # 处理乘法表达式
        elif isinstance(expr_node, exp.Mul):
            left_col = self._evaluate_expression(expr_node.this, df)
            right_col = self._evaluate_expression(expr_node.expression, df)
            df[temp_col] = pd.to_numeric(df[left_col], errors='coerce') * pd.to_numeric(df[right_col], errors='coerce')
        
        # 处理除法表达式
        elif isinstance(expr_node, exp.Div):
            left_col = self._evaluate_expression(expr_node.this, df)
            right_col = self._evaluate_expression(expr_node.expression, df)
            df[temp_col] = pd.to_numeric(df[left_col], errors='coerce') / pd.to_numeric(df[right_col], errors='coerce')
        
        else:
            raise ValueError(f"不支持的表达式类型: {type(expr_node)}")
        
        return temp_col

    def _apply_having_clause(self, parsed_sql: exp.Expression, df) -> pd.DataFrame:
        """应用HAVING条件"""
        having_clause = parsed_sql.args.get('having')
        if not having_clause:
            return df

        # 构建聚合表达式->SELECT别名的映射(HAVING COUNT(*) > 1 需要找到 cnt 列)
        self._having_agg_alias_map = {}
        for select_expr in parsed_sql.expressions:
            if isinstance(select_expr, exp.Alias) and isinstance(select_expr.this, exp.AggFunc):
                # 显式别名:AVG(伤害) as avg_dmg
                agg_sql = select_expr.this.sql()
                self._having_agg_alias_map[agg_sql] = select_expr.alias
            elif isinstance(select_expr, exp.AggFunc) and not isinstance(select_expr, exp.Alias):
                # 无别名:AVG(伤害) -> 自动生成 avg_damage
                agg_sql = select_expr.sql()
                auto_alias = self._generate_aggregate_alias(select_expr)
                self._having_agg_alias_map[agg_sql] = auto_alias

        # HAVING子句处理类似于WHERE,但作用于聚合后的数据
        try:
            condition_str = self._sql_condition_to_pandas(having_clause.this, df)
        except (ValueError, TypeError):
            # 不支持的条件类型(如COALESCE/CASE),回退到逐行过滤
            return self._apply_row_filter(having_clause.this, df)

        if condition_str:
            try:
                return df.query(condition_str)
            except Exception:
                # 备用方案:逐行过滤
                return self._apply_row_filter(having_clause.this, df)

        logger.warning("HAVING条件转换为pandas表达式失败,回退到逐行过滤: %s", having_clause.this)
        return self._apply_row_filter(having_clause.this, df)

    def _extract_select_aliases(self, parsed_sql: exp.Expression) -> Dict[str, Any]:
        """提取SELECT子句中的别名映射(委托_extract_select_alias统一逻辑)

        Returns:
            Dict: {alias_name: original_expression}
        """
        aliases = {}
        for i, select_expr in enumerate(parsed_sql.expressions):
            if isinstance(select_expr, exp.Star):
                continue
            alias_name, original_expr = self._extract_select_alias(select_expr, i)
            # 解包Paren表达式
            while isinstance(original_expr, exp.Paren):
                original_expr = original_expr.this
            aliases[alias_name] = original_expr
        return aliases

    def _resolve_order_column(self, col_name: str, df, select_aliases=None) -> Optional[str]:
        """解析ORDER BY列名:先查SELECT别名对应的基础列,再查原始列名

        Args:
            col_name: ORDER BY中引用的列名
            df: 当前DataFrame
            select_aliases: SELECT别名映射

        Returns:
            解析后的实际列名,找不到返回None
        """
        # 1. 如果列名直接在DataFrame中,直接返回
        if col_name in df.columns:
            return col_name

        # 2. 如果有SELECT别名映射,检查别名对应的基础列
        if select_aliases and col_name in select_aliases:
            expr = select_aliases[col_name]
            if isinstance(expr, exp.Column) and expr.name in df.columns:
                return expr.name
            # 别名对应的是计算表达式,临时计算后用于排序
            temp_col = self._compute_temp_column(expr, df, f"__order_temp_{col_name}")
            if temp_col is None:
                return None
            df.rename(columns={temp_col: col_name}, inplace=True)
            return col_name

        # 3. 列名不存在
        return None

    def _compute_temp_column(self, expr, df, temp_prefix="__temp__") -> Optional[str]:
        """将表达式计算结果写入临时列,支持数学/字符串/CASE/COALESCE

        Args:
            expr: SQL表达式
            df: DataFrame
            temp_prefix: 临时列名前缀

        Returns:
            临时列名,不支持的表达式返回None
        """
        temp_col = f"{temp_prefix}_{id(expr)}"
        try:
            if self._is_string_function(expr):
                df[temp_col] = self._evaluate_string_function(expr, df)
            elif isinstance(expr, exp.Case):
                df[temp_col] = self._evaluate_case_expression(expr, df)
            elif isinstance(expr, exp.Coalesce):
                df[temp_col] = self._evaluate_coalesce_vectorized(expr, df)
            elif self._is_mathematical_expression(expr):
                df[temp_col] = self._evaluate_math_expression(expr, df)
            else:
                return None
            return temp_col
        except Exception:
            return None

    def _resolve_order_expression(self, expr, df) -> Optional[str]:
        """解析ORDER BY中的函数表达式,临时计算并添加为排序列

        支持: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT,
              CASE WHEN, COALESCE, 数学表达式
        """
        return self._compute_temp_column(expr, df, "__order_expr")

    def _apply_order_by(self, parsed_sql: exp.Expression, df, select_aliases=None) -> pd.DataFrame:
        """应用ORDER BY排序

        Args:
            parsed_sql: 解析后的SQL表达式
            df: 数据DataFrame
            select_aliases: SELECT子句的别名映射(允许ORDER BY引用别名)
        """
        order_clause = parsed_sql.args.get('order')
        if not order_clause:
            return df

        sort_columns = []
        ascending = []

        for order_expr in order_clause.expressions:
            # 统一处理Ordered和简单列引用
            if isinstance(order_expr, exp.Ordered):
                col_expr = order_expr.this
                is_desc = order_expr.args.get('desc', False)
            elif isinstance(order_expr, exp.Column):
                col_expr = order_expr
                is_desc = False
            else:
                # 函数表达式: ORDER BY UPPER(name), ORDER BY LENGTH(name) 等
                col_expr = order_expr
                is_desc = False

            col_name = col_expr.name
            table_part = col_expr.table if hasattr(col_expr, 'table') and col_expr.table else None
            qualified = f"{table_part}.{col_name}" if table_part else None

            # 先查限定名,再查简单名,再查SELECT别名
            resolved_name = qualified if qualified and qualified in df.columns else None
            if resolved_name is None:
                resolved_name = self._resolve_order_column(col_name, df, select_aliases)
            if resolved_name is None and qualified and qualified in df.columns:
                resolved_name = qualified

            # 函数表达式: ORDER BY UPPER(col), LENGTH(col), COALESCE(col, 0) 等
            if resolved_name is None and not isinstance(col_expr, exp.Column):
                resolved_name = self._resolve_order_expression(col_expr, df)

            if resolved_name is None:
                suggestion = self._suggest_column_name(col_name, list(df.columns))
                raise ValueError(f"排序列 '{qualified or col_name}' 不存在.可用列: {list(df.columns)}.{suggestion}")

            sort_columns.append(resolved_name)
            ascending.append(not is_desc if is_desc is not None else True)

        if sort_columns:
            # Handle mixed data types in ORDER BY columns
            for col in sort_columns:
                if col in df.columns:
                    # Check if column has mixed data types (numbers and strings)
                    col_data = df[col]
                    has_numbers = False
                    has_strings = False
                    
                    for val in col_data.dropna():
                        if isinstance(val, (int, float)):
                            has_numbers = True
                        elif isinstance(val, str):
                            has_strings = True
                        
                        if has_numbers and has_strings:
                            # Mixed types - convert to string for consistent sorting
                            df[f'_temp_sort_{col}'] = col_data.astype(str)
                            sort_columns = [f'_temp_sort_{c}' if c == col else c for c in sort_columns]
                            break
            
            sorted_df = df.sort_values(by=sort_columns, ascending=ascending)
            
            # Clean up temporary columns
            for col in list(sort_columns):
                if col.startswith('_temp_sort_'):
                    sorted_df.drop(columns=[col], inplace=True)
            
            return sorted_df

        return df

    def _build_total_row(self, result_df: pd.DataFrame, group_by_columns: List[str] = []) -> Optional[List]:
        """构建GROUP BY聚合结果的TOTAL汇总行

        Args:
            result_df (pd.DataFrame): 结果DataFrame
            group_by_columns (List[str]): GROUP BY列名列表,跳过这些列的求和

        Returns:
            Optional[List]: TOTAL汇总行,如果没有数值列则返回None
        """
        if result_df.empty or len(result_df) <= 1:
            return None
        
        # Create a copy to avoid modifying original
        total_row = [''] * len(result_df.columns)
        
        # First pass: ensure GROUP BY columns are preserved correctly
        for i, col in enumerate(result_df.columns):
            if col in group_by_columns:
                # Copy the first value from each group for GROUP BY columns
                total_row[i] = result_df.iloc[0, i]
        
        # Second pass: calculate sums for non-GROUP BY numeric columns
        has_numeric = False
        for i, col in enumerate(result_df.columns):
            if col in group_by_columns:
                continue
                
            # Try to convert to numeric
            series = pd.to_numeric(result_df[col], errors='coerce')
            
            # Check if this column is numeric enough
            if series.notna().sum() > 0:
                # Calculate sum and serialize
                col_sum = series.sum()
                total_row[i] = self._serialize_value(col_sum)
                has_numeric = True
        
        # Mark as TOTAL row only if we have numeric aggregations
        if has_numeric:
            total_row[0] = 'TOTAL'
        
        return total_row if has_numeric else None

    def _generate_markdown_table(self, data: List, max_rows: int = MARKDOWN_TABLE_MAX_ROWS) -> str:
        """生成Markdown格式表格
        
        Args:
            data (List): 表格数据
            max_rows (int): 最大行数,默认使用配置值
            
        Returns:
            str: Markdown格式表格字符串
        """
        """将查询结果数据转为Markdown表格"""
        if not data:
            return ''
        md_lines = ['| ' + ' | '.join(str(c) for c in data[0]) + ' |']
        md_lines.append('| ' + ' | '.join(['---'] * len(data[0])) + ' |')
        display_rows = min(len(data) - 1, max_rows)
        for row in data[1:1 + display_rows]:
            md_lines.append('| ' + ' | '.join(str(c) for c in row) + ' |')
        if len(data) - 1 > max_rows:
            md_lines.append(f'| ... 共{len(data) - 1}行,仅显示前{max_rows}行 |')
        return '\n'.join(md_lines)

    def _format_export_output(self, data: List, output_format: str,
                               include_headers: bool) -> Dict[str, Any]:
        """生成JSON/CSV格式输出"""
        if not data or output_format == 'table':
            return {}
        headers_row = data[0]
        data_rows = data[1:]
        records = [dict(zip([str(h) for h in headers_row], row)) for row in data_rows]
        result = {'query_info': {'record_count': len(records)}}
        if output_format == 'json':
            result['formatted_output'] = json.dumps(records, ensure_ascii=False, indent=2)
            result['query_info']['output_format'] = 'json'
        elif output_format == 'csv':
            output = io.StringIO()
            writer = csv.writer(output)
            if include_headers:
                writer.writerow([str(h) for h in headers_row])
            for row in data_rows:
                writer.writerow([str(v) if v is not None else '' for v in row])
            result['formatted_output'] = output.getvalue()
            result['query_info']['output_format'] = 'csv'
        return result

    def _format_query_result(
        self,
        result_df: pd.DataFrame,
        file_path: str,
        sql: str,
        worksheets_data: Dict[str, pd.DataFrame],
        include_headers: bool,
        has_group_by: bool = False,
        has_having: bool = False,
        parsed_sql: exp.Expression = None,
        df_before_where: pd.DataFrame = None,
        output_format: str = "table"
    ) -> Dict[str, Any]:
        """格式化查询结果

        Args:
            has_group_by: 如果为True且有数值聚合列,自动追加TOTAL行
            has_having: 如果为True,表示使用了HAVING过滤,不添加TOTAL行
            parsed_sql: 解析后的SQL表达式,用于空结果智能建议
            df_before_where: WHERE过滤前的DataFrame,用于空结果智能建议
            output_format: 输出格式 table/json/csv
        """

        # 计算原始数据统计
        total_original_rows = sum(len(df) for df in worksheets_data.values())

        # 准备返回数据
        data = []
        if include_headers:
            data.append(list(result_df.columns))
        if not result_df.empty:
            for row in result_df.itertuples(index=False, name=None):
                data.append([self._serialize_value(val) for val in row])

        # 大结果自动截断:保护AI上下文窗口(MAX_RESULT_ROWS=500)
        truncated = False
        data_row_count = len(result_df)
        if data_row_count > MAX_RESULT_ROWS:
            # 保留表头行 + 前MAX_RESULT_ROWS行数据
            keep_rows = MAX_RESULT_ROWS + (1 if include_headers else 0)
            data = data[:keep_rows]
            truncated = True

        # GROUP BY 聚合结果自动追加 TOTAL 行(HAVING过滤时不添加)
        has_total_row = False
        if has_group_by and include_headers and not has_having:
            total_row = self._build_total_row(result_df, self._group_by_columns)
            if total_row:
                data.append(total_row)
                has_total_row = True

        # 双行表头:构建列描述映射
        column_descriptions = {}
        if hasattr(self, '_header_descriptions') and self._header_descriptions:
            for table_name, desc_map in self._header_descriptions.items():
                for col in (result_df.columns if not result_df.empty else []):
                    if col in desc_map:
                        column_descriptions[col] = desc_map[col]

        # 性能提示:无LIMIT且返回行数过多时建议加LIMIT
        perf_hint = ''
        if len(result_df) > 100:
            has_limit = parsed_sql is not None and parsed_sql.args.get('limit') is not None
            if not has_limit:
                perf_hint = f'(结果较多,建议加 LIMIT 缩小范围)'
        if truncated:
            perf_hint += f'(结果已截断为前{MAX_RESULT_ROWS}行,共{data_row_count}行,请加 LIMIT 精确查询)'

        result = {
            'success': True,
            'message': f'SQL查询成功执行,返回 {data_row_count} 行结果' + ('(含TOTAL汇总行)' if has_total_row else '') + perf_hint,
            'data': data,
            'query_info': {
                'original_rows': total_original_rows,
                'filtered_rows': data_row_count,
                'returned_rows': len(data) - (1 if include_headers else 0) - (1 if has_total_row else 0),
                'truncated': truncated,
                'query_applied': True,
                'sql_query': sql,
                'columns_returned': len(result_df.columns) if not result_df.empty else 0,
                'available_tables': list(worksheets_data.keys()),
                'returned_columns': list(result_df.columns) if not result_df.empty else [],
                'data_types': self._infer_data_types(result_df) if not result_df.empty else {}
            }
        }

        # 空结果智能建议:分析WHERE/HAVING条件类型,给出针对性提示
        if result_df.empty:
            suggestion = self._generate_empty_result_suggestion(
                parsed_sql, df_before_where, worksheets_data
            )
            # HAVING空结果追加聚合中间结果信息
            df_before_having = getattr(self, '_df_before_having', None)
            if df_before_having is not None and not df_before_having.empty:
                having_clause = parsed_sql.args.get('having')
                if having_clause:
                    suggestion += self._generate_having_empty_suggestion(
                        having_clause, df_before_having
                    )
            result['query_info']['suggestion'] = suggestion

        # 生成Markdown表格(方便AI和人类阅读)
        if data:
            result['query_info']['markdown_table'] = self._generate_markdown_table(data)

        # 生成JSON/CSV格式输出
        if data:
            export = self._format_export_output(data, output_format, include_headers)
            for key, value in export.items():
                if key == 'query_info':
                    result['query_info'].update(value)
                else:
                    result[key] = value

        # 双行表头时附加描述信息
        if column_descriptions:
            result['query_info']['dual_header'] = True
            result['query_info']['column_descriptions'] = column_descriptions

        return result

    def _infer_data_types(self, df) -> Dict[str, str]:
        """推断列的数据类型"""
        data_types = {}

        for col in df.columns:
            series = df[col]

            # 检查是否为数值类型
            numeric_series = pd.to_numeric(series, errors='coerce')
            if not numeric_series.isna().all():
                if (numeric_series % 1 == 0).all():
                    data_types[col] = 'integer'
                else:
                    data_types[col] = 'float'
                continue

            # 检查是否为日期类型
            try:
                # 先检查是否明显是日期格式
                sample_values = series.dropna().head(5)
                is_likely_date = False
                for val in sample_values:
                    if isinstance(val, str) and any(x in str(val) for x in ['-', '/', ':', '年', '月', '日']):
                        is_likely_date = True
                        break

                if is_likely_date:
                    converted = pd.to_datetime(series, errors='coerce', format='mixed')
                    if not converted.isna().all():
                        data_types[col] = 'datetime'
                        continue
            except Exception:
                pass

            # 默认为字符串类型
            data_types[col] = 'string'

        return data_types

    def _update_error(self, message: str, elapsed_ms: float = 0) -> Dict[str, Any]:
        """构造UPDATE操作的统一错误响应"""
        result = {'success': False, 'message': message,
                  'affected_rows': 0, 'changes': []}
        if elapsed_ms:
            result['execution_time_ms'] = round(elapsed_ms, 1)
        return result

    def execute_update_query(
        self,
        file_path: str,
        sql: str,
        sheet_name: Optional[str] = None,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        执行UPDATE语句,基于WHERE条件批量修改Excel数据

        支持语法: UPDATE 表名 SET 列1=值1, 列2=值2 [WHERE 条件]
        SET表达式支持: 列=常量, 列=列, 列=算术表达式(如 伤害*1.1)
        WHERE条件复用查询引擎的所有条件语法

        Args:
            file_path: Excel文件路径
            sql: UPDATE SQL语句
            sheet_name: 工作表名称(可选)
            dry_run: 预览模式,只返回影响行数不实际修改

        Returns:
            Dict: 更新结果
        """
        start_time = time.time()

        # 验证文件
        if not os.path.exists(file_path):
            return self._update_error(f'文件不存在: {file_path}')

        if not SQLGLOT_AVAILABLE:
            return self._update_error('SQLGLOT未安装,无法使用UPDATE功能')

        # 加载数据(使用缓存)
        worksheets_data = self._load_data_with_cache(file_path, sheet_name)

        if not worksheets_data:
            return self._update_error('无法加载Excel数据')

        # 清理ANSI转义序列(终端粘贴可能带入的不可见字符)
        sql = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', sql)

        # 解析UPDATE语句
        # 中文列名替换(与SELECT查询保持一致)
        try:
            sql = self._replace_cn_columns_in_sql(sql, worksheets_data)
        except Exception:
            pass  # 替换失败时继续用原始SQL

        try:
            parsed = sqlglot.parse_one(sql, read='mysql')
        except ParseError as e:
            return self._update_error(f'SQL语法错误: {e}')

        # 验证是UPDATE语句
        if not isinstance(parsed, exp.Update):
            return self._update_error('只支持UPDATE语句.💡 写入操作只支持UPDATE,查询请用 excel_query')

        # 提取表名(sqlglot中table在this属性)
        table_node = parsed.this if isinstance(parsed.this, exp.Table) else None
        if not table_node:
            return self._update_error('UPDATE语句缺少表名')
        target_table = table_node.name

        # 匹配工作表(支持中英文表名)
        matched_sheet = None
        for sheet in worksheets_data:
            if sheet == target_table:
                matched_sheet = sheet
                break
            # 模糊匹配
            if target_table.lower() == sheet.lower():
                matched_sheet = sheet
                break

        if not matched_sheet:
            available = list(worksheets_data.keys())
            suggestion = self._suggest_column_name(target_table, available)
            return self._update_error(
                f"工作表 '{target_table}' 不存在.可用工作表: {available}.{suggestion}")

        df = worksheets_data[matched_sheet].copy()
        original_df = df.copy()

        # 中文列名替换
        cn_map = {}
        desc_map = self._header_descriptions.get(matched_sheet, {})
        for en_col, cn_desc in desc_map.items():
            if en_col in df.columns:
                cn_map[cn_desc] = en_col

        # 解析SET子句(sqlglot中在expressions属性)
        set_exprs = parsed.args.get('expressions', [])
        if not set_exprs:
            return self._update_error('UPDATE语句缺少SET子句')

        set_operations = []  # [(col_name, expression_node)]
        for set_item in set_exprs:
            # sqlglot Update SET items: EQ expression (col = value)
            if isinstance(set_item, exp.EQ):
                col_name = set_item.left.name
                # 中文列名替换
                if col_name in cn_map:
                    col_name = cn_map[col_name]
                if col_name not in df.columns:
                    suggestion = self._suggest_column_name(col_name, list(df.columns))
                    return self._update_error(
                        f"列 '{col_name}' 不存在.可用列: {list(df.columns)}.{suggestion}")
                set_operations.append((col_name, set_item.right))
            else:
                return self._update_error(f'不支持的SET表达式: {set_item}')

        # 应用WHERE条件筛选
        where_clause = parsed.args.get('where')
        if where_clause:
            condition_str = self._sql_condition_to_pandas(where_clause.this, df)
            if condition_str:
                try:
                    filtered_df = df.query(condition_str)
                except Exception:
                    filtered_df = self._apply_row_filter(where_clause.this, df)
            else:
                logger.warning("UPDATE WHERE条件转换为pandas表达式失败,回退到逐行过滤: %s", where_clause.this)
                filtered_df = self._apply_row_filter(where_clause.this, df)
        else:
            filtered_df = df

        if filtered_df.empty:
            return {'success': True, 'message': '没有匹配WHERE条件的行,无需更新',
                    'affected_rows': 0, 'changes': [], 'execution_time_ms': 0}

        affected_indices = filtered_df.index.tolist()
        changes = []

        # 应用SET操作
        for col_name, value_expr in set_operations:
            for idx in affected_indices:
                old_val = df.at[idx, col_name]
                new_val = self._evaluate_update_expression(value_expr, df, idx)

                # 类型兼容性:数值类型可互通,其他类型尝试转为旧值类型
                if old_val != '' and new_val != '' and type(old_val) != type(new_val):
                    if isinstance(old_val, (int, float)) and isinstance(new_val, (int, float)):
                        pass  # 数值互通:不转换
                    else:
                        try:
                            new_val = type(old_val)(new_val)
                        except (ValueError, TypeError):
                            pass

                if old_val != new_val:
                    changes.append({
                        'row': int(idx) + 2,  # +2 for header offset (0-indexed + header row)
                        'column': col_name,
                        'old_value': self._serialize_update_value(old_val),
                        'new_value': self._serialize_update_value(new_val)
                    })
                df.at[idx, col_name] = new_val

        if not changes:
            elapsed = (time.time() - start_time) * 1000
            return {'success': True,
                    'message': f'匹配 {len(affected_indices)} 行,但值无变化',
                    'affected_rows': len(affected_indices),
                    'changes': [], 'execution_time_ms': round(elapsed, 1)}

        if dry_run:
            elapsed = (time.time() - start_time) * 1000
            return {'success': True,
                    'message': f'[预览] 将修改 {len(changes)} 个单元格({len(affected_indices)} 行)',
                    'affected_rows': len(affected_indices),
                    'changes': changes, 'dry_run': True,
                    'execution_time_ms': round(elapsed, 1)}

        # 写回Excel(事务保护:失败自动回滚)
        try:
            return self._write_changes_to_excel(
                file_path, matched_sheet, changes, df,
                len(affected_indices), start_time)
        except Exception as e:
            elapsed = (time.time() - start_time) * 1000
            return {'success': False,
                    'message': f'写入Excel失败,已自动回滚: {e}',
                    'affected_rows': 0, 'changes': changes,
                    'execution_time_ms': round(elapsed, 1)}

    def _write_changes_to_excel(self, file_path: str, sheet_name: str,
                                changes: list, df: pd.DataFrame,
                                affected_rows: int, start_time: float) -> Dict[str, Any]:
        """事务保护写入变更到Excel(失败自动回滚)
        支持流式写入:大文件批量修改时使用write_only模式提升性能
        """
        backup_path = None
        try:
            with self._file_lock(file_path):
                backup_path = tempfile.mktemp(suffix='.xlsx.bak')
                shutil.copy2(file_path, backup_path)

                # 决策:使用流式写入的条件
                file_size = os.path.getsize(file_path)
                use_streaming = (
                    affected_rows >= STREAMING_WRITE_MIN_ROWS or  # 影响行数>=阈值
                    len(changes) >= STREAMING_WRITE_MIN_CHANGES or  # 修改单元格数>=阈值
                    file_size > STREAMING_WRITE_MIN_FILE_SIZE_MB * 1024 * 1024  # 文件大小>阈值
                )

                if use_streaming and StreamingWriter.is_available():
                    # 使用流式写入(高性能路径)
                    # _copy_modify_write 会传递 (rows, header_row=1, col_map) 给 modify_fn
                    # col_map: {列名: 列索引(1-based)}
                    # rows: 包含表头行的所有行数据(0-based索引)
                    # header_row: 表头行数(固定为1)
                    # change['row']: 1-based的Excel行号(execute_update_query中 int(idx)+2)
                    
                    def modify_fn(rows, header_row, col_map):
                        """修改函数:应用UPDATE变更到行数据

                        Args:
                            rows: 行数据列表
                            header_row: 表头行
                            col_map: 列名到列索引的映射
                        """
                        modified_rows = [row[:] for row in rows]  # 深拷贝
                        
                        for change in changes:
                            col_name = change['column']
                            # 在col_map中查找列索引
                            col_idx = col_map.get(col_name) or col_map.get(str(col_name))
                            if col_idx is None:
                                continue
                            
                            # change['row']是1-based的Excel行号
                            # rows是0-based的列表(rows[0]是表头)
                            # 所以 Excel行号N -> rows[N-1]
                            list_idx = change['row'] - 1
                            
                            if 0 <= list_idx < len(modified_rows):
                                row = modified_rows[list_idx]
                                while len(row) < col_idx:
                                    row.append('')
                                row[col_idx - 1] = change['new_value']
                        
                        return True, "流式写入完成", modified_rows, {}
                    
                    success, message, meta = StreamingWriter._copy_modify_write(
                        file_path, sheet_name, modify_fn, preserve_col_widths=True
                    )
                    
                    if success:
                        self._df_cache.pop(file_path, None)
                        elapsed = (time.time() - start_time) * 1000
                        return {
                            'success': True,
                            'message': f'流式更新 {len(changes)} 个单元格({affected_rows} 行)',
                            'affected_rows': affected_rows,
                            'changes': changes,
                            'execution_time_ms': round(elapsed, 1),
                            'method': 'streaming'
                        }
                    else:
                        # 流式写入失败,降级到传统方式
                        logger.warning(f"流式写入失败,降级到传统方式: {message}")

                # 传统写入方式(兼容性路径)
                header_row_offset = 0
                header_desc = getattr(self, '_header_descriptions', {})
                if header_desc.get(sheet_name, {}):
                    header_row_offset = 1

                wb = openpyxl.load_workbook(file_path)
                ws = wb[sheet_name]

                for change in changes:
                    excel_row = change['row'] + header_row_offset
                    col_idx = list(df.columns).index(change['column']) + 1
                    ws.cell(row=excel_row, column=col_idx, value=change['new_value'])

                wb.save(file_path)
                wb.close()

                if backup_path and os.path.exists(backup_path):
                    os.remove(backup_path)

                self._df_cache.pop(file_path, None)

                elapsed = (time.time() - start_time) * 1000
                return {'success': True,
                        'message': f'成功更新 {len(changes)} 个单元格({affected_rows} 行)',
                        'affected_rows': affected_rows,
                        'changes': changes,
                        'execution_time_ms': round(elapsed, 1),
                        'method': 'traditional'}
        except Exception:
            if backup_path and os.path.exists(backup_path):
                try:
                    shutil.copy2(backup_path, file_path)
                    os.remove(backup_path)
                except Exception:
                    pass
            raise  # 重新抛出让调用方处理

    @contextmanager
    def _file_lock(self, file_path: str) -> Generator[None, None, None]:
        """文件锁上下文管理器(Linux fcntl,其他平台优雅降级)"""
        lock_fd = None
        try:
            try:
                import fcntl
                lock_fd = open(file_path + '.lock', 'w', encoding='utf-8')
                fcntl.flock(lock_fd, fcntl.LOCK_EX)
            except (ImportError, OSError):
                lock_fd = None
            yield
        finally:
            if lock_fd:
                try:
                    import fcntl
                    fcntl.flock(lock_fd, fcntl.LOCK_UN)
                    lock_path = file_path + '.lock'
                    if os.path.exists(lock_path):
                        os.remove(lock_path)
                except Exception:
                    pass
                lock_fd.close()

    def _serialize_value(self, val: Any) -> Any:
        """智能序列化值:数值保持数值类型,None/NaN转空字符串,numpy->Python原生"""
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return ''
        if isinstance(val, (np.integer,)):
            return int(val)
        if isinstance(val, (np.floating,)):
            f = float(val)
            if np.isnan(f):
                return ''
            return int(f) if f == int(f) else round(f, 2)
        if isinstance(val, float):
            if np.isnan(val):
                return ''
            return int(val) if val == int(val) else round(val, 2)
        return val

    def _serialize_update_value(self, val: Any) -> Any:
        """将值序列化为JSON安全类型(numpy->Python原生)-- 委托给_serialize_value"""
        return self._serialize_value(val)

    def _evaluate_update_expression(
        self, expr: exp.Expression, df: pd.DataFrame, row_idx: int
    ) -> Any:
        """
        评估UPDATE SET表达式,支持常量,列引用和算术运算

        Args:
            expr: SQL表达式
            df: DataFrame
            row_idx: 行索引

        Returns:
            计算后的值
        """
        if isinstance(expr, exp.Literal):
            return self._parse_literal_value(expr)

        elif isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name in df.columns:
                return df.at[row_idx, col_name]
            return ''

        elif isinstance(expr, exp.Neg):
            inner = self._evaluate_update_expression(expr.this, df, row_idx)
            try:
                return -float(inner)
            except (ValueError, TypeError):
                return inner

        elif isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div)):
            left = self._evaluate_update_expression(expr.left, df, row_idx)
            right = self._evaluate_update_expression(expr.right, df, row_idx)
            try:
                left_n = float(left) if not isinstance(left, (int, float)) else left
                right_n = float(right) if not isinstance(right, (int, float)) else right
                # 复用类级别分发表,支持所有算术运算符
                op = self._MATH_BINARY_OPS[type(expr)]
                result = op(left_n, right_n if type(expr) != exp.Div or right_n != 0 else 0)
                # 如果原值都是整数(含numpy整数)且非除法,返回整数
                if isinstance(left, (int, np.integer)) and isinstance(right, (int, np.integer)) and type(expr) != exp.Div:
                    return int(result)
                return result
            except (ValueError, TypeError):
                return ''

        else:
            # 未知表达式类型,尝试递归
            if hasattr(expr, 'this'):
                return self._evaluate_update_expression(expr.this, df, row_idx)
            return ''


# 模块级单例引擎,DataFrame缓存跨调用共享
_shared_engine: Optional[AdvancedSQLQueryEngine] = None


def _get_engine() -> AdvancedSQLQueryEngine:
    """获取共享SQL引擎实例(缓存跨调用复用)"""
    global _shared_engine
    if _shared_engine is None:
        _shared_engine = AdvancedSQLQueryEngine()
    return _shared_engine

def execute_advanced_sql_query(
    file_path: str,
    sql: str,
    sheet_name: Optional[str] = None,
    limit: Optional[int] = None,
    include_headers: bool = True,
    output_format: str = "table"
) -> Dict[str, Any]:
    """
    便捷函数:执行高级SQL查询

    Args:
        file_path: Excel文件路径
        sql: SQL查询语句
        sheet_name: 工作表名称(可选)
        limit: 结果限制
        include_headers: 是否包含表头
        output_format: 输出格式 table/json/csv

    Returns:
        Dict: 查询结果
    """
    try:
        engine = _get_engine()
        return engine.execute_sql_query(
            file_path=file_path,
            sql=sql,
            sheet_name=sheet_name,
            limit=limit,
            include_headers=include_headers,
            output_format=output_format
        )
    except ImportError as e:
        return {
            'success': False,
            'message': f'SQLGlot未安装,无法使用高级SQL功能: {str(e)}',
            'data': [],
            'query_info': {'error_type': 'missing_dependency', 'dependency': 'sqlglot'}
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'高级SQL查询失败: {str(e)}',
            'data': [],
            'query_info': {'error_type': 'engine_error', 'details': str(e)}
        }


def execute_advanced_update_query(
    file_path: str,
    sql: str,
    sheet_name: Optional[str] = None,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    便捷函数:执行UPDATE SQL语句

    Args:
        file_path: Excel文件路径
        sql: UPDATE SQL语句
        sheet_name: 工作表名称(可选)
        dry_run: 预览模式

    Returns:
        Dict: 更新结果
    """
    try:
        engine = _get_engine()
        return engine.execute_update_query(
            file_path=file_path,
            sql=sql,
            sheet_name=sheet_name,
            dry_run=dry_run
        )
    except ImportError as e:
        return {
            'success': False,
            'message': f'SQLGLOT未安装,无法使用UPDATE功能: {str(e)}',
            'affected_rows': 0, 'changes': [],
            'query_info': {'error_type': 'missing_dependency', 'dependency': 'sqlglot'}
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'UPDATE执行失败: {str(e)}',
            'affected_rows': 0, 'changes': [],
            'query_info': {'error_type': 'engine_error', 'details': str(e)}
        }