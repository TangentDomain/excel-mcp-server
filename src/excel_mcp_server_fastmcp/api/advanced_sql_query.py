"""
高级SQL查询引擎 - 基于SQLGlot实现的SQL查询支持

支持功能：
- 基础查询: SELECT, DISTINCT, 别名
- 条件筛选: WHERE, LIKE, IN, BETWEEN, AND/OR, EXISTS, 子查询
- 聚合统计: COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 排序限制: ORDER BY, LIMIT, OFFSET
- 算术运算: 加减乘除
- 条件表达式: CASE WHEN, COALESCE/IFNULL
- 表关联: INNER JOIN, LEFT JOIN, RIGHT JOIN, FULL JOIN, CROSS JOIN（同文件内工作表关联）
- 子查询: WHERE col IN (SELECT ...), 标量子查询, EXISTS
- CTE: WITH ... AS (SELECT ...)
- 字符串函数: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT
- 窗口函数: ROW_NUMBER, RANK, DENSE_RANK（OVER PARTITION BY ... ORDER BY ...）
- 合并查询: UNION, UNION ALL

不支持功能：
- FROM子查询（FROM (SELECT ...)）
- FROM子查询（FROM (SELECT ...)）
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
    logger.warning("SQLGlot未安装，将使用基础pandas查询功能")
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


class AdvancedSQLQueryEngine:
    """高级SQL查询引擎，支持完整的SQL语法"""

    def __init__(self, disable_streaming_aggregate: bool = False):
        """
        初始化SQL查询引擎

        Args:
            disable_streaming_aggregate: 禁用流式聚合优化（大文件处理）
        """
        self.disable_streaming_aggregate = disable_streaming_aggregate
        # DataFrame缓存：{file_path: (mtime, worksheets_data, header_descriptions)}
        self._df_cache = {}
        self._max_cache_size = 10  # 最大缓存文件数，防止内存泄漏

        if not SQLGLOT_AVAILABLE:
            raise ImportError("SQLGlot未安装，请运行: pip install sqlglot")

    def clear_cache(self):
        """清除DataFrame缓存，释放内存"""
        self._df_cache.clear()

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
        执行SQL查询，支持完整的SQL语法

        Args:
            file_path: Excel文件路径
            sql: SQL查询语句
            sheet_name: 工作表名称（可选，默认使用第一个）
            limit: 限制返回行数
            include_headers: 是否包含表头
            output_format: 输出格式 table/json/csv

        Returns:
            Dict: 查询结果
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

            # 检查文件大小并处理大文件
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > 100:
                return {
                    'success': False,
                    'message': f'文件过大 ({file_size_mb:.2f}MB)，建议使用小于100MB的文件',
                    'data': [],
                    'query_info': {'error_type': 'file_too_large', 'size_mb': file_size_mb}
                }

            # 加载Excel数据（带缓存）
            worksheets_data = self._load_data_with_cache(file_path, sheet_name)

            if not worksheets_data:
                return {
                    'success': False,
                    'message': '无法加载Excel数据或文件为空',
                    'data': [],
                    'query_info': {'error_type': 'data_load_failed'}
                }

            # 中文列名替换：将SQL中的中文列名替换为英文列名（在解析前）
            sql = self._replace_cn_columns_in_sql(sql, worksheets_data)

            # DESCRIBE命令友好提示
            sql_stripped = sql.strip().upper()
            if sql_stripped.startswith('DESCRIBE ') or sql_stripped.startswith('DESC '):
                table_hint = sql.strip().split(None, 1)[-1].strip(';').strip('"\'`') if len(sql.strip().split()) > 1 else ''
                hint = f'请使用 excel_describe_table 工具查看表结构'
                if table_hint:
                    hint += f'（工作表: {table_hint}）'
                return {
                    'success': False,
                    'message': f'DESCRIBE不是SQL查询语法。{hint}',
                    'data': [],
                    'query_info': {'error_type': 'describe_not_sql', 'hint': 'use_excel_describe_table'}
                }

            # 解析和执行SQL
            _query_start = time.time()
            try:
                parsed_sql = sqlglot.parse_one(sql, dialect="mysql")

                # 验证SQL支持范围
                validation_result = self._validate_sql_support(parsed_sql)
                if not validation_result['valid']:
                    return {
                        'success': False,
                        'message': f'不支持的SQL语法: {validation_result["error"]}',
                        'data': [],
                        'query_info': {'error_type': 'unsupported_sql', 'details': validation_result}
                    }

                # 执行查询（UNION/UNION ALL 或普通 SELECT）
                if isinstance(parsed_sql, exp.Union):
                    result_data = self._execute_union(parsed_sql, worksheets_data, limit)
                else:
                    result_data = self._execute_query(parsed_sql, worksheets_data, limit)
                _query_elapsed = (time.time() - _query_start) * 1000

                # 格式化结果（传入parsed_sql和WHERE前数据用于空结果智能建议）
                has_group_by = not isinstance(parsed_sql, exp.Union) and parsed_sql.args.get('group') is not None
                result = self._format_query_result(
                    result_data,
                    file_path,
                    sql,
                    worksheets_data,
                    include_headers,
                    has_group_by=has_group_by,
                    parsed_sql=parsed_sql,
                    df_before_where=self._df_before_where,
                    output_format=output_format
                )
                # 注入执行时间
                result['query_info']['execution_time_ms'] = round(_query_elapsed, 1)
                return result

            except ParseError as e:
                err_str = str(e)
                # 提取常见拼写错误的友好提示
                hint = ''
                if 'SELEC' in sql.upper() and 'SELECT' not in sql.upper():
                    hint = '\n💡 提示：可能是拼写错误，SELECT关键字是否拼对了？'
                elif 'FORM' in sql.upper() and 'FROM' not in sql.upper():
                    hint = '\n💡 提示：可能是拼写错误，FROM关键字是否拼对了？'
                elif 'WHER' in sql.upper() and 'WHERE' not in sql.upper():
                    hint = '\n💡 提示：可能是拼写错误，WHERE关键字是否拼对了？'
                return {
                    'success': False,
                    'message': f'SQL语法错误: {err_str}{hint}',
                    'data': [],
                    'query_info': {'error_type': 'syntax_error', 'details': err_str}
                }
            except UnsupportedError as e:
                return {
                    'success': False,
                    'message': f'不支持的SQL功能: {str(e)}',
                    'data': [],
                    'query_info': {'error_type': 'unsupported_feature', 'details': str(e)}
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

    def _load_data_with_cache(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[Dict[str, pd.DataFrame]]:
        """
        带缓存的Excel数据加载（公共方法，供execute_sql_query和execute_update_query复用）

        使用mtime检测文件变更，LRU淘汰防止内存泄漏。

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称（可选）

        Returns:
            worksheets_data字典，加载失败返回None
        """
        mtime = os.path.getmtime(file_path)
        cache_key = file_path
        if cache_key in self._df_cache:
            cached_mtime, cached_data, cached_desc = self._df_cache[cache_key]
            if cached_mtime == mtime:
                self._header_descriptions = cached_desc
                return cached_data
            else:
                # 文件已修改，重新加载
                worksheets_data = self._load_excel_data(file_path, sheet_name)
                self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
                return worksheets_data
        else:
            worksheets_data = self._load_excel_data(file_path, sheet_name)
            self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
            # LRU淘汰：超过最大缓存数时删除最早缓存的文件
            while len(self._df_cache) > self._max_cache_size:
                self._df_cache.pop(next(iter(self._df_cache)))
            return worksheets_data

    def _load_excel_data(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, pd.DataFrame]:
        """
        加载Excel数据到DataFrame字典，支持游戏配置表双行表头

        游戏配置表通常有双行表头：
          第1行：中文描述（如"技能ID"、"技能名称"）
          第2行：字段名（如"skill_id"、"skill_name"）
        
        本方法自动检测双行表头，用第二行（字段名）做列名，
        第一行（描述）保存在 self._header_descriptions 中。

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称（可选）

        Returns:
            Dict[str, pd.DataFrame]: 工作表名到DataFrame的映射
        """
        worksheets_data = {}
        self._header_descriptions = {}  # {sheet_name: {field_name: description}}

        try:
            # 性能优化：用calamine替代openpyxl读取（Rust引擎，速度提升10-50倍）
            # calamine一次性读取所有sheet数据，无需二次打开文件
            from python_calamine import CalamineWorkbook

            cal_wb = CalamineWorkbook.from_path(file_path)
            all_sheet_names = cal_wb.sheet_names

            if sheet_name:
                sheets_to_load = [sheet_name] if sheet_name in all_sheet_names else []
            else:
                sheets_to_load = all_sheet_names

            # 批量检测所有sheet的双行表头（calamine读取前两行，毫秒级）
            header_info = {}  # {sheet: (is_dual_header, first_row_values, second_row_values)}
            for sheet in sheets_to_load:
                try:
                    cal_ws = cal_wb.get_sheet_by_name(sheet)
                    # 跳过空工作表（无数据行）
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

                        if len(non_empty_second) >= 3:
                            second_all_field = all(
                                re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', v)
                                for v in non_empty_second
                            )
                            first_all_field = all(
                                re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', v)
                                for v in non_empty_first
                            ) if non_empty_first else False
                            if second_all_field and not first_all_field:
                                is_dual_header = True

                    header_info[sheet] = (is_dual_header, first_row_values, second_row_values)
                except Exception:
                    header_info[sheet] = (False, None, None)

            # 批量读取所有sheet数据（pd.read_excel + calamine引擎）
            for sheet, (is_dual_header, first_row_values, second_row_values) in header_info.items():
                if is_dual_header:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='calamine',
                        header=1,
                        keep_default_na=False
                    )
                    # 从calamine读取的行值构建中英文映射
                    desc_map = {}
                    if second_row_values and first_row_values:
                        for fname, desc in zip(second_row_values, first_row_values):
                            fname = fname.strip() if fname else ''
                            desc = desc.strip() if desc else ''
                            if fname:
                                desc_map[fname] = desc
                    self._header_descriptions[sheet] = desc_map
                else:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='calamine',
                        keep_default_na=False
                    )

                df = self._clean_dataframe(df)
                worksheets_data[sheet] = df

        except Exception as e:
            logger.error(f"加载Excel数据失败: {e}")
            return {}

        return worksheets_data

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
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

            # 清理特殊字符，但保持中文
            clean_col = re.sub(r'[^\w\u4e00-\u9fff\s]', '_', clean_col)
            clean_col = re.sub(r'\s+', '_', clean_col)

            # 确保列名不为空且不以数字开头
            if not clean_col or clean_col.isspace():
                clean_col = f"column_{len(clean_columns) + 1}"
            elif clean_col[0].isdigit():
                clean_col = f"col_{clean_col}"

            clean_columns[col] = clean_col

        df = df.rename(columns=clean_columns)

        # 保持原始数据不做空值替换
        # pandas groupby 默认跳过 NaN 行，不需要手动处理

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
            # 检查是否为SELECT语句或UNION（UNION内部由多个SELECT组成）
            if not isinstance(parsed_sql, (exp.Select, exp.Union)):
                return {
                    'valid': False,
                    'error': '只支持SELECT查询语句，不支持INSERT、UPDATE、DELETE等操作'
                }

            # 子查询支持已实现，不再拒绝
            # IN子查询、标量子查询、EXISTS子查询均支持

            # CTE (WITH) 支持
            with_clause = parsed_sql.args.get('with_')
            if with_clause:
                return {'valid': True}  # CTE在_execute_query中处理
            # UNION/UNION ALL 支持已实现
            # 窗口函数支持已实现 (ROW_NUMBER, RANK, DENSE_RANK)

            return {'valid': True}

        except Exception as e:
            return {
                'valid': False,
                'error': f'SQL验证失败: {str(e)}'
            }

    def _replace_cn_columns_in_sql(self, sql: str, worksheets_data: Dict[str, pd.DataFrame]) -> str:
        """
        将SQL中的中文列名替换为英文列名（在sqlglot解析前）。

        双行表头的游戏配置表中，第1行是中文描述，第2行是英文字段名。
        策划习惯用中文名查询，但SQL引擎需要英文列名。
        本方法在SQL文本层面做替换，避免给DataFrame添加临时列。

        Args:
            sql: 原始SQL语句
            worksheets_data: 已加载的工作表数据

        Returns:
            str: 替换后的SQL语句
        """
        if not hasattr(self, '_header_descriptions') or not self._header_descriptions:
            return sql

        # 收集所有中文→英文映射（去重）
        cn_to_en = {}
        for sheet_name, desc_map in self._header_descriptions.items():
            for eng_name, cn_desc in desc_map.items():
                if cn_desc and cn_desc != eng_name:
                    cn_to_en[cn_desc] = eng_name

        if not cn_to_en:
            return sql

        # 按中文列名长度降序排列，避免短名称部分匹配长名称
        sorted_names = sorted(cn_to_en.keys(), key=len, reverse=True)

        # 用正则替换：只替换SQL标识符位置（非字符串字面量中的中文）
        # 策略：先把字符串字面量占位保护，替换中文标识符，再恢复字符串
        string_literals = []
        protected_sql = sql

        # 保护单引号字符串
        def protect_string(match):
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
        """分析WHERE条件类型，生成智能空结果建议"""
        where_clause = parsed_sql.args.get('where')
        if not where_clause:
            return '查询返回0行数据。表可能为空，请检查数据是否已录入。'

        total_rows = len(df_before_where)
        if total_rows == 0:
            return '查询返回0行数据。工作表本身没有数据行。'

        hints = []
        condition = where_clause.this

        # 分析条件树，收集条件类型和涉及的列
        eq_conditions = []  # 等值条件
        range_conditions = []  # 范围条件
        like_conditions = []  # LIKE条件
        in_conditions = []  # IN条件
        between_conditions = []  # BETWEEN条件
        null_conditions = []  # IS NULL条件

        self._collect_condition_types(condition, eq_conditions, range_conditions,
                                       like_conditions, in_conditions, between_conditions, null_conditions)

        # 等值条件：提示列的唯一值
        for col, val in eq_conditions:
            if col in df_before_where.columns:
                unique_vals = df_before_where[col].dropna().unique()
                if len(unique_vals) <= 20:
                    vals_str = ', '.join(str(v) for v in unique_vals[:10])
                    if len(unique_vals) > 10:
                        vals_str += f' ... 共{len(unique_vals)}个'
                    hints.append(f'• 列"{col}"的值为: {vals_str}')
                else:
                    hints.append(f'• 列"{col}"有{len(unique_vals)}个不同值，"{val}"不在其中')

        # 范围条件：提示列的实际范围
        for col, op, val in range_conditions:
            if col in df_before_where.columns:
                numeric = pd.to_numeric(df_before_where[col], errors='coerce').dropna()
                if len(numeric) > 0:
                    hints.append(f'• 列"{col}"的实际范围: {numeric.min():.2f} ~ {numeric.max():.2f}')
                else:
                    hints.append(f'• 列"{col}"不是数值列，无法用{op}比较')

        # LIKE条件：提示匹配情况
        for col, pattern in like_conditions:
            if col in df_before_where.columns:
                sample = df_before_where[col].dropna().astype(str).head(5).tolist()
                hints.append(f'• 列"{col}"的样本数据: {", ".join(sample)}')

        # IN条件：提示实际存在的值
        for col, vals in in_conditions:
            if col in df_before_where.columns:
                unique_vals = set(df_before_where[col].dropna().unique())
                matched = unique_vals & set(vals)
                if not matched:
                    hints.append(f'• 列"{col}"中不包含指定的任何值')

        # BETWEEN条件：提示列的实际范围
        for col, low, high in between_conditions:
            if col in df_before_where.columns:
                numeric = pd.to_numeric(df_before_where[col], errors='coerce').dropna()
                if len(numeric) > 0:
                    hints.append(f'• 列"{col}"的实际范围: {numeric.min():.2f} ~ {numeric.max():.2f}')

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
            hints.append('• 多个AND条件同时满足的行可能不存在，尝试减少条件或改用OR')

        # 通用提示
        hints.append(f'• 源表共{total_rows}行，WHERE过滤后为0行')
        hints.append('• 可用 DESCRIBE 查看表结构，或去掉WHERE先查看全部数据')

        return '查询返回0行数据。分析：\n' + '\n'.join(hints)

    def _collect_condition_types(self, condition, eq, rng, like, in_list, between, null_list):
        """递归收集WHERE条件树中的各类条件"""
        if isinstance(condition, exp.EQ):
            col = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            if col:
                eq.append((col, val))
        elif isinstance(condition, (exp.GT, exp.GTE, exp.LT, exp.LTE)):
            col = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            op_map = {exp.GT: '>', exp.GTE: '>=', exp.LT: '<', exp.LTE: '<='}
            if col:
                rng.append((col, op_map.get(type(condition), '?'), val))
        elif isinstance(condition, exp.Like):
            col = self._extract_column_name(condition.left)
            val = self._extract_literal_value(condition.right)
            if col:
                like.append((col, val))
        elif isinstance(condition, exp.In):
            col = self._extract_column_name(condition.this)
            vals = []
            if hasattr(condition, 'expressions'):
                for e in condition.expressions:
                    v = self._extract_literal_value(e)
                    if v is not None:
                        vals.append(v)
            if col and vals:
                in_list.append((col, vals))
        elif isinstance(condition, exp.Between):
            col = self._extract_column_name(condition.this)
            low = self._extract_literal_value(condition.args.get('low'))
            high = self._extract_literal_value(condition.args.get('high'))
            if col:
                between.append((col, low, high))
        elif isinstance(condition, exp.Is):
            col = self._extract_column_name(condition.this)
            if col:
                null_list.append((col, True))
        elif isinstance(condition, exp.Not):
            inner = condition.this
            if isinstance(inner, exp.Is):
                col = self._extract_column_name(inner.this)
                if col:
                    null_list.append((col, False))
            else:
                self._collect_condition_types(inner, eq, rng, like, in_list, between, null_list)
        elif isinstance(condition, exp.And):
            for child in condition.flatten():
                if child is not condition:
                    self._collect_condition_types(child, eq, rng, like, in_list, between, null_list)
        elif isinstance(condition, exp.Or):
            for child in condition.flatten():
                if child is not condition:
                    self._collect_condition_types(child, eq, rng, like, in_list, between, null_list)

    def _extract_column_name(self, expr):
        """从表达式中提取列名"""
        if isinstance(expr, exp.Column):
            return expr.name
        return None

    def _extract_literal_value(self, expr):
        """从表达式中提取字面值"""
        if isinstance(expr, exp.Literal):
            val = expr.this
            try:
                return int(val)
            except (ValueError, TypeError):
                try:
                    return float(val)
                except (ValueError, TypeError):
                    return val
        return None

    def _generate_having_empty_suggestion(self, having_expr, df_before_having) -> str:
        """生成HAVING导致空结果时的智能建议

        Args:
            having_expr: HAVING表达式（完整的having clause）
            df_before_having: HAVING过滤前的聚合结果DataFrame
        """
        hints = ['\nHAVING分析：']
        hints.append(f'• GROUP BY聚合后有{len(df_before_having)}组数据')

        condition = having_expr.this
        col = self._extract_column_name(condition.left)
        val = self._extract_literal_value(condition.right)

        # HAVING中聚合函数表达式的列名可能是别名，尝试从DataFrame列匹配
        if not col:
            left_str = str(condition.left).lower()
            # 策略1：子串匹配
            for c in df_before_having.columns:
                if c.lower() in left_str or left_str in c.lower():
                    col = c
                    break
            # 策略2：拆分表达式中的标识符（如AVG(damage) → 检查含avg和damage的列）
            if not col:
                tokens = set(re.findall(r'[a-zA-Z_]+', left_str))
                if tokens:
                    for c in df_before_having.columns:
                        c_tokens = set(re.findall(r'[a-zA-Z_]+', c.lower()))
                        # 至少有一个token匹配（排除通用token如avg, sum, count, min, max）
                        generic = {'avg', 'sum', 'count', 'min', 'max'}
                        specific = tokens - generic
                        if specific and specific & c_tokens:
                            col = c
                            break

        if not col or col not in df_before_having.columns:
            # 无法匹配列名，显示所有聚合列的实际范围
            if len(df_before_having.columns) > 0:
                for c in df_before_having.columns:
                    numeric = pd.to_numeric(df_before_having[c], errors='coerce').dropna()
                    if len(numeric) > 0:
                        hints.append(f'• 列"{c}"范围: {numeric.min()} ~ {numeric.max()}')
            hints.append('• HAVING条件较复杂，建议去掉HAVING先查看聚合结果')
            hints.append('• 可先去掉HAVING查看全部分组结果，再调整过滤条件')
            return '\n'.join(hints)

        numeric = pd.to_numeric(df_before_having[col], errors='coerce').dropna()
        if len(numeric) == 0:
            hints.append(f'• 列"{col}"没有数值数据')
            hints.append('• 可先去掉HAVING查看全部分组结果，再调整过滤条件')
            return '\n'.join(hints)

        if isinstance(condition, exp.GT):
            col_max = numeric.max()
            hints.append(f'• 列"{col}"的最大值为{col_max}，HAVING要求 >{val}，无满足条件的组')
        elif isinstance(condition, exp.GTE):
            col_max = numeric.max()
            hints.append(f'• 列"{col}"的最大值为{col_max}，HAVING要求 >={val}，无满足条件的组')
        elif isinstance(condition, exp.LT):
            col_min = numeric.min()
            hints.append(f'• 列"{col}"的最小值为{col_min}，HAVING要求 <{val}，无满足条件的组')
        elif isinstance(condition, exp.LTE):
            col_min = numeric.min()
            hints.append(f'• 列"{col}"的最小值为{col_min}，HAVING要求 <={val}，无满足条件的组')
        elif isinstance(condition, exp.EQ):
            unique_vals = df_before_having[col].dropna().unique()
            if len(unique_vals) <= 10:
                vals_str = ', '.join(str(v) for v in unique_vals)
                hints.append(f'• 列"{col}"的值为: {vals_str}，不等于{val}')
            else:
                hints.append(f'• 列"{col}"有{len(unique_vals)}个不同值，不等于{val}')
        else:
            hints.append(f'• HAVING条件较复杂，建议去掉HAVING先查看聚合结果')

        hints.append('• 可先去掉HAVING查看全部分组结果，再调整过滤条件')
        return '\n'.join(hints)

    def _suggest_column_name(self, col_name: str, available_cols: List[str], max_suggestions: int = 3) -> str:
        """
        当列名不存在时，用编辑距离找出最相似的列名作为建议。
        同时检查中文列名描述（双行表头场景）。

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

        # 英文列名匹配不到时，尝试匹配中文列名描述（双行表头）
        if hasattr(self, '_header_descriptions') and self._header_descriptions:
            cn_names = []
            for sheet_name, desc_map in self._header_descriptions.items():
                for eng_name, cn_desc in desc_map.items():
                    if cn_desc:
                        cn_names.append(cn_desc)
            if cn_names:
                cn_matches = difflib.get_close_matches(col_name, cn_names, n=max_suggestions, cutoff=0.3)
                if cn_matches:
                    # 反查中文→英文映射
                    en_lookup = {}
                    for desc_map in self._header_descriptions.values():
                        for eng_name, cn_desc in desc_map.items():
                            if cn_desc:
                                en_lookup[cn_desc] = eng_name
                    mapped = [f"{cn}({en_lookup.get(cn, '?')})" for cn in cn_matches]
                    return f" 中文名称匹配: {', '.join(mapped)}"

        return ""

    def _apply_union_order_by(self, df: pd.DataFrame, order_clause: exp.Expression) -> pd.DataFrame:
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

            # 尝试列名匹配（包括中文列名映射）
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

        从 Union 表达式中提取所有 SELECT 语句，分别执行后合并结果。
        UNION 去重，UNION ALL 保留所有行。
        支持 ORDER BY 和 LIMIT 应用于合并后的结果。

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
                # this 可能是 Union（链式）或 Select
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

        # 合并所有结果（列名对齐）
        if not result_dfs:
            return pd.DataFrame()

        # 以第一个 SELECT 的列名为基准，统一列名
        base_columns = list(result_dfs[0].columns)
        aligned_dfs = []
        for df in result_dfs:
            aligned = df.reindex(columns=base_columns)
            aligned_dfs.append(aligned)

        combined = pd.concat(aligned_dfs, ignore_index=True)

        # UNION（去重） vs UNION ALL（保留重复）
        is_union_all = not parsed_sql.args.get('distinct', True)
        if not is_union_all:
            combined = combined.drop_duplicates().reset_index(drop=True)

        # 应用 ORDER BY（如果有，sqlglot 将其放在外层 Union 上）
        order_clause = parsed_sql.args.get('order')
        if order_clause:
            # 构造一个最小化的 Select 用于 _apply_order_by 的签名
            # _apply_order_by(self, parsed_sql, df, select_aliases) 期望完整的 parsed_sql
            # 但 UNION 的 ORDER BY 是独立的，直接解析排序列
            combined = self._apply_union_order_by(combined, order_clause)

        # 应用 LIMIT（如果有）
        union_limit = parsed_sql.args.get('limit')
        if union_limit:
            try:
                limit_expr = union_limit.expression or union_limit.this
                if limit_expr is not None:
                    limit_val = int(limit_expr.this if hasattr(limit_expr, 'this') else limit_expr)
                    combined = combined.head(limit_val)
            except (ValueError, AttributeError, TypeError):
                pass

        # 应用外部传入的 limit
        if limit is not None:
            combined = combined.head(limit)

        return combined

    def _has_window_function(self, parsed_sql: exp.Expression) -> bool:
        """检查SQL是否包含窗口函数"""
        return bool(parsed_sql.find(exp.Window))

    def _apply_window_functions(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """
        计算窗口函数并将结果添加到DataFrame
        支持: ROW_NUMBER, RANK, DENSE_RANK
        语法: func() OVER ([PARTITION BY col ...] ORDER BY col [ASC|DESC] ...)
        """
        if not self._has_window_function(parsed_sql):
            return df

        df = df.copy()

        # 构建SELECT别名映射（用于将聚合表达式映射到别名，如AVG(damage)→avg_dmg）
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
        """计算单个窗口函数，返回结果Series"""
        func_type = type(window_expr.this).__name__

        # 支持的窗口函数类型
        supported_funcs = {'RowNumber', 'Rank', 'DenseRank'}
        if func_type not in supported_funcs:
            raise ValueError(f"不支持的窗口函数: {func_type}。支持的: ROW_NUMBER, RANK, DENSE_RANK")

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
                # 如果列名不在DataFrame中，尝试映射
                if col_name not in df.columns:
                    col_name = self._resolve_window_column(col_name, df.columns, select_alias_map)
                order_cols.append(col_name)
                ascending.append(not ordered_expr.args.get('desc', False))

        # 验证列存在
        for col in partition_cols + order_cols:
            if col not in df.columns:
                suggestion = self._suggest_column_name(col, list(df.columns))
                raise ValueError(f"窗口函数中列 '{col}' 不存在。可用列: {list(df.columns)}。{suggestion}")

        if func_type == 'RowNumber':
            return self._compute_row_number(df, partition_cols, order_cols, ascending)
        elif func_type == 'Rank':
            return self._compute_rank(df, partition_cols, order_cols, ascending)
        elif func_type == 'DenseRank':
            return self._compute_dense_rank(df, partition_cols, order_cols, ascending)

    def _resolve_window_column(self, col_name: str, df_columns: list,
                                select_alias_map: Dict[str, str]) -> str:
        """解析窗口函数中的列名（支持聚合表达式→别名映射）"""
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

        return col_name  # 未找到映射，返回原名

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
            if order_cols:
                sorted_group = group.sort_values(order_cols, ascending=ascending)
                result = pd.Series(range(1, len(sorted_group) + 1), index=sorted_group.index, dtype=int)
                return result.reindex(group.index)
            else:
                return pd.Series(range(1, len(group) + 1), index=group.index, dtype=int)

        if grouped is not None:
            result = grouped.apply(assign_row_number, include_groups=False)
            # groupby.apply可能返回MultiIndex Series，需要展平
            if isinstance(result.index, pd.MultiIndex):
                result = result.droplevel(result.index.names[:-1])
        else:
            result = assign_row_number(df)

        return result

    def _compute_rank(self, df: pd.DataFrame, partition_cols: list,
                      order_cols: list, ascending: list) -> pd.Series:
        """RANK: 相同值相同排名，下一个排名跳过（1,2,2,4）"""
        if not order_cols:
            raise ValueError("RANK() 窗口函数需要 ORDER BY 子句")

        if partition_cols:
            grouped = df.groupby(partition_cols, sort=False)
        else:
            grouped = None

        def assign_rank(group):
            sorted_group = group.sort_values(order_cols, ascending=ascending)
            # 使用pandas rank(method='first')模拟RANK行为
            # RANK: 相同值取相同排名，下一个排名跳过
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
        """DENSE_RANK: 相同值相同排名，下一个排名不跳过（1,2,2,3）"""
        if not order_cols:
            raise ValueError("DENSE_RANK() 窗口函数需要 ORDER BY 子句")

        if partition_cols:
            grouped = df.groupby(partition_cols, sort=False)
        else:
            grouped = None

        def assign_dense_rank(group):
            sorted_group = group.sort_values(order_cols, ascending=ascending)
            # DENSE_RANK: 相同值取相同排名，下一个排名连续
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
        with_clause = parsed_sql.args.get('with_')
        if with_clause:
            # 复制worksheets_data避免修改原始数据，逐步添加CTE结果
            cte_data = dict(worksheets_data)
            for cte_expr in with_clause.expressions:
                cte_name = cte_expr.alias
                cte_query = cte_expr.this  # inner Select
                try:
                    # 每个CTE在已有的cte_data上执行（支持CTE引用前面的CTE）
                    cte_result = self._execute_query(cte_query, cte_data, limit=None)
                    cte_data[cte_name] = cte_result
                except Exception as e:
                    raise ValueError(f"CTE '{cte_name}' 执行失败: {e}")
            # 从parsed_sql中移除with_子句，让后续逻辑正常处理
            parsed_sql = parsed_sql.copy()
            parsed_sql.set('with_', None)

        # 获取FROM子句中的表名
        from_table = self._get_from_table(parsed_sql)

        # 查找表名时也搜索CTE定义的临时表
        effective_data = cte_data if with_clause else worksheets_data

        if from_table not in effective_data:
            raise ValueError(f"表 '{from_table}' 不存在。可用表: {list(effective_data.keys())}")

        base_df = effective_data[from_table].copy()

        # 构建表别名映射
        self._table_aliases = {}
        self._table_aliases[from_table] = from_table
        # 检查FROM子句是否有别名 (FROM 技能表 a)
        from_clause = parsed_sql.args.get('from')
        if from_clause:
            for alias in from_clause.find_all(exp.Alias):
                parent_table = from_clause.this.name if hasattr(from_clause.this, 'name') else str(from_clause.this)
                self._table_aliases[alias.alias] = parent_table
                self._table_aliases[parent_table] = parent_table

        # 应用JOIN子句
        joins = parsed_sql.args.get('joins')
        if joins:
            base_df = self._apply_join_clause(joins, base_df, effective_data, from_table)

        # 应用WHERE条件
        # 保存WHERE前的DataFrame，用于空结果智能建议
        base_df_before_where = base_df.copy()
        self._df_before_where = base_df_before_where
        # 保存当前工作表数据供子查询使用
        self._current_worksheets = effective_data
        base_df = self._apply_where_clause(parsed_sql, base_df)

        # 检查是否有聚合函数
        has_aggregate = self._check_has_aggregate_function(parsed_sql)

        # 应用GROUP BY和聚合
        if parsed_sql.args.get('group') or has_aggregate:
            # 有GROUP BY或有聚合函数时，应用分组聚合
            base_df = self._apply_group_by_aggregation(parsed_sql, base_df)

            # 应用HAVING条件
            if parsed_sql.args.get('having'):
                # 保存HAVING前的DataFrame，用于HAVING空结果建议
                self._df_before_having = base_df.copy()
                base_df = self._apply_having_clause(parsed_sql, base_df)

        # 应用窗口函数（ROW_NUMBER, RANK, DENSE_RANK）
        # 窗口函数在GROUP BY/HAVING之后、ORDER BY/SELECT之前计算
        base_df = self._apply_window_functions(parsed_sql, base_df)

        if parsed_sql.args.get('group') or has_aggregate:
            # ORDER BY（聚合查询：在GROUP BY之后）
            if parsed_sql.args.get('order'):
                base_df = self._apply_order_by(parsed_sql, base_df)
        else:
            # 非聚合查询：提取SELECT别名，然后ORDER BY（支持引用别名和原始列），最后SELECT
            select_aliases = self._extract_select_aliases(parsed_sql)
            if parsed_sql.args.get('order'):
                base_df = self._apply_order_by(parsed_sql, base_df, select_aliases=select_aliases)

            # 应用SELECT表达式（裁剪列、计算字段、别名）
            base_df = self._apply_select_expressions(parsed_sql, base_df)

        # 应用OFFSET（在LIMIT之前）
        offset_clause = parsed_sql.args.get('offset')
        if offset_clause:
            if hasattr(offset_clause, 'expression'):
                offset_value = int(offset_clause.expression.this)
            else:
                offset_value = int(offset_clause.this)
            base_df = base_df.iloc[offset_value:]

        # 应用LIMIT
        limit_clause = parsed_sql.args.get('limit')
        if limit_clause:
            if hasattr(limit_clause, 'expression'):
                limit_value = int(limit_clause.expression.this)
            else:
                limit_value = int(limit_clause.this)
            base_df = base_df.head(limit_value)
        elif limit:
            base_df = base_df.head(limit)

        # 应用SELECT DISTINCT去重
        if parsed_sql.args.get('distinct'):
            base_df = base_df.drop_duplicates()

        return base_df

    def _check_has_aggregate_function(self, parsed_sql: exp.Expression) -> bool:
        """检查SQL查询是否包含聚合函数"""
        for select_expr in parsed_sql.expressions:
            if self._is_aggregate_function(select_expr):
                return True
        return False

    def _apply_select_expressions(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """
        应用SELECT表达式（非聚合查询）
        处理计算字段、别名等

        Args:
            parsed_sql: 解析后的SQL表达式
            df: 数据DataFrame

        Returns:
            pd.DataFrame: 处理后的DataFrame
        """
        result_data = {}
        ordered_columns = []

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
            if isinstance(select_expr, exp.Alias):
                alias_name = select_expr.alias  # alias直接是字符串
                original_expr = select_expr.this
            else:
                original_expr = select_expr
                # 如果没有别名，使用原始列名
                if isinstance(original_expr, exp.Column):
                    alias_name = original_expr.name
                else:
                    alias_name = f"col_{i}"

            # 计算表达式值
            try:
                if isinstance(original_expr, exp.Column):
                    # 普通列引用（支持表限定符 a.column）
                    column_name = original_expr.name
                    table_part = original_expr.table if hasattr(original_expr, 'table') and original_expr.table else None
                    qualified = f"{table_part}.{column_name}" if table_part else None

                    if qualified and qualified in df.columns:
                        result_data[alias_name] = df[qualified]
                    elif column_name in df.columns:
                        result_data[alias_name] = df[column_name]
                    else:
                        suggestion = self._suggest_column_name(column_name, list(df.columns))
                        raise ValueError(f"列 '{qualified or column_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                elif isinstance(original_expr, exp.Case):
                    # CASE WHEN表达式
                    result_data[alias_name] = self._evaluate_case_expression(original_expr, df)

                elif isinstance(original_expr, exp.Coalesce):
                    # COALESCE/IFNULL表达式
                    results = []
                    for idx in range(len(df)):
                        row = df.iloc[idx]
                        results.append(self._evaluate_coalesce_for_row(original_expr, row))
                    result_data[alias_name] = pd.Series(results, index=df.index)

                elif self._is_string_function(original_expr):
                    # 字符串函数: UPPER, LOWER, TRIM, LENGTH, CONCAT, REPLACE, SUBSTRING, LEFT, RIGHT
                    result_data[alias_name] = self._evaluate_string_function(original_expr, df)

                elif isinstance(original_expr, exp.Window):
                    # 窗口函数: ROW_NUMBER, RANK, DENSE_RANK（已由_apply_window_functions预计算）
                    if alias_name in df.columns:
                        result_data[alias_name] = df[alias_name]
                    else:
                        raise ValueError(f"窗口函数结果列 '{alias_name}' 未找到")

                elif self._is_mathematical_expression(original_expr):
                    # 数学表达式
                    result_data[alias_name] = self._evaluate_math_expression(original_expr, df)

                elif isinstance(original_expr, exp.Literal):
                    # SELECT中的字面量值（如 SELECT 1, SELECT 'hello'）
                    val = original_expr.this
                    if original_expr.is_string:
                        result_data[alias_name] = pd.Series([val] * len(df), index=df.index)
                    else:
                        try:
                            num_val = int(val) if '.' not in str(val) else float(val)
                            result_data[alias_name] = pd.Series([num_val] * len(df), index=df.index)
                        except (ValueError, TypeError):
                            result_data[alias_name] = pd.Series([val] * len(df), index=df.index)

                else:
                    # 其他表达式，尝试作为列处理
                    if hasattr(original_expr, 'name') and original_expr.name in df.columns:
                        result_data[alias_name] = df[original_expr.name]
                    else:
                        raise ValueError(f"不支持的表达式: {original_expr}")

                ordered_columns.append(alias_name)

            except Exception as e:
                # 表达式处理失败，尝试返回原始值
                if hasattr(original_expr, 'name') and original_expr.name in df.columns:
                    result_data[alias_name] = df[original_expr.name]
                    ordered_columns.append(alias_name)
                else:
                    raise ValueError(f"处理SELECT表达式失败: {e}")

        # 构建结果DataFrame，保持SELECT顺序
        if result_data:
            result_df = pd.DataFrame(result_data)
            # 按照SQL SELECT顺序重新排列列
            result_df = result_df[ordered_columns]
            return result_df
        else:
            return df

    def _is_mathematical_expression(self, expr) -> bool:
        """检查是否为数学表达式"""
        return isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div, exp.Mod))

    # 数学运算符分发表：二元运算符统一处理
    _MATH_BINARY_OPS = {
        exp.Add: operator.add,
        exp.Sub: operator.sub,
        exp.Mul: operator.mul,
        exp.Div: operator.truediv,
        exp.Mod: operator.mod,
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
            # COALESCE在数学表达式中
            results = []
            for idx in range(len(df)):
                row = df.iloc[idx]
                results.append(self._evaluate_coalesce_for_row(expr, row))
            return pd.Series(results, index=df.index)
        else:
            raise ValueError(f"不支持的数学表达式部分: {expr}")

    def _is_string_function(self, expr) -> bool:
        """检查是否为字符串函数"""
        return isinstance(expr, (exp.Upper, exp.Lower, exp.Trim, exp.Length,
                                exp.Concat, exp.Replace, exp.Substring, exp.Left, exp.Right))

    # 简单字符串函数分发表：一元操作，统一模式 val_series.astype(str).str.<op>()
    _SIMPLE_STR_OPS = {
        exp.Upper: 'upper',
        exp.Lower: 'lower',
        exp.Trim: 'strip',
        exp.Length: 'len',
    }

    def _evaluate_string_function(self, expr, df: pd.DataFrame) -> pd.Series:
        """计算字符串函数，返回pd.Series"""
        func_type = type(expr)

        # 简单字符串函数：分发表处理
        if func_type in self._SIMPLE_STR_OPS:
            val_series = self._expr_to_series(expr.this, df)
            return getattr(val_series.astype(str).str, self._SIMPLE_STR_OPS[func_type])()

        func_name = func_type.__name__.lower()

        if func_name == 'concat':
            # CONCAT(a, b, ...) — expressions列表包含所有参数
            parts = [self._expr_to_series(arg, df).astype(str) for arg in expr.expressions]
            if parts:
                result = parts[0]
                for p in parts[1:]:
                    result = result + p
                return result
            return pd.Series([''] * len(df), index=df.index)

        elif func_name == 'replace':
            # REPLACE(str, old, new) — sqlglot: this=string, expression=old, replacement=new
            val_series = self._expr_to_series(expr.this, df).astype(str)
            old_val = str(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else ''
            new_val = str(self._literal_value(expr.args.get('replacement'))) if expr.args.get('replacement') else ''
            return val_series.str.replace(old_val, new_val, regex=False)

        elif func_name in ('substring', 'left', 'right'):
            val_series = self._expr_to_series(expr.this, df).astype(str)
            if func_name == 'substring':
                # sqlglot: this=string, start=1-based, length=count
                start = int(self._literal_value(expr.args.get('start'))) - 1 if expr.args.get('start') else 0
                length = int(self._literal_value(expr.args.get('length'))) if expr.args.get('length') else len(val_series.iloc[0])
                return val_series.str.slice(start, start + length)
            elif func_name == 'left':
                # sqlglot: this=string, expression=count
                n = int(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else 1
                return val_series.str.slice(0, n)
            elif func_name == 'right':
                # sqlglot: this=string, expression=count
                n = int(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else 1
                return val_series.str.slice(-n)

        raise ValueError(f"不支持的字符串函数: {func_name}")

    def _evaluate_string_function_for_row(self, expr, row: pd.Series) -> Any:
        """逐行评估字符串函数"""
        func_name = type(expr).__name__.lower()
        val = self._get_row_value(expr.this, row)
        if val is None:
            return None
        val = str(val)

        if func_name == 'upper':
            return val.upper()
        elif func_name == 'lower':
            return val.lower()
        elif func_name == 'trim':
            return val.strip()
        elif func_name == 'length':
            return len(val)
        elif func_name == 'concat':
            parts = [str(self._get_row_value(arg, row) or '') for arg in expr.expressions]
            return ''.join(parts)
        elif func_name == 'replace':
            old_val = str(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else ''
            new_val = str(self._literal_value(expr.args.get('replacement'))) if expr.args.get('replacement') else ''
            return val.replace(old_val, new_val)
        elif func_name == 'substring':
            start = int(self._literal_value(expr.args.get('start'))) - 1 if expr.args.get('start') else 0
            length = int(self._literal_value(expr.args.get('length'))) if expr.args.get('length') else len(val)
            return val[start:start + length]
        elif func_name == 'left':
            n = int(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else 1
            return val[:n]
        elif func_name == 'right':
            n = int(self._literal_value(expr.args.get('expression'))) if expr.args.get('expression') else 1
            return val[-n:] if n > 0 else ''
        return val

    def _expr_to_series(self, expr, df: pd.DataFrame) -> pd.Series:
        """将表达式转换为pd.Series（支持列引用、字面量、数学表达式、字符串函数）"""
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
            results = [self._evaluate_coalesce_for_row(expr, df.iloc[i]) for i in range(len(df))]
            return pd.Series(results, index=df.index)
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

    def _get_from_table(self, parsed_sql: exp.Expression) -> str:
        """获取FROM子句中的表名"""
        from_clause = parsed_sql.args.get('from')
        if not from_clause:
            # 尝试使用 from_ 键（sqlglot的另一种存储方式）
            from_clause = parsed_sql.args.get('from_')
        if from_clause:
            # 检查FROM子句是否是子查询（FROM (SELECT ...)）
            if hasattr(from_clause, 'this') and isinstance(from_clause.this, (exp.Subquery, exp.Select)):
                raise ValueError("不支持FROM子查询（FROM (SELECT ...)）。请使用子查询作为WHERE条件：WHERE col IN (SELECT ...)")
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'name'):
                return from_clause.this.name
            # 兼容 Table 对象
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'this'):
                return from_clause.this.this

        # 如果没有明确的FROM子句，返回第一个表名
        raise ValueError("无法确定FROM子句中的表名")

    def _apply_join_clause(self, joins, left_df: pd.DataFrame, worksheets_data: Dict[str, pd.DataFrame], left_table: str) -> pd.DataFrame:
        """
        应用JOIN子句，支持INNER/LEFT/RIGHT/FULL/CROSS JOIN

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

        result_df = left_df

        for join in joins:
            # 解析JOIN类型
            join_kind = 'inner'
            join_side = join.side  # LEFT, RIGHT, etc.
            join_kind_name = join.kind  # INNER, CROSS, etc.

            if join_side and str(join_side).upper() == 'LEFT':
                join_kind = 'left'
            elif join_side and str(join_side).upper() == 'RIGHT':
                join_kind = 'right'
            elif (join_side and str(join_side).upper() == 'FULL') or \
                 (join_kind_name and str(join_kind_name).upper() == 'FULL'):
                join_kind = 'outer'
            elif join_kind_name and str(join_kind_name).upper() == 'CROSS':
                join_kind = 'cross'

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
            for alias in join.find_all(exp.Alias):
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
                raise ValueError(f"JOIN表 '{right_table}' 不存在。可用表: {available}")

            right_df = worksheets_data[right_table].copy()

            # 解析ON条件（CROSS JOIN不需要ON）
            on_clause = join.args.get('on')
            left_on_col = None
            right_on_col = None
            actual_right_on = None

            if join_kind == 'cross':
                # CROSS JOIN: 笛卡尔积，不需要ON条件
                pass
            elif not on_clause:
                raise ValueError("JOIN缺少ON条件")
            else:
                left_on_col, right_on_col = self._parse_join_on_condition(on_clause, left_table, right_table, right_alias)

                # 验证列存在
                if left_on_col not in result_df.columns:
                    raise ValueError(f"JOIN ON条件: 左表 '{left_table}' 没有列 '{left_on_col}'。可用列: {list(result_df.columns)}")
                if right_on_col not in right_df.columns:
                    raise ValueError(f"JOIN ON条件: 右表 '{right_table}' 没有列 '{right_on_col}'。可用列: {list(right_df.columns)}")

            # 执行JOIN
            # 为右表列添加别名前缀避免冲突
            right_df_renamed = right_df.copy()
            col_mapping = {}
            for col in right_df_renamed.columns:
                if col in result_df.columns and (left_on_col is None or col != left_on_col):
                    new_col = f"{right_alias}.{col}"
                    col_mapping[col] = new_col
                elif left_on_col and col == left_on_col and left_on_col == right_on_col:
                    # ON列同名：右表列重命名避免合并后重复
                    new_col = f"{right_alias}.{col}"
                    col_mapping[col] = new_col
            right_df_renamed = right_df_renamed.rename(columns=col_mapping)

            # 调整右表ON列名（如果被重命名了）
            if right_on_col:
                actual_right_on = col_mapping.get(right_on_col, right_on_col)

            # 合并双行表头描述
            if right_table in self._header_descriptions:
                for orig_col, new_col in col_mapping.items():
                    if orig_col in self._header_descriptions[right_table]:
                        self._header_descriptions[right_table][new_col] = self._header_descriptions[right_table][orig_col]

            if join_kind == 'cross':
                # CROSS JOIN: 笛卡尔积（无需ON列）
                result_df = result_df.merge(right_df_renamed, how='cross')
            else:
                result_df = result_df.merge(
                    right_df_renamed,
                    left_on=left_on_col,
                    right_on=actual_right_on,
                    how=join_kind
                )

            # 合并后删除重复的ON列（右表侧）
            if actual_right_on and actual_right_on in result_df.columns and actual_right_on != left_on_col:
                result_df = result_df.drop(columns=[actual_right_on])

        return result_df

    def _parse_join_on_condition(self, on_clause, left_table: str, right_table: str, right_alias: str) -> Tuple[str, str]:
        """
        解析JOIN ON条件，返回(左列, 右列)

        支持格式：
        - ON a.id = b.id（带表别名）
        - ON id = id（不带别名，自动匹配）
        - ON a.id = id（混合）
        """
        if isinstance(on_clause, exp.EQ):
            left_expr = on_clause.left
            right_expr = on_clause.right
        elif isinstance(on_clause, exp.And):
            # 多条件JOIN暂不支持，取第一个等值条件
            for child in on_clause.find_all(exp.EQ):
                left_expr = child.left
                right_expr = child.right
                break
            else:
                raise ValueError("JOIN ON多条件暂不支持，请使用单个等值连接条件")
        else:
            raise ValueError(f"JOIN ON条件格式不支持，请使用等值连接: ON a.id = b.id")

        def resolve_column(col_expr) -> str:
            if isinstance(col_expr, exp.Column):
                col_name = col_expr.name
                # 检查是否有表限定符
                table_part = col_expr.table if hasattr(col_expr, 'table') and col_expr.table else None
                return col_name, table_part
            return str(col_expr), None

        left_col, left_tbl = resolve_column(left_expr)
        right_col, right_tbl = resolve_column(right_expr)

        # 判断哪个属于左表，哪个属于右表
        if left_tbl:
            resolved_left_tbl = self._table_aliases.get(left_tbl, left_tbl)
            if resolved_left_tbl == right_table or left_tbl == right_alias:
                # 左表达式实际指向右表，交换
                return right_col, left_col
        if right_tbl:
            resolved_right_tbl = self._table_aliases.get(right_tbl, right_tbl)
            if resolved_right_tbl == left_table:
                # 右表达式实际指向左表，交换
                return right_col, left_col

        # 无表限定符：根据列是否存在于左右表来判断
        # 默认左=左表, 右=右表
        return left_col, right_col

    def _apply_where_clause(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用WHERE条件"""
        where_clause = parsed_sql.args.get('where')
        if not where_clause:
            return df

        # 如果WHERE包含复杂表达式（pandas query不支持的类型），直接使用逐行过滤
        where_expr = where_clause.this
        # 集合检查：新增复杂表达式只需在集合中添加一行
        _COMPLEX_EXPR_TYPES = {
            exp.Coalesce, exp.Case, exp.Exists,
            exp.Upper, exp.Lower, exp.Trim, exp.Length,
            exp.Concat, exp.Replace, exp.Substring, exp.Left, exp.Right,
        }
        has_complex = any(where_expr.find(t) is not None for t in _COMPLEX_EXPR_TYPES)

        if has_complex:
            return self._apply_row_filter(where_expr, df)

        # 将SQLGlot表达式转换为pandas查询条件
        condition_str = self._sql_condition_to_pandas(where_expr, df)

        if condition_str:
            try:
                return df.query(condition_str)
            except Exception as e:
                # 如果查询失败，尝试逐行过滤
                return self._apply_row_filter(where_clause.this, df)

        return df

    def _sql_condition_to_pandas(self, condition: exp.Expression, df: pd.DataFrame) -> str:
        """将SQL条件转换为pandas查询字符串"""
        # 比较运算符分发表 → pandas查询字符串
        _PANDAS_OPS = {
            exp.EQ: '==', exp.NEQ: '!=',
            exp.GT: '>', exp.GTE: '>=',
            exp.LT: '<', exp.LTE: '<=',
        }
        op_type = type(condition)
        if op_type in _PANDAS_OPS:
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} {_PANDAS_OPS[op_type]} {right}"

        elif isinstance(condition, exp.And):
            left = self._sql_condition_to_pandas(condition.left, df)
            right = self._sql_condition_to_pandas(condition.right, df)
            return f"({left}) & ({right})"

        elif isinstance(condition, exp.Or):
            left = self._sql_condition_to_pandas(condition.left, df)
            right = self._sql_condition_to_pandas(condition.right, df)
            return f"({left}) | ({right})"

        elif isinstance(condition, exp.Paren):
            # 括号表达式，直接处理内部表达式
            return self._sql_condition_to_pandas(condition.this, df)

        elif isinstance(condition, exp.Not):
            inner = condition.this
            # NOT LIKE: NOT > Like → 排除匹配行
            if isinstance(inner, exp.Like):
                left = self._expression_to_column_reference(inner.this, df)
                right = self._expression_to_value(inner.expression, df)
                pattern = str(right).strip("'\"")
                pattern = pattern.replace('%', '.*').replace('_', '.')
                return f"~{left}.str.match('{pattern}', case=False, na=False)"
            # NOT IN: NOT > In → 排除列表中值
            if isinstance(inner, exp.In):
                left = self._expression_to_column_reference(inner.this, df)
                # 检查是否有子查询
                subquery = inner.args.get('query')
                if subquery and isinstance(subquery, exp.Subquery):
                    try:
                        sub_result = self._execute_subquery(subquery, self._current_worksheets)
                        if len(sub_result.columns) > 0:
                            sub_values = sub_result.iloc[:, 0].dropna().tolist()
                            values_str = ', '.join(repr(v) for v in sub_values)
                            return f"~{left}.isin([{values_str}])"
                        return f"~{left}.isin([])"
                    except Exception as e:
                        raise ValueError(f"NOT IN子查询执行失败: {e}")
                values = []
                for value in inner.expressions:
                    values.append(self._expression_to_value(value, df))
                values_str = ', '.join(str(v) for v in values)
                return f"~{left}.isin([{values_str}])"
            # 其他NOT表达式（IS NOT NULL等）
            pandas_expr = self._sql_condition_to_pandas(inner, df)
            return f"~({pandas_expr})"

        elif isinstance(condition, exp.Like):
            left = self._expression_to_column_reference(condition.this, df)
            right = self._expression_to_value(condition.expression, df)
            pattern = str(right).strip("'\"")
            # 转换SQL LIKE模式为pandas str.contains（默认大小写不敏感，适配游戏配置表场景）
            pattern = pattern.replace('%', '.*').replace('_', '.')
            return f"{left}.str.match('{pattern}', case=False, na=False)"

        elif isinstance(condition, exp.In):
            left_col_name = condition.this.name if isinstance(condition.this, exp.Column) else None
            # 检查是否有子查询 (IN (SELECT ...))
            subquery = condition.args.get('query')
            if subquery and isinstance(subquery, exp.Subquery):
                left = self._expression_to_column_reference(condition.this, df)
                try:
                    sub_result = self._execute_subquery(subquery, self._current_worksheets)
                    if len(sub_result.columns) > 0:
                        sub_values = sub_result.iloc[:, 0].dropna().tolist()
                        values_str = ', '.join(repr(v) for v in sub_values)
                        return f"{left}.isin([{values_str}])"
                    return f"{left}.isin([])"
                except Exception as e:
                    raise ValueError(f"IN子查询执行失败: {e}")
            else:
                left = self._expression_to_column_reference(condition.this, df)
                values = []
                for value in condition.expressions:
                    values.append(self._expression_to_value(value, df))
                values_str = ', '.join(str(v) for v in values)
                return f"{left}.isin([{values_str}])"

        # EXISTS (子查询)
        elif isinstance(condition, exp.Exists):
            # EXISTS需要逐行评估（特别是关联子查询），返回None触发行过滤回退
            return None

        # IS NULL (sqlglot解析为 exp.Is)
        elif isinstance(condition, exp.Is):
            left = self._expression_to_column_reference(condition.this, df)
            # IS NOT NULL会被解析为 Not > Is
            return f"{left}.isna()"

        # BETWEEN x AND y
        elif isinstance(condition, exp.Between):
            left = self._expression_to_column_reference(condition.this, df)
            low = self._expression_to_value(condition.args['low'], df)
            high = self._expression_to_value(condition.args['high'], df)
            return f"({left} >= {low}) & ({left} <= {high})"

        else:
            raise ValueError(f"不支持的条件类型: {type(condition)}")

    def _expression_to_column_reference(self, expr: exp.Expression, df: pd.DataFrame) -> str:
        """将表达式转换为列引用（支持表限定符 a.column）"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            # 处理表限定符 (a.column_name → 查找 "a.column_name" 或 "column_name")
            table_part = expr.table if hasattr(expr, 'table') and expr.table else None
            qualified = f"{table_part}.{col_name}" if table_part else None

            if qualified and qualified in df.columns:
                return f"`{qualified}`"
            if col_name in df.columns:
                return f"`{col_name}`"
            # 兜底：尝试搜索别名前缀匹配
            if table_part and hasattr(self, '_table_aliases'):
                for col in df.columns:
                    if col == f"{table_part}.{col_name}":
                        return f"`{col}`"
            suggestion = self._suggest_column_name(col_name, list(df.columns))
            raise ValueError(f"列 '{qualified or col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

        elif isinstance(expr, exp.Literal):
            return str(expr.this)

        elif isinstance(expr, exp.AggFunc):
            # 对于HAVING子句中的聚合函数，需要查找对应的列
            func_name = type(expr).__name__.lower()

            # 优先：通过SELECT别名映射（HAVING COUNT(*) → SELECT COUNT(*) as cnt）
            if hasattr(self, '_having_agg_alias_map') and self._having_agg_alias_map:
                agg_sql = expr.sql()
                for map_sql, alias in self._having_agg_alias_map.items():
                    if agg_sql == map_sql:
                        if alias in df.columns:
                            return f"`{alias}`"
                        break

            # 精确匹配：查找列名等于函数名的列
            if func_name in df.columns:
                return f"`{func_name}`"
            
            # 模糊匹配：查找列名包含函数名的列
            for col in df.columns:
                if func_name in col.lower():
                    return f"`{col}`"
            
            # 对于COUNT(*)，尝试查找包含"count"的列
            if func_name == 'count':
                for col in df.columns:
                    if 'count' in col.lower():
                        return f"`{col}`"
            
            # 如果只有一个数值列，就使用它（常见于全表聚合）
            numeric_cols = []
            for col in df.columns:
                try:
                    pd.to_numeric(df[col], errors='coerce')
                    numeric_cols.append(col)
                except Exception:
                    pass
            
            if len(numeric_cols) == 1:
                return f"`{numeric_cols[0]}`"
            
            # 如果有多个列，尝试返回第一个（作为后备方案）
            if len(df.columns) > 0:
                return f"`{df.columns[0]}`"

            # 如果没有找到匹配的列，抛出错误
            raise ValueError(f"无法找到聚合函数 {func_name} 对应的列。可用列: {list(df.columns)}")

        else:
            raise ValueError(f"不支持的表达式类型: {type(expr)}")

    def _expression_to_value(self, expr: exp.Expression, df: pd.DataFrame) -> Union[str, int, float]:
        """将表达式转换为值"""
        if isinstance(expr, exp.Literal):
            value = expr.this
            if isinstance(value, str):
                # 尝试将字符串转换为数字
                try:
                    # 检查是否为整数
                    if value.isdigit():
                        return int(value)
                    # 检查是否为浮点数
                    elif '.' in value and value.replace('.', '').isdigit():
                        return float(value)
                    # 如果不是数字，保持字符串格式
                    else:
                        return f"'{value}'"
                except (ValueError, AttributeError):
                    return f"'{value}'"
            return value

        elif isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name not in df.columns:
                suggestion = self._suggest_column_name(col_name, list(df.columns))
                raise ValueError(f"列 '{col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")
            return f"`{col_name}`"

        elif isinstance(expr, exp.AggFunc):
            # 聚合函数作为值的处理（HAVING子句中）
            # 使用与_expression_to_column_reference相同的逻辑
            return self._expression_to_column_reference(expr, df)

        elif isinstance(expr, exp.Subquery):
            # 标量子查询: WHERE col > (SELECT AVG(...) FROM ...)
            try:
                sub_result = self._execute_subquery(expr, self._current_worksheets)
                if len(sub_result) > 0 and len(sub_result.columns) > 0:
                    scalar_val = sub_result.iloc[0, 0]
                    if isinstance(scalar_val, (int, float, np.integer, np.floating)):
                        return float(scalar_val)
                    return f"'{scalar_val}'"
                return "0"
            except Exception as e:
                raise ValueError(f"标量子查询执行失败: {e}")

        else:
            raise ValueError(f"不支持的表达式类型: {type(expr)}")

    def _apply_row_filter(self, condition: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """逐行应用过滤条件（备用方案）"""
        mask = []

        for _, row in df.iterrows():
            if self._evaluate_condition_for_row(condition, row):
                mask.append(True)
            else:
                mask.append(False)

        return df[mask]

    def _evaluate_condition_for_row(self, condition: exp.Expression, row: pd.Series) -> bool:
        """为单行评估条件"""
        try:
            # 比较运算符分发表（EQ/NEQ直接比较，GT/GTE/LT/LTE数值比较）
            _COMPARISON_OPS = {
                exp.EQ: lambda l, r: l == r,
                exp.NEQ: lambda l, r: l != r,
                exp.GT: lambda l, r: float(l) > float(r),
                exp.GTE: lambda l, r: float(l) >= float(r),
                exp.LT: lambda l, r: float(l) < float(r),
                exp.LTE: lambda l, r: float(l) <= float(r),
            }
            op_type = type(condition)
            if op_type in _COMPARISON_OPS:
                left_val = self._get_row_value(condition.left, row)
                right_val = self._get_row_value(condition.right, row)
                try:
                    return _COMPARISON_OPS[op_type](left_val, right_val)
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
                pattern = pattern.replace('%', '.*').replace('_', '.')
                return bool(re.match(pattern, val, re.IGNORECASE))

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
                # EXISTS子查询（支持关联子查询）
                inner_expr = condition.this  # 可能是Select或Subquery
                if isinstance(inner_expr, exp.Subquery):
                    inner_select = inner_expr.this
                elif isinstance(inner_expr, exp.Select):
                    inner_select = inner_expr
                else:
                    return False

                if hasattr(self, '_current_worksheets') and self._current_worksheets:
                    # 检查是否有关联引用（引用外部表列）
                    inner_from = self._get_from_table(inner_select)
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
                            # 无表限定符：检查列名是否存在于外部表但不在子查询表
                            if inner_from in self._current_worksheets:
                                inner_cols = set(self._current_worksheets[inner_from].columns)
                                # 检查是否是外部表独有的列（不在子查询FROM表中）
                                for tbl_name, tbl_df in self._current_worksheets.items():
                                    if tbl_name != inner_from and col_name in tbl_df.columns:
                                        has_correlation = True
                                        break
                    
                    if has_correlation:
                        # 关联子查询：替换外部引用为当前行值，然后执行
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
                                # 无表限定符：检查列是否只存在于外部表
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
                                        # 无表限定符：精确替换（避免误替换子查询表的列）
                                        pattern = r'\b' + re.escape(col_name) + r'\b'
                                        inner_sql = re.sub(pattern, repr(val), inner_sql, count=1)
                        try:
                            parsed_inner = sqlglot.parse_one(inner_sql)
                            sub_result = self._execute_query(parsed_inner, self._current_worksheets)
                            return len(sub_result) > 0
                        except Exception:
                            return False
                    else:
                        sub_result = self._execute_subquery(inner_expr, self._current_worksheets)
                        return len(sub_result) > 0
                return False

            # 其他条件类型...

            return True

        except Exception:
            return False

    def _get_row_value(self, expr: exp.Expression, row: pd.Series) -> Any:
        """获取行中表达式的值"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            return row.get(col_name)

        elif isinstance(expr, exp.Literal):
            return expr.this

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

    def _apply_group_by_aggregation(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用GROUP BY和聚合函数"""
        group_by_columns = []
        group_clause = parsed_sql.args.get('group')
        if group_clause:
            for group_expr in group_clause.expressions:
                if isinstance(group_expr, exp.Column):
                    group_by_columns.append(group_expr.name)

        # 检查是否有聚合函数
        aggregations = {}
        select_exprs = {}

        for i, select_expr in enumerate(parsed_sql.expressions):
            if isinstance(select_expr, exp.Alias):
                alias_name = select_expr.alias  # alias直接是字符串
                original_expr = select_expr.this
                select_exprs[alias_name] = original_expr
            else:
                alias_name = f"col_{i}"
                select_exprs[alias_name] = select_expr

        # 检查聚合函数
        for alias_name, expr in select_exprs.items():
            if self._is_aggregate_function(expr):
                aggregations[alias_name] = expr
            elif hasattr(expr, 'name') and expr.name not in group_by_columns:
                # 如果是非聚合列且不在GROUP BY中，需要添加到GROUP BY
                if isinstance(expr, (exp.Column, exp.Identifier)):
                    group_by_columns.append(expr.name)

        if not aggregations:
            # 没有聚合函数，只应用GROUP BY去重
            if group_by_columns:
                return df[group_by_columns].drop_duplicates().reset_index(drop=True)
            else:
                return df

        # 应用聚合
        if group_by_columns:
            grouped = df.groupby(group_by_columns)
        else:
            # 全表聚合
            grouped = df.groupby(lambda x: 0)  # 将所有行分组为一组

        # 按照SQL SELECT表达式的顺序构建结果
        result_data = {}
        ordered_columns = []

        # 按SELECT表达式顺序处理列
        for i, select_expr in enumerate(parsed_sql.expressions):
            if isinstance(select_expr, exp.Alias):
                alias_name = select_expr.alias
                original_expr = select_expr.this
            else:
                # 对于没有别名的表达式，生成有意义的别名
                if self._is_aggregate_function(select_expr):
                    # 聚合函数无别名：生成如 count_star, avg_damage, sum_hp 等
                    alias_name = self._generate_aggregate_alias(select_expr)
                elif hasattr(select_expr, 'name') and select_expr.name:
                    alias_name = select_expr.name
                else:
                    alias_name = f"col_{i}"
                original_expr = select_expr

            ordered_columns.append(alias_name)

            # 处理聚合函数
            # 检查当前表达式（或其内部表达式）是否是聚合函数
            is_agg = self._is_aggregate_function(select_expr if not isinstance(select_expr, exp.Alias) else select_expr.this)
            
            if is_agg:
                # 找到对应的聚合表达式
                agg_expr = original_expr if isinstance(select_expr, exp.Alias) else select_expr
                agg_result = self._apply_aggregation_function(agg_expr, grouped)
                # 如果结果是标量，转换为Series
                if isinstance(agg_result, (int, float, np.integer, np.floating)):
                    result_data[alias_name] = pd.Series([agg_result])
                else:
                    result_data[alias_name] = agg_result
            # 处理CASE WHEN表达式（在GROUP BY中作为列使用）
            elif isinstance(original_expr, exp.Case):
                # CASE WHEN在GROUP BY中需要先计算再分组
                case_series = self._evaluate_case_expression(original_expr, df)
                if alias_name not in group_by_columns:
                    group_by_columns.append(alias_name)
                # 重新分组包含CASE WHEN列
                if len(group_by_columns) > 1:
                    existing_cols = [c for c in group_by_columns if c != alias_name and c in df.columns]
                    if existing_cols:
                        temp_df = df.copy()
                        temp_df[alias_name] = case_series
                        regrouped = temp_df.groupby(group_by_columns)
                        result_data[alias_name] = regrouped[alias_name].first()
                    else:
                        result_data[alias_name] = case_series.groupby(case_series).first()
                else:
                    result_data[alias_name] = case_series.groupby(case_series).first()
            # 处理COALESCE表达式
            elif isinstance(original_expr, exp.Coalesce):
                coalesce_results = []
                for idx in range(len(df)):
                    row = df.iloc[idx]
                    coalesce_results.append(self._evaluate_coalesce_for_row(original_expr, row))
                coalesce_series = pd.Series(coalesce_results)
                if alias_name not in group_by_columns:
                    group_by_columns.append(alias_name)
                result_data[alias_name] = coalesce_series.groupby(coalesce_series).first()
            # 处理普通列（GROUP BY列）
            elif hasattr(original_expr, 'name'):
                col_name = original_expr.name
                if col_name in group_by_columns:
                    result_data[alias_name] = grouped[col_name].first()
            # 处理SELECT * 的情况
            elif alias_name == '*':
                # SELECT * 情况，返回所有GROUP BY列
                for col in group_by_columns:
                    if col not in result_data:
                        result_data[col] = grouped[col].first()

        # 组合结果，保持列顺序
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
        执行子查询，返回结果DataFrame

        Args:
            subquery_expr: sqlglot Subquery或Select表达式
            worksheets_data: 当前可用的所有工作表数据

        Returns:
            pd.DataFrame: 子查询结果
        """
        # sqlglot可能将子查询直接存储为Select（而非Subquery包装）
        if isinstance(subquery_expr, exp.Subquery):
            inner_select = subquery_expr.this
        elif isinstance(subquery_expr, exp.Select):
            inner_select = subquery_expr
        else:
            raise ValueError(f"不支持子查询类型: {type(subquery_expr)}")

        # 获取子查询的FROM表
        from_table = self._get_from_table(inner_select)
        if from_table not in worksheets_data:
            raise ValueError(f"子查询中表 '{from_table}' 不存在。可用表: {list(worksheets_data.keys())}")

        # 复用现有查询执行逻辑
        try:
            result = self._execute_query(inner_select, worksheets_data)
            return result
        except Exception as e:
            raise ValueError(f"子查询执行失败: {e}")

    def _evaluate_case_expression(self, case_expr: exp.Case, df: pd.DataFrame, row: Optional[pd.Series] = None) -> Any:
        """
        评估CASE WHEN表达式

        支持格式：
        - CASE WHEN cond1 THEN val1 WHEN cond2 THEN val2 ELSE default END
        - 单行评估模式(row参数)或向量化模式(无row参数)

        Args:
            case_expr: sqlglot Case表达式
            df: DataFrame（向量化模式）
            row: 可选，单行数据（逐行模式）

        Returns:
            向量化模式返回pd.Series，逐行模式返回单个值
        """
        ifs = case_expr.args.get('ifs', [])
        default_value = case_expr.args.get('default')

        if row is not None:
            # 逐行评估模式
            for if_clause in ifs:
                condition = if_clause.this
                if self._evaluate_condition_for_row(condition, row):
                    return self._get_row_value(if_clause.args.get('true'), row)
            # 没有匹配的WHEN，返回ELSE默认值
            if default_value is not None:
                if isinstance(default_value, exp.Literal):
                    return default_value.this if default_value.is_string else (
                        float(default_value.this) if '.' in str(default_value.this) else int(default_value.this)
                    )
                elif isinstance(default_value, exp.Column):
                    return row.get(default_value.name)
                return None
            return None
        else:
            # 向量化模式 - 逐行构建Series
            results = []
            for idx in range(len(df)):
                row = df.iloc[idx]
                matched = False
                for if_clause in ifs:
                    condition = if_clause.this
                    try:
                        if self._evaluate_condition_for_row(condition, row):
                            true_expr = if_clause.args.get('true')
                            val = self._get_expression_value(true_expr, row)
                            results.append(val)
                            matched = True
                            break
                    except Exception:
                        continue
                if not matched:
                    if default_value is not None:
                        if isinstance(default_value, exp.Literal):
                            results.append(default_value.this if default_value.is_string else (
                                float(default_value.this) if '.' in str(default_value.this) else int(default_value.this)
                            ))
                        elif isinstance(default_value, exp.Column):
                            results.append(row.get(default_value.name))
                        else:
                            results.append(None)
                    else:
                        results.append(None)
            return pd.Series(results, index=df.index)

    def _get_expression_value(self, expr: exp.Expression, row: pd.Series) -> Any:
        """获取表达式在指定行的值（支持列引用、字面量、算术表达式）"""
        if isinstance(expr, exp.Literal):
            return expr.this if expr.is_string else (
                float(expr.this) if '.' in str(expr.this) else int(expr.this)
            )
        elif isinstance(expr, exp.Column):
            return row.get(expr.name)
        elif isinstance(expr, (exp.Add, exp.Sub, exp.Mul, exp.Div)):
            return self._evaluate_math_for_row(expr, row)
        elif isinstance(expr, exp.Coalesce):
            return self._evaluate_coalesce_for_row(expr, row)
        return None

    def _evaluate_math_for_row(self, expr: exp.Expression, row: pd.Series) -> Any:
        """逐行评估数学表达式"""
        if isinstance(expr, exp.Add):
            left = self._get_expression_value(expr.left, row)
            right = self._get_expression_value(expr.right, row)
            try:
                return float(left) + float(right)
            except (TypeError, ValueError):
                return None
        elif isinstance(expr, exp.Sub):
            left = self._get_expression_value(expr.left, row)
            right = self._get_expression_value(expr.right, row)
            try:
                return float(left) - float(right)
            except (TypeError, ValueError):
                return None
        elif isinstance(expr, exp.Mul):
            left = self._get_expression_value(expr.left, row)
            right = self._get_expression_value(expr.right, row)
            try:
                return float(left) * float(right)
            except (TypeError, ValueError):
                return None
        elif isinstance(expr, exp.Div):
            left = self._get_expression_value(expr.left, row)
            right = self._get_expression_value(expr.right, row)
            try:
                r = float(right)
                return float(left) / r if r != 0 else None
            except (TypeError, ValueError):
                return None
        return None

    def _evaluate_coalesce_for_row(self, coalesce_expr: exp.Coalesce, row: pd.Series) -> Any:
        """逐行评估COALESCE/IFNULL表达式"""
        # COALESCE结构: this=第一个参数, expressions=[后续参数]
        values = [coalesce_expr.this] + list(coalesce_expr.expressions)
        for val_expr in values:
            val = self._get_expression_value(val_expr, row)
            if val is not None and not (isinstance(val, float) and np.isnan(val)):
                return val
        return None


    def _generate_aggregate_alias(self, expr: exp.Expression) -> str:
        """为无别名的聚合函数生成有意义的列名

        例: COUNT(*) → count_star, AVG(damage) → avg_damage, SUM(hp) → sum_hp
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

    def _apply_aggregation_function(self, expr: exp.Expression, grouped) -> pd.Series:
        """应用聚合函数"""
        if isinstance(expr, exp.AggFunc):
            # 从聚合函数类型获取函数名
            func_name = type(expr).__name__.lower()

            # 应用对应的聚合函数
            if func_name == 'count':
                if isinstance(expr.this, exp.Star):
                    # COUNT(*)的情况
                    return grouped.size()
                elif isinstance(expr.this, exp.Distinct):
                    # COUNT(DISTINCT column)的情况
                    distinct_expr = expr.this.expressions[0]
                    if isinstance(distinct_expr, exp.Column):
                        col_name = distinct_expr.name
                    elif hasattr(distinct_expr, 'name'):
                        col_name = distinct_expr.name
                    else:
                        raise ValueError(f"COUNT(DISTINCT)参数格式错误: {distinct_expr}")
                    return grouped[col_name].nunique()
                else:
                    # COUNT(column)的情况
                    if isinstance(expr.this, exp.Column):
                        col_name = expr.this.name
                    elif hasattr(expr.this, 'name'):
                        col_name = expr.this.name
                    else:
                        raise ValueError(f"COUNT函数参数格式错误: {expr.this}")
                    return grouped[col_name].count()
            elif isinstance(expr.this, exp.Star):
                # 其他函数不支持*
                raise ValueError(f"函数 {func_name} 不支持 * 参数")

            # 对于其他聚合函数，提取列名
            if isinstance(expr.this, exp.Column):
                col_name = expr.this.name
            elif hasattr(expr.this, 'name'):
                col_name = expr.this.name
            else:
                raise ValueError(f"聚合函数 {func_name} 参数格式错误: {expr.this}")

            # 获取原始列数据并转换为数值类型
            # 注意：需要从grouped的obj获取原始DataFrame
            try:
                original_df = grouped.obj
            except Exception:
                original_df = None
            
            # 应用对应的聚合函数
            # 直接使用 groupby agg，避免 apply+func 返回标量的问题
            if func_name == 'sum':
                return grouped[col_name].agg(lambda x: pd.to_numeric(x, errors='coerce').sum())
            elif func_name == 'avg':
                return grouped[col_name].agg(lambda x: pd.to_numeric(x, errors='coerce').mean())
            elif func_name == 'max':
                return grouped[col_name].agg(lambda x: pd.to_numeric(x, errors='coerce').max())
            elif func_name == 'min':
                return grouped[col_name].agg(lambda x: pd.to_numeric(x, errors='coerce').min())
            else:
                raise ValueError(f"不支持的聚合函数: {func_name}")

        elif isinstance(expr, exp.Alias):
            return self._apply_aggregation_function(expr.this, grouped)

        else:
            raise ValueError(f"不是聚合函数: {type(expr)}")

    def _apply_having_clause(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用HAVING条件"""
        having_clause = parsed_sql.args.get('having')
        if not having_clause:
            return df

        # 构建聚合表达式→SELECT别名的映射（HAVING COUNT(*) > 1 需要找到 cnt 列）
        self._having_agg_alias_map = {}
        for select_expr in parsed_sql.expressions:
            if isinstance(select_expr, exp.Alias) and isinstance(select_expr.this, exp.AggFunc):
                agg_sql = select_expr.this.sql()
                self._having_agg_alias_map[agg_sql] = select_expr.alias

        # HAVING子句处理类似于WHERE，但作用于聚合后的数据
        condition_str = self._sql_condition_to_pandas(having_clause.this, df)

        if condition_str:
            try:
                return df.query(condition_str)
            except Exception as e:
                # 备用方案：逐行过滤
                mask = []
                for _, row in df.iterrows():
                    if self._evaluate_condition_for_row(having_clause.this, row):
                        mask.append(True)
                    else:
                        mask.append(False)
                return df[mask]

        return df

    def _extract_select_aliases(self, parsed_sql: exp.Expression) -> Dict[str, Any]:
        """提取SELECT子句中的别名映射

        Returns:
            Dict: {alias_name: original_expression} 或 {column_name: column_name}
        """
        aliases = {}
        for i, select_expr in enumerate(parsed_sql.expressions):
            if isinstance(select_expr, exp.Alias):
                aliases[select_expr.alias] = select_expr.this
            elif isinstance(select_expr, exp.Column) and hasattr(select_expr, 'name'):
                aliases[select_expr.name] = select_expr
            elif self._is_aggregate_function(select_expr):
                # 无别名的聚合函数，生成列名
                aliases[self._generate_aggregate_alias(select_expr)] = select_expr
            elif isinstance(select_expr, exp.Case):
                # CASE WHEN无别名的处理
                pass
        return aliases

    def _resolve_order_column(self, col_name: str, df: pd.DataFrame, select_aliases: Optional[Dict] = None) -> Optional[str]:
        """解析ORDER BY列名：先查SELECT别名对应的基础列，再查原始列名

        Args:
            col_name: ORDER BY中引用的列名
            df: 当前DataFrame
            select_aliases: SELECT别名映射

        Returns:
            解析后的实际列名，找不到返回None
        """
        # 1. 如果列名直接在DataFrame中，直接返回
        if col_name in df.columns:
            return col_name

        # 2. 如果有SELECT别名映射，检查别名对应的基础列
        if select_aliases and col_name in select_aliases:
            expr = select_aliases[col_name]
            if isinstance(expr, exp.Column) and expr.name in df.columns:
                return expr.name
            # 别名对应的是计算表达式，无法在SELECT之前排序
            # 这种情况需要特殊处理：先计算表达式列，排序后删除
            if self._is_mathematical_expression(expr):
                # 临时计算该表达式
                temp_col = f"__order_temp_{col_name}"
                df[temp_col] = self._evaluate_math_expression(expr, df)
                # 重命名到目标列名（SELECT后会被处理）
                df.rename(columns={temp_col: col_name}, inplace=True)
                return col_name
            # CASE WHEN表达式：临时计算
            if isinstance(expr, exp.Case):
                temp_col = f"__order_temp_{col_name}"
                df[temp_col] = self._evaluate_case_expression(expr, df)
                df.rename(columns={temp_col: col_name}, inplace=True)
                return col_name
            # COALESCE表达式：临时计算
            if isinstance(expr, exp.Coalesce):
                temp_col = f"__order_temp_{col_name}"
                results = []
                for idx in range(len(df)):
                    row = df.iloc[idx]
                    results.append(self._evaluate_coalesce_for_row(expr, row))
                df[temp_col] = pd.Series(results, index=df.index)
                df.rename(columns={temp_col: col_name}, inplace=True)
                return col_name

        # 3. 列名不存在
        return None

    def _apply_order_by(self, parsed_sql: exp.Expression, df: pd.DataFrame, select_aliases: Optional[Dict] = None) -> pd.DataFrame:
        """应用ORDER BY排序

        Args:
            parsed_sql: 解析后的SQL表达式
            df: 数据DataFrame
            select_aliases: SELECT子句的别名映射（允许ORDER BY引用别名）
        """
        order_clause = parsed_sql.args.get('order')
        if not order_clause:
            return df

        sort_columns = []
        ascending = []

        for order_expr in order_clause.expressions:
            if isinstance(order_expr, exp.Ordered):
                col_expr = order_expr.this
                col_name = col_expr.name
                table_part = col_expr.table if hasattr(col_expr, 'table') and col_expr.table else None
                qualified = f"{table_part}.{col_name}" if table_part else None

                # 先查限定名，再查简单名，再查SELECT别名
                resolved_name = qualified if qualified and qualified in df.columns else None
                if resolved_name is None:
                    resolved_name = self._resolve_order_column(col_name, df, select_aliases)
                if resolved_name is None and qualified and qualified in df.columns:
                    resolved_name = qualified
                if resolved_name is None:
                    suggestion = self._suggest_column_name(col_name, list(df.columns))
                    raise ValueError(f"排序列 '{qualified or col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                sort_columns.append(resolved_name)
                is_desc = order_expr.args.get('desc', False)
                ascending.append(not is_desc if is_desc is not None else True)
            else:
                # 简单列引用，默认升序
                if isinstance(order_expr, exp.Column):
                    col_name = order_expr.name
                    table_part = order_expr.table if hasattr(order_expr, 'table') and order_expr.table else None
                    qualified = f"{table_part}.{col_name}" if table_part else None

                    resolved_name = qualified if qualified and qualified in df.columns else None
                    if resolved_name is None:
                        resolved_name = self._resolve_order_column(col_name, df, select_aliases)
                    if resolved_name is None and qualified and qualified in df.columns:
                        resolved_name = qualified
                    if resolved_name is None:
                        suggestion = self._suggest_column_name(col_name, list(df.columns))
                        raise ValueError(f"排序列 '{qualified or col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                    sort_columns.append(resolved_name)
                    ascending.append(True)

        if sort_columns:
            return df.sort_values(by=sort_columns, ascending=ascending)

        return df

    def _build_total_row(self, result_df: pd.DataFrame) -> Optional[List]:
        """构建GROUP BY聚合结果的TOTAL汇总行"""
        if result_df.empty or len(result_df) <= 1:
            return None
        total_row = [''] * len(result_df.columns)
        total_row[0] = 'TOTAL'
        has_numeric = False
        for i, col in enumerate(result_df.columns):
            series = pd.to_numeric(result_df[col], errors='coerce')
            if series.notna().sum() > len(result_df) * 0.5:
                total_row[i] = self._serialize_value(series.sum())
                has_numeric = True
        return total_row if has_numeric else None

    def _generate_markdown_table(self, data: List, max_rows: int = 50) -> str:
        """将查询结果数据转为Markdown表格"""
        if not data:
            return ''
        md_lines = ['| ' + ' | '.join(str(c) for c in data[0]) + ' |']
        md_lines.append('| ' + ' | '.join(['---'] * len(data[0])) + ' |')
        display_rows = min(len(data) - 1, max_rows)
        for row in data[1:1 + display_rows]:
            md_lines.append('| ' + ' | '.join(str(c) for c in row) + ' |')
        if len(data) - 1 > max_rows:
            md_lines.append(f'| ... 共{len(data) - 1}行，仅显示前{max_rows}行 |')
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
        parsed_sql: exp.Expression = None,
        df_before_where: pd.DataFrame = None,
        output_format: str = "table"
    ) -> Dict[str, Any]:
        """格式化查询结果

        Args:
            has_group_by: 如果为True且有数值聚合列，自动追加TOTAL行
            parsed_sql: 解析后的SQL表达式，用于空结果智能建议
            df_before_where: WHERE过滤前的DataFrame，用于空结果智能建议
            output_format: 输出格式 table/json/csv
        """

        # 计算原始数据统计
        total_original_rows = sum(len(df) for df in worksheets_data.values())

        # 准备返回数据
        data = []
        if include_headers:
            data.append(list(result_df.columns))
        if not result_df.empty:
            for _, row in result_df.iterrows():
                data.append([self._serialize_value(val) for val in row])

        # 大结果自动截断：保护AI上下文窗口（MAX_RESULT_ROWS=500）
        MAX_RESULT_ROWS = 500
        truncated = False
        data_row_count = len(result_df)
        if data_row_count > MAX_RESULT_ROWS:
            # 保留表头行 + 前MAX_RESULT_ROWS行数据
            keep_rows = MAX_RESULT_ROWS + (1 if include_headers else 0)
            data = data[:keep_rows]
            truncated = True

        # GROUP BY 聚合结果自动追加 TOTAL 行
        has_total_row = False
        if has_group_by and include_headers:
            total_row = self._build_total_row(result_df)
            if total_row:
                data.append(total_row)
                has_total_row = True

        # 双行表头：构建列描述映射
        column_descriptions = {}
        if hasattr(self, '_header_descriptions') and self._header_descriptions:
            for table_name, desc_map in self._header_descriptions.items():
                for col in (result_df.columns if not result_df.empty else []):
                    if col in desc_map:
                        column_descriptions[col] = desc_map[col]

        # 性能提示：无LIMIT且返回行数过多时建议加LIMIT
        perf_hint = ''
        if len(result_df) > 100:
            has_limit = parsed_sql is not None and parsed_sql.args.get('limit') is not None
            if not has_limit:
                perf_hint = f'（结果较多，建议加 LIMIT 缩小范围）'
        if truncated:
            perf_hint += f'（结果已截断为前{MAX_RESULT_ROWS}行，共{data_row_count}行，请加 LIMIT 精确查询）'

        result = {
            'success': True,
            'message': f'SQL查询成功执行，返回 {data_row_count} 行结果' + ('（含TOTAL汇总行）' if has_total_row else '') + perf_hint,
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

        # 空结果智能建议：分析WHERE/HAVING条件类型，给出针对性提示
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

        # 生成Markdown表格（方便AI和人类阅读）
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

    def _infer_data_types(self, df: pd.DataFrame) -> Dict[str, str]:
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
        执行UPDATE语句，基于WHERE条件批量修改Excel数据

        支持语法: UPDATE 表名 SET 列1=值1, 列2=值2 [WHERE 条件]
        SET表达式支持: 列=常量, 列=列, 列=算术表达式(如 伤害*1.1)
        WHERE条件复用查询引擎的所有条件语法

        Args:
            file_path: Excel文件路径
            sql: UPDATE SQL语句
            sheet_name: 工作表名称（可选）
            dry_run: 预览模式，只返回影响行数不实际修改

        Returns:
            Dict: 更新结果
        """
        start_time = time.time()

        # 验证文件
        if not os.path.exists(file_path):
            return self._update_error(f'文件不存在: {file_path}')

        if not SQLGLOT_AVAILABLE:
            return self._update_error('SQLGLOT未安装，无法使用UPDATE功能')

        # 加载数据（使用缓存）
        worksheets_data = self._load_data_with_cache(file_path, sheet_name)

        if not worksheets_data:
            return self._update_error('无法加载Excel数据')

        # 解析UPDATE语句
        # 中文列名替换（与SELECT查询保持一致）
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
            return self._update_error('只支持UPDATE语句。💡 写入操作只支持UPDATE，查询请用 excel_query')

        # 提取表名（sqlglot中table在this属性）
        table_node = parsed.this if isinstance(parsed.this, exp.Table) else None
        if not table_node:
            return self._update_error('UPDATE语句缺少表名')
        target_table = table_node.name

        # 匹配工作表（支持中英文表名）
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
                f"工作表 '{target_table}' 不存在。可用工作表: {available}。{suggestion}")

        df = worksheets_data[matched_sheet].copy()
        original_df = df.copy()

        # 中文列名替换
        cn_map = {}
        desc_map = self._header_descriptions.get(matched_sheet, {})
        for en_col, cn_desc in desc_map.items():
            if en_col in df.columns:
                cn_map[cn_desc] = en_col

        # 解析SET子句（sqlglot中在expressions属性）
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
                        f"列 '{col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")
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
                filtered_df = df
        else:
            filtered_df = df

        if filtered_df.empty:
            return {'success': True, 'message': '没有匹配WHERE条件的行，无需更新',
                    'affected_rows': 0, 'changes': [], 'execution_time_ms': 0}

        affected_indices = filtered_df.index.tolist()
        changes = []

        # 应用SET操作
        for col_name, value_expr in set_operations:
            for idx in affected_indices:
                old_val = df.at[idx, col_name]
                new_val = self._evaluate_update_expression(value_expr, df, idx)

                # 类型兼容性：数值类型可互通，其他类型尝试转为旧值类型
                if old_val != '' and new_val != '' and type(old_val) != type(new_val):
                    if isinstance(old_val, (int, float)) and isinstance(new_val, (int, float)):
                        pass  # 数值互通：不转换
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
                    'message': f'匹配 {len(affected_indices)} 行，但值无变化',
                    'affected_rows': len(affected_indices),
                    'changes': [], 'execution_time_ms': round(elapsed, 1)}

        if dry_run:
            elapsed = (time.time() - start_time) * 1000
            return {'success': True,
                    'message': f'[预览] 将修改 {len(changes)} 个单元格（{len(affected_indices)} 行）',
                    'affected_rows': len(affected_indices),
                    'changes': changes, 'dry_run': True,
                    'execution_time_ms': round(elapsed, 1)}

        # 写回Excel（事务保护：失败自动回滚）
        backup_path = None
        try:
            with self._file_lock(file_path):
                # 创建临时备份（事务保护）
                backup_path = tempfile.mktemp(suffix='.xlsx.bak')
                shutil.copy2(file_path, backup_path)

                # 检测双行表头偏移
                header_row_offset = 0
                desc_map = self._header_descriptions.get(matched_sheet, {})
                if desc_map:
                    header_row_offset = 1  # 双行表头，数据从第3行开始

                wb = openpyxl.load_workbook(file_path)
                ws = wb[matched_sheet]

                for change in changes:
                    excel_row = change['row'] + header_row_offset
                    col_idx = list(df.columns).index(change['column']) + 1
                    ws.cell(row=excel_row, column=col_idx, value=change['new_value'])

                wb.save(file_path)
                wb.close()

                # 写入成功，删除备份
                if backup_path and os.path.exists(backup_path):
                    os.remove(backup_path)

                # 清除缓存（文件已修改）
                self._df_cache.pop(file_path, None)

                elapsed = (time.time() - start_time) * 1000
                return {'success': True,
                        'message': f'成功更新 {len(changes)} 个单元格（{len(affected_indices)} 行）',
                        'affected_rows': len(affected_indices),
                        'changes': changes,
                        'execution_time_ms': round(elapsed, 1)}

        except Exception as e:
            # 事务回滚：从备份恢复
            if backup_path and os.path.exists(backup_path):
                try:
                    shutil.copy2(backup_path, file_path)
                    os.remove(backup_path)
                except Exception:
                    pass
            return {'success': False,
                    'message': f'写入Excel失败，已自动回滚: {e}',
                    'affected_rows': 0, 'changes': changes,
                    'execution_time_ms': round((time.time() - start_time) * 1000, 1)}

    @contextmanager
    def _file_lock(self, file_path: str) -> Generator[None, None, None]:
        """文件锁上下文管理器（Linux fcntl，其他平台优雅降级）"""
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
        """智能序列化值：数值保持数值类型，None/NaN转空字符串，numpy→Python原生"""
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
        """将值序列化为JSON安全类型（numpy→Python原生）— 委托给_serialize_value"""
        return self._serialize_value(val)

    def _evaluate_update_expression(
        self, expr: exp.Expression, df: pd.DataFrame, row_idx: int
    ) -> Any:
        """
        评估UPDATE SET表达式，支持常量、列引用和算术运算

        Args:
            expr: SQL表达式
            df: DataFrame
            row_idx: 行索引

        Returns:
            计算后的值
        """
        if isinstance(expr, exp.Literal):
            val = expr.this
            if expr.is_string:
                return str(val)
            if isinstance(val, str):
                try:
                    if '.' in val:
                        return float(val)
                    return int(val)
                except ValueError:
                    return val
            return val

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
                # 分发表：与 _MATH_BINARY_OPS 风格统一，新增运算符只需一行
                _OPS = {exp.Add: operator.add, exp.Sub: operator.sub,
                        exp.Mul: operator.mul, exp.Div: operator.truediv}
                result = _OPS[type(expr)](left_n, right_n if type(expr) != exp.Div or right_n != 0 else 0)
                # 如果原值都是整数且非除法，返回整数
                if isinstance(left, int) and isinstance(right, int) and type(expr) != exp.Div:
                    return int(result)
                return result
            except (ValueError, TypeError):
                return ''

        else:
            # 未知表达式类型，尝试递归
            if hasattr(expr, 'this'):
                return self._evaluate_update_expression(expr.this, df, row_idx)
            return ''


# 模块级单例引擎，DataFrame缓存跨调用共享
_shared_engine: Optional[AdvancedSQLQueryEngine] = None


def _get_engine() -> AdvancedSQLQueryEngine:
    """获取共享SQL引擎实例（缓存跨调用复用）"""
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
    便捷函数：执行高级SQL查询

    Args:
        file_path: Excel文件路径
        sql: SQL查询语句
        sheet_name: 工作表名称（可选）
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
            'message': f'SQLGlot未安装，无法使用高级SQL功能: {str(e)}',
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
    便捷函数：执行UPDATE SQL语句

    Args:
        file_path: Excel文件路径
        sql: UPDATE SQL语句
        sheet_name: 工作表名称（可选）
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
            'message': f'SQLGLOT未安装，无法使用UPDATE功能: {str(e)}',
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