"""
高级SQL查询引擎 - 基于SQLGlot实现的SQL查询支持

支持功能：
- 基础查询: SELECT, DISTINCT, 别名
- 条件筛选: WHERE, LIKE, IN, BETWEEN, AND/OR
- 聚合统计: COUNT, SUM, AVG, MAX, MIN, GROUP BY, HAVING
- 排序限制: ORDER BY, LIMIT, OFFSET
- 算术运算: 加减乘除

不支持功能：
- 子查询、CTE (WITH)、JOIN、UNION
- 窗口函数 (ROW_NUMBER等)
- CASE WHEN、EXISTS
- INSERT/UPDATE/DELETE
"""

import os
import re
from typing import Dict, List, Any, Optional, Union, Tuple
import pandas as pd
import numpy as np

# SQLGlot导入 - 核心SQL解析引擎
try:
    import sqlglot
    from sqlglot import expressions as exp
    from sqlglot.errors import ParseError, UnsupportedError
    SQLGLOT_AVAILABLE = True
except ImportError:
    SQLGLOT_AVAILABLE = False
    print("警告: SQLGlot未安装，将使用基础pandas查询功能")
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
        include_headers: bool = True
    ) -> Dict[str, Any]:
        """
        执行SQL查询，支持完整的SQL语法

        Args:
            file_path: Excel文件路径
            sql: SQL查询语句
            sheet_name: 工作表名称（可选，默认使用第一个）
            limit: 限制返回行数
            include_headers: 是否包含表头

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
            import time as _time
            mtime = os.path.getmtime(file_path)
            cache_key = file_path
            if cache_key in self._df_cache:
                cached_mtime, cached_data, cached_desc = self._df_cache[cache_key]
                if cached_mtime == mtime:
                    worksheets_data = cached_data
                    self._header_descriptions = cached_desc
                else:
                    # 文件已修改，重新加载
                    worksheets_data = self._load_excel_data(file_path, sheet_name)
                    self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
            else:
                worksheets_data = self._load_excel_data(file_path, sheet_name)
                self._df_cache[cache_key] = (mtime, worksheets_data, self._header_descriptions)
                # LRU淘汰：超过最大缓存数时删除最早缓存的文件
                while len(self._df_cache) > self._max_cache_size:
                    self._df_cache.pop(next(iter(self._df_cache)))

            if not worksheets_data:
                return {
                    'success': False,
                    'message': '无法加载Excel数据或文件为空',
                    'data': [],
                    'query_info': {'error_type': 'data_load_failed'}
                }

            # 中文列名替换：将SQL中的中文列名替换为英文列名（在解析前）
            sql = self._replace_cn_columns_in_sql(sql, worksheets_data)

            # 解析和执行SQL
            import time as _time
            _query_start = _time.time()
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

                # 执行查询
                result_data = self._execute_query(parsed_sql, worksheets_data, limit)
                _query_elapsed = (_time.time() - _query_start) * 1000

                # 格式化结果（传入parsed_sql判断是否需要总计行）
                has_group_by = parsed_sql.args.get('group') is not None
                result = self._format_query_result(
                    result_data,
                    file_path,
                    sql,
                    worksheets_data,
                    include_headers,
                    has_group_by=has_group_by
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
            if sheet_name:
                sheets_to_load = [sheet_name]
            else:
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                sheets_to_load = excel_file.sheet_names

            for sheet in sheets_to_load:
                # 先用 openpyxl 读取前两行，检测双行表头
                try:
                    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                    ws = wb[sheet]
                    rows_iter = ws.iter_rows(max_row=2, values_only=False)
                    first_row_cells = next(rows_iter, None)
                    second_row_cells = next(rows_iter, None)
                    wb.close()
                except Exception:
                    first_row_cells = None
                    second_row_cells = None

                is_dual_header = False
                if first_row_cells and second_row_cells:
                    second_row_values = [str(c.value).strip() if c.value else '' for c in second_row_cells]
                    first_row_values = [str(c.value).strip() if c.value else '' for c in first_row_cells]
                    
                    non_empty_second = [v for v in second_row_values if v]
                    non_empty_first = [v for v in first_row_values if v]
                    
                    # 严格双行表头检测：
                    # 1. 至少3个非空值
                    # 2. 第二行全部匹配字段名模式（^[a-zA-Z_]\w*$）
                    # 3. 第一行不全匹配字段名模式（排除两行都是字段名的普通表）
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

                if is_dual_header:
                    # 双行表头：用第二行做列名，跳过第一行
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='openpyxl',
                        header=1,  # 第二行做列名
                        keep_default_na=False
                    )
                    # 记录描述映射
                    field_names = [str(c.value).strip() if c.value else '' for c in second_row_cells]
                    descriptions = [str(c.value).strip() if c.value else '' for c in first_row_cells]
                    desc_map = {}
                    for fname, desc in zip(field_names, descriptions):
                        if fname:
                            desc_map[fname] = desc
                    self._header_descriptions[sheet] = desc_map
                else:
                    # 单行表头
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='openpyxl',
                        keep_default_na=False
                    )

                df = self._clean_dataframe(df)
                worksheets_data[sheet] = df

        except Exception as e:
            print(f"加载Excel数据失败: {e}")
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
                except:
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
            # 检查是否为SELECT语句
            if not isinstance(parsed_sql, exp.Select):
                return {
                    'valid': False,
                    'error': '只支持SELECT查询语句，不支持INSERT、UPDATE、DELETE等操作'
                }

            # 检查不支持的子查询
            for subquery in parsed_sql.find_all(exp.Subquery):
                return {
                    'valid': False,
                    'error': '不支持子查询'
                }

            # 检查不支持的CTE (WITH子句)
            if parsed_sql.find(exp.With):
                return {
                    'valid': False,
                    'error': '不支持WITH子句（公用表表达式）'
                }

            # 检查窗口函数
            for window in parsed_sql.find_all(exp.Window):
                return {
                    'valid': False,
                    'error': '不支持窗口函数（OVER子句）'
                }

            # 检查不支持的JOIN
            if parsed_sql.args.get('joins') or parsed_sql.find(exp.Join):
                available_tables = list(self._pending_tables) if hasattr(self, '_pending_tables') else []
                return {
                    'valid': False,
                    'error': '不支持JOIN查询。💡 游戏配置表关联查询替代方案：\n'
                             '1. 先用 excel_get_range 读取两个表的数据\n'
                             '2. 在AI层面做关联匹配\n'
                             '3. 或将需要关联的字段合并到同一个工作表中'
                }

            # 检查不支持的CASE WHEN
            if parsed_sql.find(exp.Case):
                return {
                    'valid': False,
                    'error': '不支持CASE WHEN表达式。💡 替代方案：\n'
                             '1. 用多个 WHERE 条件分别查询再合并\n'
                             '2. 用 excel_update_range 批量修改数据\n'
                             '3. 使用 excel_get_range 读取后在外部处理条件逻辑'
                }

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

    def _suggest_column_name(self, col_name: str, available_cols: List[str], max_suggestions: int = 3) -> str:
        """
        当列名不存在时，用编辑距离找出最相似的列名作为建议。

        Args:
            col_name: 用户输入的列名
            available_cols: 可用的列名列表
            max_suggestions: 最多返回几个建议

        Returns:
            str: 格式化的建议字符串，如 "你是否想用: skill_name, skill_id, skill_type?"
        """
        import difflib
        if not available_cols:
            return ""

        matches = difflib.get_close_matches(col_name, available_cols, n=max_suggestions, cutoff=0.3)
        if not matches:
            return ""

        return f" 你是否想用: {', '.join(matches)}?"

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
        # 获取FROM子句中的表名
        from_table = self._get_from_table(parsed_sql)

        if from_table not in worksheets_data:
            raise ValueError(f"表 '{from_table}' 不存在。可用表: {list(worksheets_data.keys())}")

        base_df = worksheets_data[from_table].copy()

        # 应用WHERE条件
        base_df = self._apply_where_clause(parsed_sql, base_df)

        # 检查是否有聚合函数
        has_aggregate = self._check_has_aggregate_function(parsed_sql)

        # 应用GROUP BY和聚合
        if parsed_sql.args.get('group') or has_aggregate:
            # 有GROUP BY或有聚合函数时，应用分组聚合
            base_df = self._apply_group_by_aggregation(parsed_sql, base_df)

            # 应用HAVING条件
            if parsed_sql.args.get('having'):
                base_df = self._apply_having_clause(parsed_sql, base_df)

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
                    # 普通列引用
                    column_name = original_expr.name
                    if column_name in df.columns:
                        result_data[alias_name] = df[column_name]
                    else:
                        suggestion = self._suggest_column_name(column_name, list(df.columns))
                        raise ValueError(f"列 '{column_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                elif self._is_mathematical_expression(original_expr):
                    # 数学表达式
                    result_data[alias_name] = self._evaluate_math_expression(original_expr, df)

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

    def _evaluate_math_expression(self, expr, df: pd.DataFrame):
        """计算数学表达式"""
        if isinstance(expr, exp.Add):
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return left + right
        elif isinstance(expr, exp.Sub):
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return left - right
        elif isinstance(expr, exp.Mul):
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return left * right
        elif isinstance(expr, exp.Div):
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return left / right
        elif isinstance(expr, exp.Mod):
            left = self._evaluate_math_expression(expr.left, df)
            right = self._evaluate_math_expression(expr.right, df)
            return left % right
        elif isinstance(expr, exp.Column):
            return df[expr.name]
        elif isinstance(expr, exp.Literal):
            return self._expression_to_value(expr, df)
        else:
            raise ValueError(f"不支持的数学表达式部分: {expr}")

    def _get_from_table(self, parsed_sql: exp.Expression) -> str:
        """获取FROM子句中的表名"""
        from_clause = parsed_sql.args.get('from')
        if not from_clause:
            # 尝试使用 from_ 键（sqlglot的另一种存储方式）
            from_clause = parsed_sql.args.get('from_')
        if from_clause:
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'name'):
                return from_clause.this.name
            # 兼容 Table 对象
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'this'):
                return from_clause.this.this

        # 如果没有明确的FROM子句，返回第一个表名
        raise ValueError("无法确定FROM子句中的表名")

    def _apply_where_clause(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用WHERE条件"""
        where_clause = parsed_sql.args.get('where')
        if not where_clause:
            return df

        # 将SQLGlot表达式转换为pandas查询条件
        condition_str = self._sql_condition_to_pandas(where_clause.this, df)

        if condition_str:
            try:
                return df.query(condition_str)
            except Exception as e:
                # 如果查询失败，尝试逐行过滤
                return self._apply_row_filter(where_clause.this, df)

        return df

    def _sql_condition_to_pandas(self, condition: exp.Expression, df: pd.DataFrame) -> str:
        """将SQL条件转换为pandas查询字符串"""
        if isinstance(condition, exp.EQ):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} == {right}"

        elif isinstance(condition, exp.NEQ):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} != {right}"

        elif isinstance(condition, exp.GT):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} > {right}"

        elif isinstance(condition, exp.GTE):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} >= {right}"

        elif isinstance(condition, exp.LT):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} < {right}"

        elif isinstance(condition, exp.LTE):
            left = self._expression_to_column_reference(condition.left, df)
            right = self._expression_to_value(condition.right, df)
            return f"{left} <= {right}"

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
            left = self._expression_to_column_reference(condition.this, df)
            values = []
            for value in condition.expressions:
                values.append(self._expression_to_value(value, df))
            values_str = ', '.join(str(v) for v in values)
            return f"{left}.isin([{values_str}])"

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
        """将表达式转换为列引用"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name not in df.columns:
                suggestion = self._suggest_column_name(col_name, list(df.columns))
                raise ValueError(f"列 '{col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")
            return f"`{col_name}`"

        elif isinstance(expr, exp.Literal):
            return str(expr.this)

        elif isinstance(expr, exp.AggFunc):
            # 对于HAVING子句中的聚合函数，需要查找对应的列
            # 由于聚合后的DataFrame只有有限列，直接返回第一个非GROUP BY列
            # 或者如果只有一个数值列，就使用它
            func_name = type(expr).__name__.lower()
            
            # 首先尝试精确匹配：查找列名等于函数名的列
            if func_name in df.columns:
                return f"`{func_name}`"
            
            # 尝试模糊匹配：查找列名包含函数名的列
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
                except:
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
            if isinstance(condition, exp.EQ):
                left_val = self._get_row_value(condition.left, row)
                right_val = self._get_row_value(condition.right, row)
                return left_val == right_val

            elif isinstance(condition, exp.NEQ):
                left_val = self._get_row_value(condition.left, row)
                right_val = self._get_row_value(condition.right, row)
                return left_val != right_val

            elif isinstance(condition, exp.And):
                return (self._evaluate_condition_for_row(condition.left, row) and
                       self._evaluate_condition_for_row(condition.right, row))

            elif isinstance(condition, exp.Or):
                return (self._evaluate_condition_for_row(condition.left, row) or
                       self._evaluate_condition_for_row(condition.right, row))

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
            except:
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
                col_name = order_expr.this.name
                # 先查SELECT别名，再查原始列
                resolved_name = self._resolve_order_column(col_name, df, select_aliases)
                if resolved_name is None:
                    suggestion = self._suggest_column_name(col_name, list(df.columns))
                    raise ValueError(f"排序列 '{col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                sort_columns.append(resolved_name)
                is_desc = order_expr.args.get('desc', False)
                ascending.append(not is_desc if is_desc is not None else True)
            else:
                # 简单列引用，默认升序
                if isinstance(order_expr, exp.Column):
                    col_name = order_expr.name
                    resolved_name = self._resolve_order_column(col_name, df, select_aliases)
                    if resolved_name is None:
                        suggestion = self._suggest_column_name(col_name, list(df.columns))
                        raise ValueError(f"排序列 '{col_name}' 不存在。可用列: {list(df.columns)}。{suggestion}")

                    sort_columns.append(resolved_name)
                    ascending.append(True)

        if sort_columns:
            return df.sort_values(by=sort_columns, ascending=ascending)

        return df

    def _format_query_result(
        self,
        result_df: pd.DataFrame,
        file_path: str,
        sql: str,
        worksheets_data: Dict[str, pd.DataFrame],
        include_headers: bool,
        has_group_by: bool = False
    ) -> Dict[str, Any]:
        """格式化查询结果

        Args:
            has_group_by: 如果为True且有数值聚合列，自动追加TOTAL行
        """

        # 计算原始数据统计
        total_original_rows = sum(len(df) for df in worksheets_data.values())

        # 准备返回数据
        def _serialize_value(val):
            """智能序列化值：数值保持数值类型，None转空字符串"""
            if val is None:
                return ''
            if isinstance(val, float) and val == int(val):
                return int(val)  # 170.0 → 170
            if isinstance(val, (np.integer,)):
                return int(val)
            if isinstance(val, (np.floating,)):
                f = float(val)
                if f == int(f):
                    return int(f)
                # 非整数浮点数保留2位小数，避免166.66666666666666
                return round(f, 2)
            return val

        data = []
        if include_headers:
            # 包含表头（无论是否有数据）
            headers = list(result_df.columns)
            data.append(headers)

            # 添加数据行（如果有的话）
            if not result_df.empty:
                for _, row in result_df.iterrows():
                    data.append([_serialize_value(val) for val in row])
        else:
            # 不包含表头，只返回数据
            if not result_df.empty:
                for _, row in result_df.iterrows():
                    data.append([_serialize_value(val) for val in row])

        # GROUP BY 聚合结果自动追加 TOTAL 行
        has_total_row = False
        if has_group_by and not result_df.empty and len(result_df) > 1 and include_headers:
            numeric_cols = []
            for i, col in enumerate(result_df.columns):
                # 检测数值列（聚合结果通常是数值）
                series = pd.to_numeric(result_df[col], errors='coerce')
                if series.notna().sum() > len(result_df) * 0.5:
                    numeric_cols.append(i)
            if numeric_cols:
                total_row = [''] * len(result_df.columns)
                total_row[0] = 'TOTAL'
                for i in numeric_cols:
                    col_sum = 0.0
                    for row_idx in range(1, len(data)):  # 跳过表头行
                        val = data[row_idx][i]
                        if isinstance(val, (int, float)):
                            col_sum += val
                    total_row[i] = _serialize_value(col_sum)
                data.append(total_row)
                has_total_row = True

        # 双行表头：构建列描述映射
        column_descriptions = {}
        if hasattr(self, '_header_descriptions') and self._header_descriptions:
            for table_name, desc_map in self._header_descriptions.items():
                for col in (result_df.columns if not result_df.empty else []):
                    if col in desc_map:
                        column_descriptions[col] = desc_map[col]

        result = {
            'success': True,
            'message': f'SQL查询成功执行，返回 {len(result_df)} 行结果' + ('（含TOTAL汇总行）' if has_total_row else ''),
            'data': data,
            'query_info': {
                'original_rows': total_original_rows,
                'filtered_rows': len(result_df),
                'query_applied': True,
                'sql_query': sql,
                'columns_returned': len(result_df.columns) if not result_df.empty else 0,
                'available_tables': list(worksheets_data.keys()),
                'returned_columns': list(result_df.columns) if not result_df.empty else [],
                'data_types': self._infer_data_types(result_df) if not result_df.empty else {}
            }
        }

        # 空结果友好提示
        if result_df.empty:
            result['query_info']['suggestion'] = '查询返回0行数据。可能原因：WHERE条件过严、列名拼写错误（可用DESCRIBE查看列名）、或数据尚未录入。'

        # 生成Markdown表格（方便AI和人类阅读）
        if data and len(data) > 0:
            md_lines = []
            # 表头
            md_lines.append('| ' + ' | '.join(str(c) for c in data[0]) + ' |')
            md_lines.append('| ' + ' | '.join(['---'] * len(data[0])) + ' |')
            # 数据行（最多50行，避免超大输出）
            max_md_rows = min(len(data) - 1, 50)
            for row in data[1:1 + max_md_rows]:
                md_lines.append('| ' + ' | '.join(str(c) for c in row) + ' |')
            if len(data) - 1 > 50:
                md_lines.append(f'| ... 共{len(data) - 1}行，仅显示前50行 |')
            result['query_info']['markdown_table'] = '\n'.join(md_lines)

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
                    # 对可能是日期的列进行转换
                    pd.to_datetime(series, errors='coerce', format='mixed')
                    if not pd.to_datetime(series, errors='coerce').isna().all():
                        data_types[col] = 'datetime'
                        continue
            except Exception:
                pass

            # 默认为字符串类型
            data_types[col] = 'string'

        return data_types


def execute_advanced_sql_query(
    file_path: str,
    sql: str,
    sheet_name: Optional[str] = None,
    limit: Optional[int] = None,
    include_headers: bool = True
) -> Dict[str, Any]:
    """
    便捷函数：执行高级SQL查询

    Args:
        file_path: Excel文件路径
        sql: SQL查询语句
        sheet_name: 工作表名称（可选）
        limit: 结果限制
        include_headers: 是否包含表头

    Returns:
        Dict: 查询结果
    """
    try:
        engine = AdvancedSQLQueryEngine()
        return engine.execute_sql_query(
            file_path=file_path,
            sql=sql,
            sheet_name=sheet_name,
            limit=limit,
            include_headers=include_headers
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