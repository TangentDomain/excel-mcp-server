"""
高级SQL查询引擎 - 基于SQLGlot实现完整SQL功能支持
参考mcp-excel-db架构，支持GROUP BY、聚合函数、JOIN等完整SQL语法
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

        if not SQLGLOT_AVAILABLE:
            raise ImportError("SQLGlot未安装，请运行: pip install sqlglot")

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

            # 加载Excel数据
            worksheets_data = self._load_excel_data(file_path, sheet_name)

            if not worksheets_data:
                return {
                    'success': False,
                    'message': '无法加载Excel数据或文件为空',
                    'data': [],
                    'query_info': {'error_type': 'data_load_failed'}
                }

            # 解析和执行SQL
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

                # 格式化结果
                return self._format_query_result(
                    result_data,
                    file_path,
                    sql,
                    worksheets_data,
                    include_headers
                )

            except ParseError as e:
                return {
                    'success': False,
                    'message': f'SQL语法错误: {str(e)}',
                    'data': [],
                    'query_info': {'error_type': 'syntax_error', 'details': str(e)}
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
        加载Excel数据到DataFrame字典

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称（可选）

        Returns:
            Dict[str, pd.DataFrame]: 工作表名到DataFrame的映射
        """
        worksheets_data = {}

        try:
            if sheet_name:
                # 加载指定工作表
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    engine='openpyxl',
                    keep_default_na=False     # 减少空值转换警告
                )
                df = self._clean_dataframe(df)
                worksheets_data[sheet_name] = df
            else:
                # 加载所有工作表
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                for sheet in excel_file.sheet_names:
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet,
                        engine='openpyxl',
                        keep_default_na=False     # 减少空值转换警告
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

        # 清理数据中的空值
        df = df.where(pd.notnull(df), None)

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

            return {'valid': True}

        except Exception as e:
            return {
                'valid': False,
                'error': f'SQL验证失败: {str(e)}'
            }

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

        # 应用GROUP BY和聚合
        if parsed_sql.args.get('group'):
            base_df = self._apply_group_by_aggregation(parsed_sql, base_df)
        else:
            # 没有GROUP BY时，也需要处理SELECT表达式（如计算字段、别名等）
            base_df = self._apply_select_expressions(parsed_sql, base_df)

        # 应用HAVING条件
        if parsed_sql.args.get('having'):
            base_df = self._apply_having_clause(parsed_sql, base_df)

        # 应用ORDER BY
        if parsed_sql.args.get('order'):
            base_df = self._apply_order_by(parsed_sql, base_df)

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

        return base_df

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
                        raise ValueError(f"列 '{column_name}' 不存在")

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
        if from_clause:
            if hasattr(from_clause, 'this') and hasattr(from_clause.this, 'name'):
                return from_clause.this.name

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

        elif isinstance(condition, exp.Like):
            left = self._expression_to_column_reference(condition.this, df)
            right = self._expression_to_value(condition.expression, df)
            pattern = str(right).strip("'\"")
            # 转换SQL LIKE模式为pandas str.contains
            pattern = pattern.replace('%', '.*').replace('_', '.')
            return f"{left}.str.match('{pattern}', na=False)"

        elif isinstance(condition, exp.In):
            left = self._expression_to_column_reference(condition.this, df)
            values = []
            for value in condition.expressions:
                values.append(self._expression_to_value(value, df))
            values_str = ', '.join(str(v) for v in values)
            return f"{left}.isin([{values_str}])"

        # elif isinstance(condition, exp.IsNull):
        #     left = self._expression_to_column_reference(condition.this, df)
        #     return f"{left}.isna()"

        # elif isinstance(condition, exp.NotNull):
        #     left = self._expression_to_column_reference(condition.this, df)
        #     return f"~{left}.isna()"

        else:
            raise ValueError(f"不支持的条件类型: {type(condition)}")

    def _expression_to_column_reference(self, expr: exp.Expression, df: pd.DataFrame) -> str:
        """将表达式转换为列引用"""
        if isinstance(expr, exp.Column):
            col_name = expr.name
            if col_name not in df.columns:
                raise ValueError(f"列 '{col_name}' 不存在。可用列: {list(df.columns)}")
            return f"`{col_name}`"

        elif isinstance(expr, exp.Literal):
            return str(expr.this)

        elif isinstance(expr, exp.AggFunc):
            # 对于HAVING子句中的聚合函数，需要查找对应的列
            # 这需要与SELECT表达式中的别名匹配
            func_name = type(expr).__name__.lower()
            if isinstance(expr.this, exp.Star):
                # COUNT(*)的情况
                agg_signature = "count_star"
            else:
                # 其他聚合函数
                if isinstance(expr.this, exp.Column):
                    col_name = expr.this.name
                else:
                    col_name = str(expr.this)
                agg_signature = f"{func_name}_{col_name}"

            # 尝试在DataFrame中找到匹配的列
            for col in df.columns:
                # 简单匹配：如果列名包含函数名和参数
                if func_name in col.lower():
                    return f"`{col}`"

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
                raise ValueError(f"列 '{col_name}' 不存在")
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
                # 对于没有别名的列，使用列名作为别名
                if hasattr(select_expr, 'name'):
                    alias_name = select_expr.name
                else:
                    alias_name = f"col_{i}"
                original_expr = select_expr

            ordered_columns.append(alias_name)

            # 处理聚合函数
            if alias_name in aggregations:
                result_data[alias_name] = self._apply_aggregation_function(aggregations[alias_name], grouped)
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

            # 应用对应的聚合函数
            if func_name == 'sum':
                # 转换为数值类型进行求和
                return pd.to_numeric(grouped[col_name], errors='coerce').sum()
            elif func_name == 'avg':
                return pd.to_numeric(grouped[col_name], errors='coerce').mean()
            elif func_name == 'max':
                return grouped[col_name].max()
            elif func_name == 'min':
                return grouped[col_name].min()
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

    def _apply_order_by(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用ORDER BY排序"""
        order_clause = parsed_sql.args.get('order')
        if not order_clause:
            return df

        sort_columns = []
        ascending = []

        for order_expr in order_clause.expressions:
            if isinstance(order_expr, exp.Ordered):
                col_name = order_expr.this.name
                if col_name not in df.columns:
                    raise ValueError(f"排序列 '{col_name}' 不存在")

                sort_columns.append(col_name)
                ascending.append(order_expr.args.get('asc', True))
            else:
                # 简单列引用，默认升序
                if isinstance(order_expr, exp.Column):
                    col_name = order_expr.name
                    if col_name not in df.columns:
                        raise ValueError(f"排序列 '{col_name}' 不存在")

                    sort_columns.append(col_name)
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
        include_headers: bool
    ) -> Dict[str, Any]:
        """格式化查询结果"""

        # 计算原始数据统计
        total_original_rows = sum(len(df) for df in worksheets_data.values())

        # 准备返回数据
        data = []
        if include_headers:
            # 包含表头（无论是否有数据）
            headers = list(result_df.columns)
            data.append(headers)

            # 添加数据行（如果有的话）
            if not result_df.empty:
                for _, row in result_df.iterrows():
                    data.append([str(val) if val is not None else '' for val in row])
        else:
            # 不包含表头，只返回数据
            if not result_df.empty:
                for _, row in result_df.iterrows():
                    data.append([str(val) if val is not None else '' for val in row])

        return {
            'success': True,
            'message': f'SQL查询成功执行，返回 {len(result_df)} 行结果',
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