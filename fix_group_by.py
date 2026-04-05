import re

# Read the original file
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'r') as f:
    content = f.read()

# Fix the _apply_group_by_aggregation method
# The main issue is in how aggregations are applied and how the total row is calculated

old_method = '''    def _apply_group_by_aggregation(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用GROUP BY和聚合函数
        性能优化：
        - 大数据集使用向量化操作代替逐行计算
        - 减少不必要的DataFrame复制
        - 优化groupby操作使用observed=True避免稀疏分组
        """
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
            alias_name, original_expr = self._extract_select_alias(select_expr, i)
            select_exprs[alias_name] = original_expr

        # 检查聚合函数
        for alias_name, expr in select_exprs.items():
            if self._is_aggregate_function(expr):
                aggregations[alias_name] = expr
            elif hasattr(expr, 'name') and expr.name not in group_by_columns:
                # 如果是非聚合列且不在GROUP BY中，需要添加到GROUP BY
                if isinstance(expr, (exp.Column, exp.Identifier)):
                    group_by_columns.append(expr.name)

        # 保存GROUP BY列到实例变量，供_build_total_row使用
        self._group_by_columns = group_by_columns

        if not aggregations:
            # 没有聚合函数，只应用GROUP BY去重
            if group_by_columns:
                # 性能优化：使用drop_duplicates的subset参数避免全列比较
                return df[group_by_columns].drop_duplicates(subset=group_by_columns).reset_index(drop=True)
            else:
                return df

        # 预计算CASE WHEN/COALESCE/标量子查询表达式，添加到df副本，使grouped可访问
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
        # 性能优化：使用observed=True减少分组计算开销
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
            ordered_columns.append(alias_name)

            # 处理聚合函数
            is_agg = self._is_aggregate_function(select_expr if not isinstance(select_expr, exp.Alias) else select_expr.this)
            
            if is_agg:
                agg_expr = original_expr if isinstance(select_expr, exp.Alias) else select_expr
                agg_result = self._apply_aggregation_function(agg_expr, grouped, df)
                if isinstance(agg_result, (int, float, np.integer, np.floating)):
                    result_data[alias_name] = pd.Series([agg_result])
                else:
                    result_data[alias_name] = agg_result
            # 处理CASE WHEN表达式（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Case):
                result_data[alias_name] = grouped[alias_name].first()
            # 处理COALESCE表达式（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Coalesce):
                result_data[alias_name] = grouped[alias_name].first()
            # 处理标量子查询（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Subquery):
                result_data[alias_name] = grouped[alias_name].first()
            # 处理普通列（GROUP BY列）
            elif hasattr(original_expr, 'name'):
                col_name = original_expr.name
                if col_name in group_by_columns:
                    result_data[alias_name] = grouped[col_name].first()
            # 处理SELECT * 的情况
            elif alias_name == '*':
                for col in group_by_columns:
                    if col not in result_data:
                        result_data[col] = grouped[col].first()

        # 组合结果，保持列顺序
        result_df = pd.DataFrame(result_data, columns=ordered_columns).reset_index(drop=True)

        return result_df'''

new_method = '''    def _apply_group_by_aggregation(self, parsed_sql: exp.Expression, df: pd.DataFrame) -> pd.DataFrame:
        """应用GROUP BY和聚合函数
        性能优化：
        - 大数据集使用向量化操作代替逐行计算
        - 减少不必要的DataFrame复制
        - 优化groupby操作使用observed=True避免稀疏分组
        """
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
            alias_name, original_expr = self._extract_select_alias(select_expr, i)
            select_exprs[alias_name] = original_expr

        # 检查聚合函数
        for alias_name, expr in select_exprs.items():
            if self._is_aggregate_function(expr):
                aggregations[alias_name] = expr
            elif hasattr(expr, 'name') and expr.name not in group_by_columns:
                # 如果是非聚合列且不在GROUP BY中，需要添加到GROUP BY
                if isinstance(expr, (exp.Column, exp.Identifier)):
                    group_by_columns.append(expr.name)

        # 保存GROUP BY列到实例变量，供_build_total_row使用
        self._group_by_columns = group_by_columns

        if not aggregations:
            # 没有聚合函数，只应用GROUP BY去重
            if group_by_columns:
                # 性能优化：使用drop_duplicates的subset参数避免全列比较
                return df[group_by_columns].drop_duplicates(subset=group_by_columns).reset_index(drop=True)
            else:
                return df

        # 预计算CASE WHEN/COALESCE/标量子查询表达式，添加到df副本，使grouped可访问
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
        # 性能优化：使用observed=True减少分组计算开销
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
            ordered_columns.append(alias_name)

            # 处理聚合函数
            is_agg = self._is_aggregate_function(select_expr if not isinstance(select_expr, exp.Alias) else select_expr.this)
            
            if is_agg:
                agg_expr = original_expr if isinstance(select_expr, exp.Alias) else select_expr
                agg_result = self._apply_aggregation_function(agg_expr, grouped, df)
                # 修复：确保聚合结果是正确格式
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
            # 处理CASE WHEN表达式（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Case):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理COALESCE表达式（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Coalesce):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理标量子查询（已预计算到df，直接从grouped取first）
            elif isinstance(original_expr, exp.Subquery):
                result_data[alias_name] = grouped[alias_name].first().reset_index(drop=True)
            # 处理普通列（GROUP BY列）
            elif hasattr(original_expr, 'name'):
                col_name = original_expr.name
                if col_name in group_by_columns:
                    result_data[alias_name] = grouped[col_name].first().reset_index(drop=True)
            # 处理SELECT * 的情况
            elif alias_name == '*':
                for col in group_by_columns:
                    if col not in result_data:
                        result_data[col] = grouped[col].first().reset_index(drop=True)

        # 组合结果，保持列顺序
        try:
            result_df = pd.DataFrame(result_data, columns=ordered_columns).reset_index(drop=True)
        except Exception as e:
            # 如果创建DataFrame失败，尝试逐列构建
            print(f"警告：构建结果DataFrame时出错: {e}")
            result_data = {k: v.reset_index(drop=True) for k, v in result_data.items()}
            result_df = pd.DataFrame(result_data, columns=ordered_columns).reset_index(drop=True)

        return result_df'''

# Replace the method
content = content.replace(old_method, new_method)

# Write back to file
with open('src/excel_mcp_server_fastmcp/api/advanced_sql_query.py', 'w') as f:
    f.write(content)

print("GROUP BY聚合逻辑修复完成")
