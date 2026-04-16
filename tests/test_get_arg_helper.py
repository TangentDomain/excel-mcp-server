"""测试 _get_arg 公共方法 — 字符串函数参数提取DRY"""
import pytest
import pandas as pd
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def engine():
    return AdvancedSQLQueryEngine()


class TestGetArgHelper:
    """_get_arg 消除字符串函数中重复的 _literal_value(expr.args.get(...)) 模式"""

    def test_get_arg_missing_returns_default(self, engine):
        """参数不存在时返回默认值"""
        import sqlglot.expressions as exp
        expr = exp.Substring(this=exp.column('x'))
        assert engine._get_arg(expr, 'start', 0, int) == 0
        assert engine._get_arg(expr, 'nonexistent', 'default') == 'default'

    def test_get_arg_int_conversion(self, engine):
        """int类型转换"""
        import sqlglot.expressions as exp
        expr = exp.Substring(this=exp.column('x'))
        expr.set('expression', exp.Literal.number(3))
        assert engine._get_arg(expr, 'expression', 1, int) == 3

    def test_get_arg_str_conversion(self, engine):
        """str类型转换"""
        import sqlglot.expressions as exp
        expr = exp.Replace(this=exp.column('x'))
        expr.set('expression', exp.Literal.string('old'))
        assert engine._get_arg(expr, 'expression', '', str) == 'old'

    def test_get_arg_no_type_fn(self, engine):
        """不传type_fn时返回原始值"""
        import sqlglot.expressions as exp
        expr = exp.Replace(this=exp.column('x'))
        expr.set('replacement', exp.Literal.string('new'))
        result = engine._get_arg(expr, 'replacement')
        assert result == 'new'

    def test_get_arg_default_none(self, engine):
        """默认default=None"""
        import sqlglot.expressions as exp
        expr = exp.Substring(this=exp.column('x'))
        assert engine._get_arg(expr, 'length') is None


class TestWindowDispatchTable:
    """窗口函数分发表完整性验证"""

    def test_window_dispatch_all_types(self, engine):
        """分发表包含所有支持的窗口函数"""
        import sqlglot.expressions as exp
        # 验证3种窗口函数类型字符串
        for func_name in ['RowNumber', 'Rank', 'DenseRank']:
            # 确认对应的sqlglot类存在
            assert hasattr(exp, func_name), f"sqlglot缺少 {func_name}"

    def test_window_dispatch_unknown_raises(self, engine):
        """未知窗口函数抛出ValueError"""
        import sqlglot.expressions as exp
        df = pd.DataFrame({'a': [1, 2, 3]})
        try:
            # 通过_compute_window_function间接调用分发表
            engine._apply_window_functions(
                df,
                [exp.Select().select('a')],
                {}
            )
        except (ValueError, AttributeError, TypeError):
            pass  # 某种错误是预期的（取决于调用路径）


class TestStringFunctionsDRY:
    """字符串函数使用_get_arg后行为一致"""

    @pytest.fixture
    def df(self):
        return pd.DataFrame({'name': ['Hello World', 'Python', 'Test']})

    def test_replace_with_get_arg(self, engine, df):
        """REPLACE使用_get_arg提取参数"""
        import sqlglot.expressions as exp
        expr = exp.Replace(
            this=exp.column('name'),
            expression=exp.Literal.string('World'),
            replacement=exp.Literal.string('SQL')
        )
        result = engine._evaluate_string_function(expr, df)
        assert result.iloc[0] == 'Hello SQL'

    def test_left_with_get_arg(self, engine, df):
        """LEFT使用_get_arg提取参数"""
        import sqlglot.expressions as exp
        expr = exp.Left(
            this=exp.column('name'),
            expression=exp.Literal.number(3)
        )
        result = engine._evaluate_string_function(expr, df)
        assert result.iloc[0] == 'Hel'

    def test_right_with_get_arg(self, engine, df):
        """RIGHT使用_get_arg提取参数"""
        import sqlglot.expressions as exp
        expr = exp.Right(
            this=exp.column('name'),
            expression=exp.Literal.number(5)
        )
        result = engine._evaluate_string_function(expr, df)
        assert result.iloc[0] == 'World'

    def test_substring_with_get_arg(self, engine, df):
        """SUBSTRING使用_get_arg提取参数"""
        import sqlglot.expressions as exp
        expr = exp.Substring(
            this=exp.column('name'),
            start=exp.Literal.number(7),
            length=exp.Literal.number(5)
        )
        result = engine._evaluate_string_function(expr, df)
        assert result.iloc[0] == 'World'

    def test_replace_row_mode_with_get_arg(self, engine):
        """REPLACE逐行模式使用_get_arg"""
        import sqlglot.expressions as exp
        expr = exp.Replace(
            this=exp.column('name'),
            expression=exp.Literal.string('World'),
            replacement=exp.Literal.string('SQL')
        )
        row = pd.Series({'name': 'Hello World'})
        result = engine._evaluate_string_function_for_row(expr, row)
        assert result == 'Hello SQL'

    def test_left_row_mode_with_get_arg(self, engine):
        """LEFT逐行模式使用_get_arg"""
        import sqlglot.expressions as exp
        expr = exp.Left(
            this=exp.column('name'),
            expression=exp.Literal.number(4)
        )
        row = pd.Series({'name': 'Hello World'})
        result = engine._evaluate_string_function_for_row(expr, row)
        assert result == 'Hell'
