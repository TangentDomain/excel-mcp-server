"""测试 Literal 解析 DRY 统一：_extract_literal_value 和 _parse_literal_value 委托关系"""
import pytest
import sys
import os

# 添加源码路径（worktree兼容）
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from src.excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def engine():
    return AdvancedSQLQueryEngine()


class TestExtractLiteralValueDelegation:
    """验证 _extract_literal_value 委托 _parse_literal_value"""

    def test_integer_literal(self, engine):
        """整数字面量通过委托正确解析"""
        import sqlglot.expressions as exp
        lit = exp.Literal.number(42)
        result = engine._extract_literal_value(lit)
        assert result == 42
        assert isinstance(result, int)

    def test_float_literal(self, engine):
        """浮点数字面量通过委托正确解析"""
        import sqlglot.expressions as exp
        lit = exp.Literal.number(3.14)
        result = engine._extract_literal_value(lit)
        assert result == 3.14
        assert isinstance(result, float)

    def test_string_literal(self, engine):
        """字符串字面量通过委托正确解析"""
        import sqlglot.expressions as exp
        lit = exp.Literal.string("hello")
        result = engine._extract_literal_value(lit)
        assert result == "hello"
        assert isinstance(result, str)

    def test_non_literal_returns_none(self, engine):
        """非Literal表达式返回None"""
        import sqlglot.expressions as exp
        col = exp.Column(this=exp.Identifier(this="test"))
        result = engine._extract_literal_value(col)
        assert result is None

    def test_integer_string(self, engine):
        """纯整数字符串解析为int而非float"""
        import sqlglot.expressions as exp
        lit = exp.Literal.number(100)
        result = engine._extract_literal_value(lit)
        assert result == 100
        assert isinstance(result, int)

    def test_consistency_with_parse_literal_value(self, engine):
        """_extract_literal_value 和 _parse_literal_value 对 Literal 返回相同结果"""
        import sqlglot.expressions as exp
        test_cases = [
            exp.Literal.number(42),
            exp.Literal.number(3.14),
            exp.Literal.string("test"),
        ]
        for lit in test_cases:
            assert engine._extract_literal_value(lit) == engine._parse_literal_value(lit)


class TestSelectLiteralParsesCorrectly:
    """验证 _apply_select_expressions 中 Literal 使用 _parse_literal_value"""

    def test_select_literal_integer(self, engine):
        """SELECT 1 返回整数而非字符串"""
        import pandas as pd
        import sqlglot
        df = pd.DataFrame({'a': [1, 2, 3]})
        parsed = sqlglot.parse_one("SELECT 1 AS one, a FROM df", read='mysql')
        result = engine._apply_select_expressions(parsed, df)
        assert 'one' in result.columns
        # 应该是int(1)而非字符串'1'
        assert result['one'].iloc[0] == 1
        assert result['one'].iloc[0] != '1'  # 确认不是字符串

    def test_select_literal_float(self, engine):
        """SELECT 3.14 返回浮点数"""
        import pandas as pd
        import sqlglot
        df = pd.DataFrame({'a': [1, 2, 3]})
        parsed = sqlglot.parse_one("SELECT 3.14 AS pi, a FROM df", read='mysql')
        result = engine._apply_select_expressions(parsed, df)
        assert result['pi'].iloc[0] == 3.14

    def test_select_literal_string(self, engine):
        """SELECT 'hello' 返回字符串"""
        import pandas as pd
        import sqlglot
        df = pd.DataFrame({'a': [1, 2, 3]})
        parsed = sqlglot.parse_one("SELECT 'hello' AS greeting, a FROM df", read='mysql')
        result = engine._apply_select_expressions(parsed, df)
        assert result['greeting'].iloc[0] == 'hello'


class TestCoalesceLiteralParsesCorrectly:
    """验证 COALESCE 向量化中 Literal 使用 _parse_literal_value"""

    def test_coalesce_with_integer_literal(self, engine):
        """COALESCE(col, 0) 中 0 应为整数0而非字符串'0'"""
        import pandas as pd
        import sqlglot
        df = pd.DataFrame({'val': [1, None, 3]})
        parsed = sqlglot.parse_one("SELECT COALESCE(val, 0) AS result FROM df", read='mysql')
        result = engine._apply_select_expressions(parsed, df)
        # None值应被替换为整数0
        assert result['result'].iloc[1] == 0
        assert isinstance(result['result'].iloc[1], int)

    def test_coalesce_with_string_literal(self, engine):
        """COALESCE(col, 'N/A') 中 'N/A' 应为字符串"""
        import pandas as pd
        import sqlglot
        df = pd.DataFrame({'val': ['a', None, 'c']})
        parsed = sqlglot.parse_one("SELECT COALESCE(val, 'N/A') AS result FROM df", read='mysql')
        result = engine._apply_select_expressions(parsed, df)
        assert result['result'].iloc[1] == 'N/A'


class TestExtractSelectAliasesDelegation:
    """验证 _extract_select_aliases 委托 _extract_select_alias"""

    def test_alias_consistency(self, engine):
        """_extract_select_aliases 和逐个 _extract_select_alias 结果一致"""
        import sqlglot
        sql = "SELECT a, b AS alias_b, COUNT(*) AS cnt, SUM(x) FROM t"
        parsed = sqlglot.parse_one(sql, read='mysql')
        aliases_dict = engine._extract_select_aliases(parsed)
        for i, select_expr in enumerate(parsed.expressions):
            alias_name, original_expr = engine._extract_select_alias(select_expr, i)
            assert alias_name in aliases_dict
            assert aliases_dict[alias_name] == original_expr

    def test_star_excluded(self, engine):
        """SELECT * 不应出现在别名映射中"""
        import sqlglot
        sql = "SELECT *, a AS alias_a FROM t"
        parsed = sqlglot.parse_one(sql, read='mysql')
        aliases = engine._extract_select_aliases(parsed)
        assert '*' not in aliases
        assert 'alias_a' in aliases
