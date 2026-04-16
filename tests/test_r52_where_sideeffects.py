"""P3-CONSIST-01: WHERE子句DataFrame副作用测试

验证 _apply_where_clause 是否对输入 DataFrame 产生意外副作用：
1. 原始 DF 的列不应被修改（增/删）
2. 原始 DF 的数据行不应被修改
3. WHERE 结果应正确（不受临时列清理影响）
4. CAST WHERE 条件应正确工作（触发临时列路径）
"""
import pytest
import pandas as pd
import numpy as np
import os
import sys
import sqlglot
from sqlglot import exp as exp

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine


@pytest.fixture
def query_engine():
    """创建查询引擎实例"""
    engine = AdvancedSQLQueryEngine()
    return engine


def _parse(sql):
    """辅助函数：解析SQL为sqlglot表达式"""
    return sqlglot.parse_one(sql, dialect="mysql")


@pytest.fixture
def sample_df():
    """创建测试用 DataFrame（含字符串数字列，可触发CAST）"""
    return pd.DataFrame({
        'id': [1, 2, 3, 4, 5],
        'name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
        'score': ['95', '87', '72', '88', '91'],  # 字符串类型数字，触发CAST
        'dept': ['Engineering', 'Sales', 'Engineering', 'Marketing', 'Sales']
    })


class TestWhereClauseSideEffects:
    """WHERE子句副作用测试套件"""

    def test_no_where_does_not_mutate_df(self, query_engine, sample_df):
        """无WHERE条件时，原始DF不应被修改"""
        original_cols = list(sample_df.columns)
        original_shape = sample_df.shape
        # 使用简单SQL（无WHERE）
        parsed = _parse("SELECT * FROM data LIMIT 1")
        result = query_engine._apply_where_clause(parsed, sample_df.copy())
        # 验证原始DF未被修改
        assert list(sample_df.columns) == original_cols, "列被修改了"
        assert sample_df.shape == original_shape, "形状被修改了"

    def test_simple_where_does_not_mutate_original(self, query_engine, sample_df):
        """简单WHERE条件（无CAST）不应修改原始DF的列"""
        df_input = sample_df.copy()
        original_cols = list(df_input.columns)
        
        parsed = _parse("SELECT * FROM data WHERE id > 2")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        # 原始DF列不变
        assert list(df_input.columns) == original_cols, f"原始列被修改: {list(df_input.columns)} vs {original_cols}"
        # 不应有_tmp_columns属性残留
        assert not hasattr(df_input, '_tmp_columns'), "_tmp_columns属性未清理"
        # 结果正确
        assert len(result) == 3, f"应返回3行，实际{len(result)}"
        assert list(result['id']) == [3, 4, 5]

    def test_cast_where_temp_cols_cleaned(self, query_engine, sample_df):
        """CAST WHERE条件产生的临时列应在返回后被清理"""
        df_input = sample_df.copy()
        original_cols = list(df_input.columns)
        
        # CAST(score AS FLOAT) > 85 触发临时列路径
        parsed = _parse("SELECT * FROM data WHERE CAST(score AS FLOAT) > 85")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        # 临时列必须清理干净
        assert list(df_input.columns) == original_cols, \
            f"临时列未清理! 当前列: {list(df_input.columns)}, 原始: {original_cols}"
        assert not hasattr(df_input, '_tmp_columns'), "_tmp_columns属性未清理"
        # 结果应正确（score > 85 的行: 95, 87, 88, 91 → 4行）
        assert len(result) == 4, f"CAST WHERE结果错误: 期望4行, 实际{len(result)}"

    def test_cast_where_result_correctness(self, query_engine, sample_df):
        """CAST WHERE结果数据应完全正确"""
        parsed = _parse("SELECT * FROM data WHERE CAST(score AS FLOAT) > 90")
        result = query_engine._apply_where_clause(parsed, sample_df.copy())
        
        # score > 90: Alice(95), Eve(91) → 2行
        assert len(result) == 2, f"期望2行, 实际{len(result)}"
        result_names = sorted(result['name'].tolist())
        assert result_names == ['Alice', 'Eve'], f"名字错误: {result_names}"

    def test_where_with_function_expr(self, query_engine, sample_df):
        """函数表达式WHERE（如UPPER(name)='ALICE'）不应泄漏临时列"""
        df_input = sample_df.copy()
        original_cols = list(df_input.columns)
        
        # 函数表达式也可能触发临时列路径
        parsed = _parse("SELECT * FROM data WHERE name = 'Alice'")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert list(df_input.columns) == original_cols, \
            f"列被修改: {list(df_input.columns)} vs {original_cols}"
        assert len(result) == 1, f"应返回1行, 实际{len(result)}"
        assert result.iloc[0]['name'] == 'Alice'

    def test_complex_where_fallback_does_not_leak(self, query_engine, sample_df):
        """复杂WHERE回退到逐行过滤时也不应泄漏临时列"""
        df_input = sample_df.copy()
        original_cols = list(df_input.columns)
        
        # 如果内部出错导致回退到_apply_row_filter
        parsed = _parse("SELECT * FROM data WHERE dept = 'Engineering'")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert list(df_input.columns) == original_cols
        assert len(result) == 2
        assert set(result['name'].tolist()) == {'Alice', 'Charlie'}

    def test_multiple_cast_conditions(self, query_engine):
        """多个CAST条件应各自创建和清理临时列"""
        df = pd.DataFrame({
            'a': ['10', '20', '30', '40', '50'],
            'b': ['1.5', '2.5', '3.5', '4.5', '5.5'],
            'cat': ['X', 'Y', 'X', 'Y', 'X']
        })
        df_input = df.copy()
        original_cols = list(df_input.columns)
        
        parsed = _parse(
            "SELECT * FROM data WHERE CAST(a AS INTEGER) > 15 AND CAST(b AS FLOAT) < 5.0"
        )
        result = query_engine._apply_where_clause(parsed, df_input)
        
        # 所有临时列清理完毕
        assert list(df_input.columns) == original_cols, \
            f"临时列残留: {[c for c in df_input.columns if c not in original_cols]}"
        # a>15 AND b<5.0: (20,2.5), (30,3.5), (40,4.5) → 3行
        assert len(result) == 3, f"期望3行, 实际{len(result)}, 数据:\n{result}"

    def test_where_data_integrity(self, query_engine, sample_df):
        """WHERE处理后原始df的数据内容不变"""
        df_input = sample_df.copy()
        original_data = df_input.copy()
        
        parsed = _parse("SELECT * FROM data WHERE id >= 3")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        # 原始df的所有值都保持不变
        pd.testing.assert_frame_equal(df_input, original_data)


class TestWhereEdgeCases:
    """WHERE边界case测试"""

    def test_empty_result_where(self, query_engine, sample_df):
        """WHERE返回空集时不应有副作用"""
        df_input = sample_df.copy()
        
        parsed = _parse("SELECT * FROM data WHERE id > 999")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert len(result) == 0
        assert list(df_input.columns) == list(sample_df.columns)

    def test_where_all_rows_match(self, query_engine, sample_df):
        """所有行匹配WHERE时结果完整"""
        df_input = sample_df.copy()
        
        parsed = _parse("SELECT * FROM data WHERE id >= 1")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert len(result) == 5

    def test_where_like_pattern(self, query_engine, sample_df):
        """LIKE模式WHERE不应有副作用"""
        df_input = sample_df.copy()
        original_cols = list(df_input.columns)
        
        parsed = _parse("SELECT * FROM data WHERE name LIKE 'A%'")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert list(df_input.columns) == original_cols
        assert len(result) == 1
        assert result.iloc[0]['name'] == 'Alice'

    def test_where_is_null(self, query_engine):
        """IS NULL WHERE不应有副作用"""
        df = pd.DataFrame({
            'id': [1, 2, 3],
            'val': ['hello', None, 'world']
        })
        df_input = df.copy()
        
        parsed = _parse("SELECT * FROM data WHERE val IS NULL")
        result = query_engine._apply_where_clause(parsed, df_input)
        
        assert len(result) == 1
        assert list(df_input.columns) == list(df.columns)


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
