"""
R44 Edge Case Tests - Window functions, ORDER BY expressions, NULL handling,
LIMIT/OFFSET, DISTINCT, IN/BETWEEN/LIKE, special characters, complex SQL

Covers: R12 (window on empty table), R13 (ORDER BY expressions), 
plus comprehensive edge cases discovered during R44 deep testing.
"""
import pytest
import pandas as pd
import tempfile
import os
from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query


# ============================================================
# Fixtures
# ============================================================

@pytest.fixture
def normal_file():
    """Normal 5-row table with various data types"""
    df = pd.DataFrame({
        "ID": [1, 2, 3, 4, 5],
        "Name": ["Alice", "Bob", "Charlie", "Diana", "Eve"],
        "Score": [85, 92, 78, 95, 88],
        "Category": ["A", "B", "A", "B", "C"],
    })
    f = tempfile.mktemp(suffix=".xlsx")
    df.to_excel(f, index=False, sheet_name="Sheet1")
    yield f
    os.unlink(f)


@pytest.fixture
def empty_file():
    """Empty table (0 rows)"""
    df = pd.DataFrame({"ID": [], "Name": [], "Score": []})
    f = tempfile.mktemp(suffix=".xlsx")
    df.to_excel(f, index=False, sheet_name="Sheet1")
    yield f
    os.unlink(f)


@pytest.fixture
def single_file():
    """Single row table"""
    df = pd.DataFrame({"ID": [1], "Name": ["Solo"], "Score": [100]})
    f = tempfile.mktemp(suffix=".xlsx")
    df.to_excel(f, index=False, sheet_name="Sheet1")
    yield f
    os.unlink(f)


@pytest.fixture
def null_file():
    """Table with NULL values"""
    df = pd.DataFrame({
        "ID": [1, 2, 3, 4, 5],
        "Name": ["Alice", None, "Charlie", None, "Eve"],
        "Score": [85, None, 78, None, 88],
    })
    f = tempfile.mktemp(suffix=".xlsx")
    df.to_excel(f, index=False, sheet_name="Sheet1")
    yield f
    os.unlink(f)


@pytest.fixture
def special_file():
    """Table with special characters in data"""
    df = pd.DataFrame({
        "ID": [1, 2, 3],
        "Name": ["O'Brien", 'Annie "Bo"', 'Test \\\\ slash'],
        "Value": [10, 20, 30],
    })
    f = tempfile.mktemp(suffix=".xlsx")
    df.to_excel(f, index=False, sheet_name="Sheet1")
    yield f
    os.unlink(f)


# ============================================================
# Helpers
# ============================================================

def row_count(data):
    """Return number of data rows (excluding header row at index 0)"""
    return max(0, len(data) - 1)


def get_row(data, idx):
    """Get data row by 0-based index (idx=0 → first data row)"""
    if len(data) > idx + 1:
        return data[idx + 1]
    return None


def query(file_path, sql):
    """Execute query and return result dict"""
    return execute_advanced_sql_query(file_path, sql)


# ============================================================
# R12: Window Function Edge Cases
# ============================================================

class TestR12WindowEdgeCases:
    """R12: Window functions on empty/single-row tables"""

    def test_row_number_on_empty_table(self, empty_file):
        """Empty table + ROW_NUMBER should return header-only (0 data rows)"""
        r = query(empty_file, "SELECT ID, Name, ROW_NUMBER() OVER (ORDER BY ID) as rn FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 0

    def test_rank_on_empty_table(self, empty_file):
        """Empty table + RANK should return 0 data rows"""
        r = query(empty_file, "SELECT ID, Score, RANK() OVER (ORDER BY Score DESC) as rk FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 0

    def test_sum_over_on_empty_table(self, empty_file):
        """Empty table + SUM OVER () should return 0 data rows"""
        r = query(empty_file, "SELECT ID, Score, SUM(Score) OVER () as total FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 0

    def test_window_function_on_single_row(self, single_file):
        """Single row + window function should return 1 data row"""
        r = query(single_file, "SELECT ID, Name, ROW_NUMBER() OVER (ORDER BY ID) as rn FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 1
        assert get_row(r["data"], 0)[-1] == 1  # rn = 1


# ============================================================
# R13: ORDER BY Expression Cases
# ============================================================

class TestR13OrderByExpressions:
    """R13: ORDER BY with arithmetic/CASE WHEN/alias expressions"""

    def test_order_by_arithmetic_expr(self, normal_file):
        """ORDER BY Score + 0 DESC should work (arithmetic expression)"""
        r = query(normal_file, "SELECT ID, Name, Score FROM Sheet1 ORDER BY Score + 0 DESC")
        assert r["success"]
        assert row_count(r["data"]) == 5

    def test_order_by_score_desc(self, normal_file):
        """ORDER BY Score DESC should sort correctly"""
        r = query(normal_file, "SELECT ID, Name, Score FROM Sheet1 ORDER BY Score DESC")
        assert r["success"]
        scores = [get_row(r["data"], i)[2] for i in range(row_count(r["data"]))]
        assert scores == sorted(scores, reverse=True)

    def test_order_by_case_when(self, normal_file):
        """ORDER BY CASE WHEN should sort Category A first, then others by Score DESC"""
        r = query(normal_file, """SELECT ID, Name, Category, Score FROM Sheet1 
            ORDER BY CASE WHEN Category='A' THEN 0 ELSE 1 END, Score DESC""")
        assert r["success"]
        # First two rows should be Category='A'
        assert get_row(r["data"], 0)[2] == "A"
        assert get_row(r["data"], 1)[2] == "A"
        # Remaining rows should not be 'A'
        assert get_row(r["data"], 2)[2] != "A"

    def test_order_by_column_alias(self, normal_file):
        """ORDER BY column alias from SELECT should work"""
        r = query(normal_file, "SELECT ID, Name, Score*2 as DoubleScore FROM Sheet1 ORDER BY DoubleScore DESC")
        assert r["success"]
        assert row_count(r["data"]) == 5


# ============================================================
# NULL Handling Edge Cases
# ============================================================

class TestNullHandling:
    """NULL value handling in aggregations and GROUP BY"""

    def test_count_ignores_nulls(self, null_file):
        """COUNT(col) should ignore NULL values"""
        r = query(null_file, "SELECT COUNT(Name) as CntName, COUNT(Score) as CntScore, COUNT(*) as Total FROM Sheet1")
        assert r["success"]
        assert get_row(r["data"], 0) == [3, 3, 5]  # 3 non-null names, 3 non-null scores, 5 total rows

    def test_avg_sum_ignore_nulls(self, null_file):
        """AVG/SUM should ignore NULL values"""
        r = query(null_file, "SELECT AVG(Score) as AvgS, SUM(Score) as SumS FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 1
        # AVG of [85, 78, 88] = 83.67, SUM = 251
        row = get_row(r["data"], 0)
        assert abs(row[0] - 83.66666666666667) < 0.001
        assert row[1] == 251

    def test_group_by_with_null_column(self, null_file):
        """GROUP BY with NULL column should include NULL group"""
        r = query(null_file, "SELECT Name, COUNT(*) as Cnt FROM Sheet1 GROUP BY Name")
        assert r["success"]
        assert row_count(r["data"]) >= 3  # Alice, Charlie, Eve + NULL group


# ============================================================
# Complex SQL Combinations
# ============================================================

class TestComplexSQLCombinations:
    """Complex SQL combining multiple features"""

    def test_cte_with_window_function(self, normal_file):
        """CTE + window function + WHERE filter"""
        r = query(normal_file, """WITH Ranked AS (
            SELECT *, ROW_NUMBER() OVER (ORDER BY Score DESC) as rn 
            FROM Sheet1
        ) SELECT * FROM Ranked WHERE rn <= 3""")
        assert r["success"]
        assert row_count(r["data"]) == 3
        # Top 3 by score: Diana(95), Bob(92), Eve(88)
        scores = [get_row(r["data"], i)[2] for i in range(3)]
        assert scores == [95, 92, 88]

    def test_subquery_in_where_with_aggregation(self, normal_file):
        """Subquery in WHERE clause with aggregation comparison"""
        r = query(normal_file, """SELECT Category, AVG(Score) as AvgSc 
            FROM Sheet1 
            WHERE Score > (SELECT AVG(Score) FROM Sheet1) 
            GROUP BY Category""")
        assert r["success"]
        assert row_count(r["data"]) > 0

    def test_multiple_aggregation_functions(self, normal_file):
        """Multiple different aggregation functions in one query"""
        r = query(normal_file, """SELECT Category, 
            COUNT(*) as Cnt, SUM(Score) as Total, AVG(Score) as AvgSc, 
            MAX(Score) as MaxSc, MIN(Score) as MinSc
            FROM Sheet1 GROUP BY Category""")
        assert r["success"]
        assert row_count(r["data"]) == 3  # A, B, C groups

    def test_nested_case_when_in_aggregation(self, normal_file):
        """CASE WHEN inside SUM aggregation"""
        r = query(normal_file, """SELECT 
            SUM(CASE WHEN Score >= 90 THEN 1 ELSE 0 END) as HighScorers,
            SUM(CASE WHEN Score < 80 THEN 1 ELSE 0 END) as LowScorers,
            COUNT(*) as Total
            FROM Sheet1""")
        assert r["success"]
        assert row_count(r["data"]) == 1
        assert get_row(r["data"], 0) == [2, 1, 5]  # 2 high(>=90), 1 low(<80), 5 total


# ============================================================
# LIMIT/OFFSET Edge Cases
# ============================================================

class TestLimitOffsetEdgeCases:
    """LIMIT/OFFSET boundary conditions"""

    def test_limit_zero(self, normal_file):
        """LIMIT 0 should return 0 data rows"""
        r = query(normal_file, "SELECT * FROM Sheet1 LIMIT 0")
        assert r["success"]
        assert row_count(r["data"]) == 0

    def test_limit_exceeds_row_count(self, normal_file):
        """LIMIT larger than row count returns all rows"""
        r = query(normal_file, "SELECT * FROM Sheet1 LIMIT 100")
        assert r["success"]
        assert row_count(r["data"]) == 5

    def test_offset_with_limit(self, normal_file):
        """OFFSET 2 LIMIT 2 should skip first 2, return next 2"""
        r = query(normal_file, "SELECT * FROM Sheet1 ORDER BY ID LIMIT 2 OFFSET 2")
        assert r["success"]
        assert row_count(r["data"]) == 2
        assert get_row(r["data"], 0)[0] == 3  # ID=3 (after skipping 1,2)

    def test_offset_exceeds_row_count(self, normal_file):
        """OFFSET beyond all rows returns 0 data rows"""
        r = query(normal_file, "SELECT * FROM Sheet1 ORDER BY ID LIMIT 10 OFFSET 100")
        assert r["success"]
        assert row_count(r["data"]) == 0


# ============================================================
# DISTINCT Edge Cases
# ============================================================

class TestDistinctEdgeCases:
    """DISTINCT behavior"""

    def test_distinct_unique_values(self, normal_file):
        """DISTINCT on Category should return 3 unique values"""
        r = query(normal_file, "SELECT DISTINCT Category FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 3  # A, B, C

    def test_distinct_constant(self, normal_file):
        """DISTINCT constant should return 1 row"""
        r = query(normal_file, "SELECT DISTINCT 'X' as Const FROM Sheet1")
        assert r["success"]
        assert row_count(r["data"]) == 1


# ============================================================
# IN/BETWEEN/LIKE Edge Cases
# ============================================================

class TestInBetweenLikeEdgeCases:
    """IN, BETWEEN, LIKE boundary conditions"""

    def test_in_all_matching(self, normal_file):
        """IN with all matching IDs returns all rows"""
        r = query(normal_file, "SELECT * FROM Sheet1 WHERE ID IN (1, 2, 3, 4, 5)")
        assert r["success"]
        assert row_count(r["data"]) == 5

    def test_in_no_matches(self, normal_file):
        """IN with no matching IDs returns 0 rows"""
        r = query(normal_file, "SELECT * FROM Sheet1 WHERE ID IN (99, 100)")
        assert r["success"]
        assert row_count(r["data"]) == 0

    def test_in_subquery(self, normal_file):
        """IN subquery should filter correctly"""
        r = query(normal_file, "SELECT * FROM Sheet1 WHERE ID IN (SELECT ID FROM Sheet1 WHERE Score > 80)")
        assert r["success"]
        assert row_count(r["data"]) == 4  # Scores > 80: 85,92,95,88

    def test_between_inclusive(self, normal_file):
        "BETWEEN should be inclusive of both bounds"
        r = query(normal_file, "SELECT * FROM Sheet1 WHERE Score BETWEEN 80 AND 90")
        assert r["success"]
        assert row_count(r["data"]) >= 2  # 85, 88 in range

    def test_like_pattern_match(self, normal_file):
        "LIKE pattern matching should work"
        r = query(normal_file, "SELECT * FROM Sheet1 WHERE Name LIKE '%a%' OR Name LIKE '%A%'")
        assert r["success"]
        assert row_count(r["data"]) >= 1  # Alice matches


# ============================================================
# Special Characters
# ============================================================

class TestSpecialCharacters:
    """Special characters in data values"""

    def test_single_quote_in_data(self, special_file):
        "Single quote (O'Brien) in data should not break queries"
        r = query(special_file, "SELECT * FROM Sheet1 WHERE ID = 1")
        assert r["success"]
        assert row_count(r["data"]) == 1

    def test_double_quotes_in_data(self, special_file):
        'Double quotes (Annie "Bo") in data should not break'
        r = query(special_file, "SELECT * FROM Sheet1 WHERE ID = 2")
        assert r["success"]
        assert row_count(r["data"]) == 1

    def test_backslash_in_data(self, special_file):
        "Backslash in data should not break"
        r = query(special_file, "SELECT * FROM Sheet1 WHERE ID = 3")
        assert r["success"]
        assert row_count(r["data"]) == 1
