"""差分对抗测试：ExcelMCP vs SQLite Oracle。

每个测试同时操作 ExcelMCP 和 SQLite，对比结果。
SQLite 是真值来源，不一致即 bug。
"""

from __future__ import annotations

import random
import sqlite3
import time

import pytest

from .conftest import (
    SEED_DATA,
    SEED_TABLE,
    ScoreCollector,
    assert_affected_rows_match,
    assert_query_match,
    query_excel,
    query_sqlite,
    write_excel,
    write_sqlite,
)

# ============================================================
# TestAdversarialUpdateReadback
# ============================================================


class TestAdversarialUpdateReadback:
    """UPDATE 写后读一致性：写完两边结果必须相同。"""

    def test_update_single_col_readback(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=999 WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            # 读回验证
            sel = "SELECT * FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("UPDATE", "single_col", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("UPDATE", "single_col", sql, True)

    def test_update_multi_col_readback(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=888, Stock=99 WHERE ID=2"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT * FROM 商品 WHERE ID=2"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("UPDATE", "multi_col", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("UPDATE", "multi_col", sql, True)

    def test_update_expression_readback(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=Price*1.1 WHERE Active='是'"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT * FROM 商品 WHERE Active='是'"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel, tol=0.1)
        except AssertionError as e:
            score_collector.record("UPDATE", "expression", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("UPDATE", "expression", sql, True)

    def test_update_no_match(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=999 WHERE ID=99999"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            assert sr == 0, f"SQLite should affect 0 rows, got {sr}"
            # 验证数据没变
            sel = "SELECT * FROM 商品"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("UPDATE", "no_match", sql, False, 0, er.get("affected_rows"), str(e))
            raise
        score_collector.record("UPDATE", "no_match", sql, True)

    def test_update_all_rows(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Stock=0"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT * FROM 商品"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("UPDATE", "all_rows", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("UPDATE", "all_rows", sql, True)


# ============================================================
# TestAdversarialInsertDelete
# ============================================================


class TestAdversarialInsertDelete:
    """INSERT/DELETE 行数守恒和读回一致性。"""

    def test_insert_readback(self, both, score_collector):
        excel_path, conn = both
        sql = "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (7, '圣剑', 500, 10, '是')"

        er = write_excel(excel_path, sql, "insert")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT * FROM 商品 WHERE ID=7"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
            # COUNT(*) 一致
            sel2 = "SELECT COUNT(*) FROM 商品"
            assert_query_match(query_excel(excel_path, sel2), query_sqlite(conn, sel2), sel2)
        except AssertionError as e:
            score_collector.record("INSERT", "basic", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("INSERT", "basic", sql, True)

    def test_delete_readback(self, both, score_collector):
        excel_path, conn = both
        sql = "DELETE FROM 商品 WHERE ID=3"

        er = write_excel(excel_path, sql, "delete")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT * FROM 商品"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
            sel2 = "SELECT COUNT(*) FROM 商品"
            assert_query_match(query_excel(excel_path, sel2), query_sqlite(conn, sel2), sel2)
        except AssertionError as e:
            score_collector.record("DELETE", "basic", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("DELETE", "basic", sql, True)

    def test_affected_rows_accuracy(self, both, score_collector):
        """各种写操作后 affected_rows == SQLite rowcount"""
        excel_path, conn = both

        cases = [
            ("UPDATE", "UPDATE 商品 SET Price=100 WHERE ID=1", "update"),
            ("UPDATE", "UPDATE 商品 SET Stock=0 WHERE Active='否'", "update"),
            ("UPDATE", "UPDATE 商品 SET Name='改名' WHERE ID=99999", "update"),
            ("INSERT", "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (10, 'X', 1, 1, '是')", "insert"),
            ("DELETE", "DELETE FROM 商品 WHERE ID=10", "delete"),
            ("DELETE", "DELETE FROM 商品 WHERE ID=99999", "delete"),
        ]

        for category, sql, op_type in cases:
            er = write_excel(excel_path, sql, op_type)
            sr = write_sqlite(conn, sql)
            try:
                assert_affected_rows_match(er, sr, sql)
            except AssertionError as e:
                score_collector.record(category, "affected_rows", sql, False, sr, er.get("affected_rows"), str(e))
                raise
            score_collector.record(category, "affected_rows", sql, True)


# ============================================================
# TestAdversarialWhereConditions
# ============================================================


class TestAdversarialWhereConditions:
    """WHERE 条件精确性：各种 WHERE 条件两边匹配的行集一致。"""

    def test_where_equal(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE ID=2"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "equal", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "equal", sql, True)

    def test_where_greater(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE Price > 150"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "greater", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "greater", sql, True)

    def test_where_less(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE Stock < 50"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "less", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "less", sql, True)

    def test_where_in(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE ID IN (1, 3, 5)"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "in", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "in", sql, True)

    def test_where_like(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE Name LIKE '%剑%'"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "like", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "like", sql, True)

    def test_where_and_or(self, both, score_collector):
        excel_path, conn = both
        sql = "SELECT * FROM 商品 WHERE (Price > 100 AND Active='是') OR ID=3"

        try:
            assert_query_match(query_excel(excel_path, sql), query_sqlite(conn, sql), sql)
        except AssertionError as e:
            score_collector.record("WHERE", "and_or", sql, False, error_msg=str(e))
            raise
        score_collector.record("WHERE", "and_or", sql, True)


# ============================================================
# TestAdversarialEdgeValues
# ============================================================


class TestAdversarialEdgeValues:
    """边界值对抗：浮点精度、负数、零、空字符串、特殊字符、中文、NULL。"""

    def test_float_precision(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=999.99*1.1 WHERE ID=6"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Price FROM 商品 WHERE ID=6"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel, tol=0.1)
        except AssertionError as e:
            score_collector.record("EDGE", "float_precision", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "float_precision", sql, True)

    def test_negative(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=-100 WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Price FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "negative", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "negative", sql, True)

    def test_zero(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=0 WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Price FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "zero", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "zero", sql, True)

    def test_empty_string(self, both, score_collector):
        """空字符串 '' 写入后读回为 None — Excel 存储层固有限制，
        差分测试中归一化 '' == None（见 _normalize_cell）。"""
        excel_path, conn = both
        sql = "UPDATE 商品 SET Name='' WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Name FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "empty_string", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "empty_string", sql, True)

    def test_special_char(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Name='O''Brien' WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Name FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "special_char", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "special_char", sql, True)

    def test_chinese(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Name='神剑·破军' WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Name FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "chinese", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "chinese", sql, True)

    def test_null_write(self, both, score_collector):
        excel_path, conn = both
        sql = "UPDATE 商品 SET Price=NULL WHERE ID=1"

        er = write_excel(excel_path, sql, "update")
        sr = write_sqlite(conn, sql)

        try:
            assert_affected_rows_match(er, sr, sql)
            sel = "SELECT Price FROM 商品 WHERE ID=1"
            assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
        except AssertionError as e:
            score_collector.record("EDGE", "null_write", sql, False, sr, er.get("affected_rows"), str(e))
            raise
        score_collector.record("EDGE", "null_write", sql, True)


# ============================================================
# TestAdversarialRandomFuzz
# ============================================================


class TestAdversarialRandomFuzz:
    """随机写入对抗：随机生成 20 个写操作，每次写后验证一致性。"""

    def test_random_write_sequence(self, both, score_collector):
        excel_path, conn = both
        random.seed(42)

        # 生成随机写操作池
        update_ops = [
            (
                "UPDATE 商品 SET Price={val} WHERE ID={rid}",
                "update",
                lambda: {
                    "val": random.choice([0, -50, 99.99, 1000, 0.01, 500]),
                    "rid": random.randint(1, 6),
                },
            ),
            (
                "UPDATE 商品 SET Stock={val} WHERE ID={rid}",
                "update",
                lambda: {
                    "val": random.randint(0, 200),
                    "rid": random.randint(1, 6),
                },
            ),
            (
                "UPDATE 商品 SET Name='{val}' WHERE ID={rid}",
                "update",
                lambda: {
                    "val": random.choice(["测试", "Test", "O''Brien", "A" * 50]),
                    "rid": random.randint(1, 6),
                },
            ),
            ("UPDATE 商品 SET Price=Price*1.1 WHERE Active='是'", "update", lambda: {}),
            ("UPDATE 商品 SET Stock=Stock+10", "update", lambda: {}),
        ]

        insert_ops = [
            (
                "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES ({rid}, '{name}', {price}, {stock}, '{active}')",
                "insert",
                lambda: {
                    "rid": random.randint(100, 999),
                    "name": random.choice(["随机物品", "RandItem", "空名"]),
                    "price": round(random.uniform(1, 1000), 2),
                    "stock": random.randint(0, 100),
                    "active": random.choice(["是", "否"]),
                },
            ),
        ]

        delete_ops = [
            (
                "DELETE FROM 商品 WHERE ID={rid}",
                "delete",
                lambda: {
                    "rid": random.randint(100, 999),  # 大概率不匹配
                },
            ),
        ]

        all_ops = update_ops + insert_ops + delete_ops

        for i in range(20):
            template, op_type, gen_params = random.choice(all_ops)
            params = gen_params()
            sql = template.format(**params)

            # 写入两边
            er = write_excel(excel_path, sql, op_type)
            sr = write_sqlite(conn, sql)

            try:
                assert_affected_rows_match(er, sr, sql)
                # 写后读：全表扫描验证
                sel = "SELECT * FROM 商品"
                assert_query_match(query_excel(excel_path, sel), query_sqlite(conn, sel), sel)
            except AssertionError as e:
                score_collector.record("FUZZ", f"step_{i}", sql, False, sr, er.get("affected_rows"), str(e))
                # 不 raise — 继续后续操作看看还会出什么问题
            else:
                score_collector.record("FUZZ", f"step_{i}", sql, True)

        # 最终检查：如果有失败就整体失败
        score_collector.summary()  # 触发可能的内部汇总副作用
        fuzz_failures = [r for r in score_collector.results if r["category"] == "FUZZ" and not r["passed"]]
        assert len(fuzz_failures) == 0, f"Random fuzz found {len(fuzz_failures)} mismatches out of 20 operations:\n" + "\n".join(
            f"  [{r['sub_category']}] {r['sql']}\n    expected={r['expected']} actual={r['actual']}\n    {r['error_msg']}" for r in fuzz_failures
        )
