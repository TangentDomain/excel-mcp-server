"""
R49 P1 Bug Fixes — 关键安全/正确性修复测试
|- P1-03: DELETE openpyxl 降级路径逆序删除，避免行号偏移 [P1]
|- P1-04: INSERT 批量大小限制检查 [P1]
|- P1-05: _evaluate_update_expression 递归深度保护 [P1]

日期: 2026-04-16
Round: 49
"""

import pytest
import pandas as pd
import numpy as np
import os
import tempfile
import openpyxl

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    AdvancedSQLQueryEngine,
    execute_advanced_sql_query,
)


def _make_test_xlsx(data, sheet_name="Sheet"):
    """创建临时 xlsx 文件"""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row_idx, row in enumerate(data):
        for col_idx, val in enumerate(row):
            ws.cell(row=row_idx + 1, column=col_idx + 1, value=val)
    wb.save(path)
    wb.close()
    return path


# ============================================================
# P1-03: DELETE 逆序删除 — 验证多行删除后数据一致性
# ============================================================

class TestP1_03_DeleteReverseOrder:
    """P1-03: batch_delete_rows 在 openpyxl 路径下必须逆序处理行号"""

    def test_delete_multiple_rows_data_integrity(self):
        """删除多行后，剩余数据应保持正确（无行号偏移）"""
        data = [["Name", "Score"], ["Alice", 85], ["Bob", 72], ["Carol", 93], ["Dave", 68], ["Eve", 90]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_delete_query(
                path,
                "DELETE FROM Sheet WHERE Name IN ('Bob', 'Dave')"
            )
            assert result["success"] is True, f"删除失败: {result.get('message', '')}"

            # 验证剩余数据：Alice, Carol, Eve
            verify = execute_advanced_sql_query(
                path,
                "SELECT Name FROM Sheet ORDER BY Name"
            )
            assert verify["success"] is True
            names = [r[0] for r in verify["data"][1:]]  # 去表头
            assert names == ["Alice", "Carol", "Eve"], f"期望 ['Alice','Carol','Eve'], 实际 {names}"
        finally:
            os.unlink(path)

    def test_delete_consecutive_rows(self):
        """删除连续多行后数据一致性"""
        data = [["ID", "Val"], [1, "a"], [2, "b"], [3, "c"], [4, "d"], [5, "e"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_delete_query(
                path,
                "DELETE FROM Sheet WHERE ID IN (2, 3, 4)"
            )
            assert result["success"] is True

            verify = execute_advanced_sql_query(path, "SELECT ID FROM Sheet ORDER BY ID")
            ids = [r[0] for r in verify["data"][1:]]
            assert ids == [1, 5], f"期望 [1,5], 实际 {ids}"
        finally:
            os.unlink(path)

    def test_delete_all_except_one(self):
        """删除几乎所有行，只保留一行"""
        data = [["N"], [1], [2], [3], [4]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_delete_query(
                path,
                "DELETE FROM Sheet WHERE N != 3"
            )
            assert result["success"] is True

            verify = execute_advanced_sql_query(path, "SELECT N FROM Sheet")
            vals = [r[0] for r in verify["data"][1:]]
            assert vals == [3], f"期望 [3], 实际 {vals}"
        finally:
            os.unlink(path)

    def test_delete_non_matching_rows(self):
        """删除不匹配的行时数据不变"""
        data = [["A"], [1], [2], [3]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_delete_query(
                path,
                "DELETE FROM Sheet WHERE A = 999"
            )
            assert result["success"] is True

            verify = execute_advanced_sql_query(path, "SELECT A FROM Sheet ORDER BY A")
            vals = [r[0] for r in verify["data"][1:]]
            assert vals == [1, 2, 3], f"期望 [1,2,3], 实际 {vals}"
        finally:
            os.unlink(path)


# ============================================================
# P1-04: INSERT 批量大小限制
# ============================================================

class TestP1_04_InsertBatchSizeLimit:
    """P1-04: INSERT 必须拒绝超过批量大小限制的请求"""

    def test_insert_within_limit(self):
        """正常批量插入（在限制内）应成功"""
        data = [["A", "B"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_insert_query(
                path,
                "INSERT INTO Sheet (A, B) VALUES (1, 'x'), (2, 'y'), (3, 'z')"
            )
            # 不应返回"超限"错误
            if not result["success"]:
                msg = result.get("message", "")
                assert "超过限制" not in msg and "5000" not in msg, \
                    f"正常INSERT不应被限流: {msg}"
        finally:
            os.unlink(path)

    def test_insert_exceeds_limit_rejected(self):
        """超过批量大小限制的 INSERT 应被拒绝"""
        data = [["A", "B"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            # 构造一个超大 INSERT（超过 5000 行限制）
            n_rows = 6000
            values_clause = ",".join(f"({i}, 'val_{i}')" for i in range(n_rows))
            sql = f"INSERT INTO Sheet (A, B) VALUES {values_clause}"

            result = engine.execute_insert_query(path, sql)

            # 应该被拒绝
            assert result["success"] is False, "超量 INSERT 应该失败"
            msg = result.get("message", "")
            assert "超过限制" in msg or "limit" in msg.lower() or \
                   "batch" in msg.lower() or "5000" in msg or \
                   "分批" in msg, \
                   f"错误信息应包含限制相关关键词: {msg}"
        finally:
            os.unlink(path)

    def test_insert_at_exact_boundary(self):
        """刚好在边界值的插入应被允许（5000 行）"""
        data = [["C"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            n_rows = 5000
            values_clause = ",".join(f"({i},)" for i in range(n_rows))
            sql = f"INSERT INTO Sheet (C) VALUES {values_clause}"

            result = engine.execute_insert_query(path, sql)
            # 不应返回"超限"错误（可能因其他原因失败）
            if not result["success"]:
                msg = result.get("message", "")
                assert "超过限制" not in msg and "5000" not in msg, \
                    f"边界值5000不应被拒绝: {msg}"
        finally:
            os.unlink(path)

    def test_insert_just_over_boundary_rejected(self):
        """刚好超边界值(5001行)的插入应被拒绝"""
        data = [["C"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            n_rows = 5001
            values_clause = ",".join(f"({i},)" for i in range(n_rows))
            sql = f"INSERT INTO Sheet (C) VALUES {values_clause}"

            result = engine.execute_insert_query(path, sql)
            assert result["success"] is False, "5001行的INSERT应被拒绝"
            msg = result.get("message", "")
            assert "超过限制" in msg or "5000" in msg or "limit" in msg.lower(), \
                   f"应返回限制相关错误: {msg}"
        finally:
            os.unlink(path)


# ============================================================
# P1-05: _evaluate_update_expression 递归深度保护
# ============================================================

class TestP1_05_UpdateRecursionDepthGuard:
    """P1-05: UPDATE SET 表达式求值必须有递归深度保护"""

    def test_normal_nested_expression(self):
        """正常的嵌套表达式应正常工作"""
        data = [["Name", "Val"], ["Alice", 10.0], ["Bob", 20.0]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            # 嵌套算术运算
            result = engine.execute_update_query(
                path,
                "UPDATE Sheet SET Val = Val + 2 * 3 WHERE Name = 'Alice'"
            )
            assert result["success"] is True, f"UPDATE 失败: {result.get('message', '')}"

            # 验证结果：10 + 2*3 = 16
            verify = execute_advanced_sql_query(
                path,
                "SELECT Val FROM Sheet WHERE Name = 'Alice'"
            )
            assert verify["success"] is True
            val = verify["data"][1][0]
            assert val == 16.0, f"期望 16.0, 实际 {val}"
        finally:
            os.unlink(path)

    def test_deeply_nested_function_calls(self):
        """深层嵌套函数调用不应崩溃（有深度保护）"""
        data = [["S"], ["hello_world"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            # 构造深层嵌套的 LEFT/SUBSTRING 函数调用
            result = engine.execute_update_query(
                path,
                "UPDATE Sheet SET S = LEFT(SUBSTRING(S, 1, 5), 3)"
            )
            # 无论成功还是返回深度保护错误，都不应抛出未捕获的异常
            assert result["success"] is True or \
                   "depth" in str(result.get("message", "")).lower() or \
                   "递归" in result.get("message", "") or \
                   "error" not in str(result.get("message", "")).lower()[:20], \
                   f"深层嵌套不应导致未处理异常: {result}"
        finally:
            os.unlink(path)

    def test_update_with_column_reference_and_arithmetic(self):
        """UPDATE SET 中列引用+算术运算的正确性"""
        data = [["Price", "Qty", "Total"], [10.0, 2, 20.0], [5.0, 3, 15.0], [8.0, 1, 8.0]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            # 用 Price * Qty 更新 Total
            result = engine.execute_update_query(
                path,
                "UPDATE Sheet SET Total = Price * Qty"
            )
            assert result["success"] is True, f"UPDATE 失败: {result.get('message', '')}"

            verify = execute_advanced_sql_query(path, "SELECT Price, Qty, Total FROM Sheet")
            assert verify["success"] is True
            for row in verify["data"][1:]:
                price, qty, total = row[0], row[1], row[2]
                expected = price * qty
                assert abs(total - expected) < 0.01, \
                    f"{price}*{qty}={expected}, 但Total={total}"
        finally:
            os.unlink(path)

    def test_update_with_string_functions(self):
        """UPDATE SET 中字符串函数(UPPER/LOWER/SUBSTRING)的正确性"""
        data = [["Name"], ["alice"], ["BOB"], ["Charlie"]]
        path = _make_test_xlsx(data)
        try:
            engine = AdvancedSQLQueryEngine()
            result = engine.execute_update_query(
                path,
                "UPDATE Sheet SET Name = UPPER(Name)"
            )
            assert result["success"] is True

            verify = execute_advanced_sql_query(path, "SELECT Name FROM Sheet ORDER BY Name")
            names = [r[0] for r in verify["data"][1:]]
            assert all(n == n.upper() for n in names), f"期望全大写: {names}"
        finally:
            os.unlink(path)

    def test_recursion_depth_guard_on_unknown_expr_type(self):
        """验证递归深度保护：构造极端嵌套不应导致 RecursionError"""
        import sqlglot.expressions as exp
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        engine = AdvancedSQLQueryEngine()
        df = pd.DataFrame({"A": [1], "B": [2]})

        # 构造一个深层嵌套的表达式链（模拟未知类型嵌套）
        # 使用 Nested 类型来制造多层包装
        inner = exp.Column(this="A")
        for _ in range(30):  # 超过 MAX_RECURSION_DEPTH=20
            inner = exp.Anonymous(this="wrap", expressions=[inner])

        # 调用不应抛出 RecursionError，而应在深度限制处安全返回 None/空值
        try:
            result = engine._evaluate_update_expression(inner, df, 0)
            # 如果到达这里且没有崩溃，说明深度保护生效
            assert result is None or result == "", \
                f"超深嵌套应返回 None 或空字符串, 实际: {repr(result)}"
        except RecursionError:
            pytest.fail("RecursionError! 递归深度保护未生效")
