"""L3 高级不变量测试（INV-19, INV-21）。

INV-19: 写操作 SQLite 对齐 — UPDATE/INSERT/DELETE 后 ExcelMCP 和 SQLite 数据对齐
INV-21: 跨文件 JOIN 真值 — 跨文件 JOIN 结果与 SQLite 真值对齐
"""

from __future__ import annotations

import math
import sqlite3

import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_delete_query,
    execute_advanced_insert_query,
    execute_advanced_sql_query,
    execute_advanced_update_query,
)
from excel_mcp_server_fastmcp.calibrator.core import cmd_import, cmd_query, get_db_path


# ============================================================
# 辅助函数
# ============================================================

_CAL_DB_WRITE = "inv_round2_write"
_CAL_DB_JOIN = "inv_round2_join"


def _sqlite_execute_sql(sql: str, db_path: str) -> dict:
    """直接用 sqlite3 执行 SQL（包含 commit），用于 DML 操作"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        conn.commit()  # 关键：显式 commit
        rows = [tuple(row) for row in cursor.fetchall()]
        headers = [desc[0] for desc in cursor.description] if cursor.description else []
        return {"success": True, "headers": headers, "rows": rows}
    except sqlite3.Error as e:
        conn.rollback()
        return {"success": False, "headers": [], "rows": [], "message": str(e)}
    finally:
        conn.close()


def _align_result(excel_result: dict, sqlite_result: dict, tol: float = 0.01) -> bool:
    """比较 ExcelMCP 和 SQLite 结果是否一致。

    - 跳过 SQLite 的 _rowid_ 列（calibrator 自动添加）
    - 浮点容差 tol
    - 空值/None 统一处理
    - 处理行顺序差异（排序后比较）
    """
    if not excel_result["success"] or not sqlite_result.get("success"):
        return False

    excel_data = excel_result["data"]
    sqlite_rows = sqlite_result.get("rows", [])
    sqlite_headers = sqlite_result.get("headers", [])

    if len(excel_data) == 0 and len(sqlite_rows) == 0:
        return True
    if len(excel_data) == 0 or len(sqlite_rows) == 0:
        return False

    # 跳过 _rowid_ 列
    rowid_idx = None
    for idx, h in enumerate(sqlite_headers):
        if h == "_rowid_":
            rowid_idx = idx
            break

    # 构建不含 _rowid_ 的 SQLite 数据
    sqlite_rows_clean = []
    for row in sqlite_rows:
        clean_row = [v for i, v in enumerate(row) if i != rowid_idx]
        sqlite_rows_clean.append(clean_row)

    # 行数一致（Excel 数据含表头，SQLite 数据不含表头）
    if len(excel_data) - 1 != len(sqlite_rows_clean):
        return False

    def _value_key(v):
        """用于排序的值归一化：None < number < string"""
        if v is None:
            return (0, 0)
        try:
            f = float(v)
            if math.isnan(f):
                return (0, 0)
            return (1, float(f))
        except (ValueError, TypeError):
            return (2, str(v))

    def _row_key(row):
        return tuple(_value_key(v) for v in row)

    # 提取数据行并排序
    excel_rows = excel_data[1:]
    sorted_excel = sorted(excel_rows, key=_row_key)
    sorted_sqlite = sorted(sqlite_rows_clean, key=_row_key)

    # 逐值比较
    for erow, srow in zip(sorted_excel, sorted_sqlite):
        if len(erow) != len(srow):
            return False
        for ev, sv in zip(erow, srow):
            if ev is None and sv is None:
                continue
            if ev is None or sv is None:
                ev_str = str(ev).strip() if ev is not None else ""
                sv_str = str(sv).strip() if sv is not None else ""
                if ev_str == "" and sv_str == "":
                    continue
                return False
            try:
                ef = float(ev)
                sf = float(sv)
                if abs(ef - sf) > tol:
                    return False
            except (ValueError, TypeError):
                if str(ev).strip() != str(sv).strip():
                    return False
    return True


# ============================================================
# INV-19: 写操作 SQLite 对齐
# ============================================================


class TestINV19WriteSQLiteAlignment:
    """INV-19: UPDATE/INSERT/DELETE 后 ExcelMCP 和 SQLite 数据对齐"""

    @pytest.fixture(autouse=True)
    def _setup(self, writable_file):
        self.file_path = writable_file
        import_result = cmd_import(writable_file, _CAL_DB_WRITE)
        assert import_result["success"], f"calibrator 导入失败: {import_result}"
        self.db_path = get_db_path(_CAL_DB_WRITE)

    def test_update_then_select_alignment(self):
        """UPDATE 后 ExcelMCP 和 SQLite 的 SELECT 结果对齐"""
        execute_advanced_update_query(self.file_path, "UPDATE 商品 SET Price = 0 WHERE Active = '是'")
        _sqlite_execute_sql("UPDATE 商品 SET Price = 0 WHERE Active = '是'", self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT ID, Name, Price, Stock, Active FROM 商品 ORDER BY ID")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT ID, Name, Price, Stock, Active FROM 商品 ORDER BY ID")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_insert_then_select_alignment(self):
        """INSERT 后 ExcelMCP 和 SQLite 的 SELECT 结果对齐"""
        execute_advanced_insert_query(self.file_path,
            "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (99, '测试', 123.45, 10, '是')")
        _sqlite_execute_sql(
            "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (99, '测试', 123.45, 10, '是')",
            self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT * FROM 商品 WHERE ID = 99")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT * FROM 商品 WHERE ID = 99")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_delete_then_select_alignment(self):
        """DELETE 后 ExcelMCP 和 SQLite 的 SELECT 结果对齐"""
        execute_advanced_delete_query(self.file_path, "DELETE FROM 商品 WHERE Active = '否'")
        _sqlite_execute_sql("DELETE FROM 商品 WHERE Active = '否'", self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT ID, Name FROM 商品 ORDER BY ID")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT ID, Name FROM 商品 ORDER BY ID")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_update_expression_alignment(self):
        """UPDATE SET 表达式后对齐"""
        execute_advanced_update_query(self.file_path, "UPDATE 商品 SET Price = ROUND(Price * 1.1, 2)")
        _sqlite_execute_sql("UPDATE 商品 SET Price = ROUND(Price * 1.1, 2)", self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT ID, Price FROM 商品 ORDER BY ID")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT ID, Price FROM 商品 ORDER BY ID")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_multiple_writes_alignment(self):
        """连续多次写操作后对齐"""
        # UPDATE
        execute_advanced_update_query(self.file_path, "UPDATE 商品 SET Price = 0 WHERE ID = 1")
        _sqlite_execute_sql("UPDATE 商品 SET Price = 0 WHERE ID = 1", self.db_path)
        # INSERT
        execute_advanced_insert_query(self.file_path,
            "INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (90, 'X', 1, 1, '是')")
        _sqlite_execute_sql("INSERT INTO 商品 (ID, Name, Price, Stock, Active) VALUES (90, 'X', 1, 1, '是')", self.db_path)
        # DELETE
        execute_advanced_delete_query(self.file_path, "DELETE FROM 商品 WHERE ID = 6")
        _sqlite_execute_sql("DELETE FROM 商品 WHERE ID = 6", self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT * FROM 商品 ORDER BY ID")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT * FROM 商品 ORDER BY ID")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_null_write_alignment(self):
        """NULL 写入后对齐"""
        execute_advanced_update_query(self.file_path, "UPDATE 商品 SET Price = NULL WHERE ID = 1")
        _sqlite_execute_sql("UPDATE 商品 SET Price = NULL WHERE ID = 1", self.db_path)

        excel_result = execute_advanced_sql_query(self.file_path, "SELECT ID, Price FROM 商品 WHERE ID = 1")
        sqlite_result = cmd_query(_CAL_DB_WRITE, "SELECT ID, Price FROM 商品 WHERE ID = 1")

        assert _align_result(excel_result, sqlite_result), (
            f"ExcelMCP 和 SQLite 结果不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )


# ============================================================
# INV-21: 跨文件 JOIN 真值
# ============================================================


class TestINV21CrossFileJoin:
    """INV-21: 跨文件 JOIN 结果与 SQLite 真值对齐"""

    @pytest.fixture(autouse=True)
    def _setup(self, skills_file, drops_file):
        self.skills_file = skills_file
        self.drops_file = drops_file
        # 导入两个文件到同一 SQLite 数据库
        r1 = cmd_import(skills_file, _CAL_DB_JOIN)
        r2 = cmd_import(drops_file, _CAL_DB_JOIN)
        assert r1["success"], f"技能表导入失败: {r1}"
        assert r2["success"], f"掉落表导入失败: {r2}"
        self.db_path = get_db_path(_CAL_DB_JOIN)

    def _query_sqlite(self, sql):
        """查询 SQLite（只读）"""
        return cmd_query(_CAL_DB_JOIN, sql)

    def test_inner_join_alignment(self):
        """INNER JOIN 结果与 SQLite 对齐"""
        sql = f"""SELECT s.技能名称, d.掉落物品, d.数量
                  FROM 技能配置@'{self.skills_file}' s
                  JOIN 掉落配置@'{self.drops_file}' d ON s.技能ID = d.关联技能"""
        excel_result = execute_advanced_sql_query(self.skills_file, sql)

        sqlite_sql = """SELECT s.技能名称, d.掉落物品, d.数量
                        FROM 技能配置 s JOIN 掉落配置 d ON s.技能ID = d.关联技能"""
        sqlite_result = self._query_sqlite(sqlite_sql)

        assert _align_result(excel_result, sqlite_result), (
            f"INNER JOIN 不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_left_join_alignment(self):
        """LEFT JOIN 结果与 SQLite 对齐"""
        sql = f"""SELECT s.技能名称, d.掉落物品
                  FROM 技能配置@'{self.skills_file}' s
                  LEFT JOIN 掉落配置@'{self.drops_file}' d ON s.技能ID = d.关联技能"""
        excel_result = execute_advanced_sql_query(self.skills_file, sql)

        sqlite_sql = """SELECT s.技能名称, d.掉落物品
                        FROM 技能配置 s LEFT JOIN 掉落配置 d ON s.技能ID = d.关联技能"""
        sqlite_result = self._query_sqlite(sqlite_sql)

        assert _align_result(excel_result, sqlite_result), (
            f"LEFT JOIN 不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_join_with_where_alignment(self):
        """JOIN + WHERE 结果与 SQLite 对齐"""
        sql = f"""SELECT s.技能名称, d.掉落物品, d.数量
                  FROM 技能配置@'{self.skills_file}' s
                  JOIN 掉落配置@'{self.drops_file}' d ON s.技能ID = d.关联技能
                  WHERE d.数量 > 2"""
        excel_result = execute_advanced_sql_query(self.skills_file, sql)

        sqlite_sql = """SELECT s.技能名称, d.掉落物品, d.数量
                        FROM 技能配置 s JOIN 掉落配置 d ON s.技能ID = d.关联技能
                        WHERE d.数量 > 2"""
        sqlite_result = self._query_sqlite(sqlite_sql)

        assert _align_result(excel_result, sqlite_result), (
            f"JOIN + WHERE 不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_join_with_aggregation_alignment(self):
        """JOIN + 聚合结果与 SQLite 对齐"""
        sql = f"""SELECT s.类型, COUNT(*) as cnt, SUM(d.数量) as total
                  FROM 技能配置@'{self.skills_file}' s
                  JOIN 掉落配置@'{self.drops_file}' d ON s.技能ID = d.关联技能
                  GROUP BY s.类型"""
        excel_result = execute_advanced_sql_query(self.skills_file, sql)

        sqlite_sql = """SELECT s.类型, COUNT(*) as cnt, SUM(d.数量) as total
                        FROM 技能配置 s JOIN 掉落配置 d ON s.技能ID = d.关联技能
                        GROUP BY s.类型"""
        sqlite_result = self._query_sqlite(sqlite_sql)

        assert _align_result(excel_result, sqlite_result), (
            f"JOIN + 聚合不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )

    def test_join_count_consistency(self):
        """JOIN 结果行数与 SQLite 一致"""
        sql = f"""SELECT COUNT(*) FROM 技能配置@'{self.skills_file}' s
                  JOIN 掉落配置@'{self.drops_file}' d ON s.技能ID = d.关联技能"""
        excel_result = execute_advanced_sql_query(self.skills_file, sql)

        sqlite_sql = """SELECT COUNT(*) FROM 技能配置 s JOIN 掉落配置 d ON s.技能ID = d.关联技能"""
        sqlite_result = self._query_sqlite(sqlite_sql)

        assert _align_result(excel_result, sqlite_result), (
            f"JOIN COUNT 不一致\n"
            f"Excel: {excel_result['data']}\n"
            f"SQLite: {sqlite_result.get('rows', [])}"
        )
