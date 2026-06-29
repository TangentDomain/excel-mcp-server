#!/usr/bin/env python3
"""SQL 准确率差分测试基准（autoresearch primary metric）

对 ExcelMCP SQL 引擎与 SQLite oracle 跑同一批 SQL，统计对齐率。
确定性：固定 fixture 数据、固定 SQL 集、固定种子，无网络/时间依赖。

用法:
    python tools/sql-accuracy-benchmark.py            # 完整跑
    python tools/sql-accuracy-benchmark.py --quick    # 快速子集（调试用）

输出:
    METRIC accuracy=<百分比>            # 主指标
    METRIC total_cases=<总数>           # 次指标
    METRIC passed=<通过数>              # 次指标
    METRIC failed=<失败数>              # 次指标
    METRIC duration_ms=<耗时>           # 次指标
"""

from __future__ import annotations

import math
import os
import sys
import tempfile
import time
from pathlib import Path

# ── 确保 src 在 path ──
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from openpyxl import Workbook  # noqa: E402

from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query  # noqa: E402
from excel_mcp_server_fastmcp.calibrator.core import cmd_import, cmd_query  # noqa: E402


# ============================================================
# Fixture 生成（自包含，不依赖 pytest fixtures）
# ============================================================


def _make_simple_wb() -> Workbook:
    """简单测试表：ID(int), Name(str), Price(float), Active(str), Tags(str)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws.append(["ID", "Name", "Price", "Active", "Tags"])
    rows = [
        [1, "铁剑", 50.0, "是", "武器"],
        [2, "法杖", 80.0, "是", "武器"],
        [3, "皮甲", 30.0, "否", "防具"],
        [4, "锁子甲", None, "是", "防具"],
        [5, "传说之刃", 200.0, "是", "武器"],
    ]
    for r in rows:
        ws.append(r)
    return wb


def _make_numbers_wb() -> Workbook:
    """数值表：用于聚合/窗口函数测试（含并列值）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "数值"
    ws.append(["id", "grp", "val"])
    # grp=A 的 val: 10,20,20,30  (20 有并列，测试 RANK vs DENSE_RANK)
    # grp=B 的 val: 5,15
    rows = [
        [1, "A", 10],
        [2, "A", 20],
        [3, "A", 20],
        [4, "A", 30],
        [5, "B", 5],
        [6, "B", 15],
    ]
    for r in rows:
        ws.append(r)
    return wb


def _make_dual_header_wb() -> Workbook:
    """双行表头表：游戏配置格式（第1行中文描述 + 第2行英文字段名）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "技能配置"
    ws.append(["技能ID", "技能名称", "伤害", "类型"])
    ws.append(["skill_id", "skill_name", "damage", "skill_type"])
    rows = [
        ["SK001", "火球术", 100, "法师"],
        ["SK002", "冰冻术", 80, "法师"],
        ["SK003", "斩击", 120, "战士"],
        ["SK004", "火球术强化", 200, "法师"],
        ["SK005", "冰风暴", 180, "法师"],
        ["SK006", "旋风斩", 150, "战士"],
        ["SK007", "暗影突袭", 90, "刺客"],
        ["SK008", "致命一击", 250, "刺客"],
    ]
    for r in rows:
        ws.append(r)
    return wb


def _make_special_char_wb() -> Workbook:
    """特殊字符表：中文/emoji/单引号/空值"""
    wb = Workbook()
    ws = wb.active
    ws.title = "特殊字符"
    ws.append(["ID", "文本", "备注"])
    rows = [
        [1, "你好世界", "普通中文"],
        [2, "🎮游戏", "emoji"],
        [3, "it's a test", "单引号"],
        [4, "path\\to\\file", "反斜杠"],
        [5, None, "空值"],
        [6, "日本語テスト", "日文"],
    ]
    for r in rows:
        ws.append(r)
    return wb


def _make_join_wb() -> Workbook:
    """JOIN 测试表：技能表 + 掉落表（同文件不同 sheet）"""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "技能"
    ws1.append(["skill_id", "skill_name", "damage"])
    for r in [
        ["S1", "火球术", 100],
        ["S2", "冰冻术", 80],
        ["S3", "斩击", 120],
        ["S4", "治疗术", 0],
    ]:
        ws1.append(r)

    ws2 = wb.create_sheet("掉落")
    ws2.append(["drop_id", "skill_ref", "item_name", "qty"])
    for r in [
        ["D1", "S1", "火焰水晶", 2],
        ["D2", "S1", "魔法粉尘", 5],
        ["D3", "S2", "冰晶", 3],
        ["D4", "S3", "铁矿石", 1],
    ]:
        ws2.append(r)
    return wb


def _save(wb: Workbook, directory: Path, name: str) -> str:
    p = directory / name
    wb.save(str(p))
    return str(p)


# ============================================================
# 对齐比较（从 test_l1_result_structure._align_result 精简）
# ============================================================


def align_result(excel_result: dict, sqlite_result: dict, tol: float = 0.01) -> bool:
    """比较 ExcelMCP 和 SQLite 结果是否一致。"""
    if not excel_result.get("success") or not sqlite_result.get("success"):
        return False

    excel_data = excel_result["data"]
    sqlite_rows = sqlite_result.get("rows", [])
    sqlite_headers = sqlite_result.get("headers", [])

    # Excel data 含表头行（data[0]）；SQLite rows 不含表头
    excel_rows = excel_data[1:] if excel_data else []
    if len(excel_rows) == 0 and len(sqlite_rows) == 0:
        return True
    if len(excel_rows) == 0 or len(sqlite_rows) == 0:
        return False

    # 跳过 _rowid_
    rowid_idx = None
    for idx, h in enumerate(sqlite_headers):
        if h == "_rowid_":
            rowid_idx = idx
            break
    sqlite_clean = [[v for i, v in enumerate(row) if i != rowid_idx] for row in sqlite_rows]

    if len(excel_rows) != len(sqlite_clean):
        return False

    def _vk(v):
        if v is None:
            return (0, 0)
        try:
            f = float(v)
            return (0, f) if not math.isnan(f) else (0, 0)
        except (ValueError, TypeError):
            return (1, str(v))

    se = sorted(excel_rows, key=lambda r: tuple(_vk(v) for v in r))
    ss = sorted(sqlite_clean, key=lambda r: tuple(_vk(v) for v in r))

    for erow, srow in zip(se, ss):
        if len(erow) != len(srow):
            return False
        for ev, sv in zip(erow, srow):
            if ev is None and sv is None:
                continue
            if ev is None or sv is None:
                ev_s = str(ev).strip() if ev is not None else ""
                sv_s = str(sv).strip() if sv is not None else ""
                if ev_s == "" and sv_s == "":
                    continue
                return False
            try:
                if abs(float(ev) - float(sv)) > tol:
                    return False
            except (ValueError, TypeError):
                if str(ev).strip() != str(sv).strip():
                    return False
    return True


# ============================================================
# SQL 测试集（确定性，按类别）
# ============================================================


def build_test_cases() -> list[dict]:
    """构建确定性 SQL 差分测试集。

    每个测试: {file_key, sql, category, note}
    file_key 对应 FIXTURE_BUILDER 中的键。
    """
    cases: list[dict] = []

    # ── 基础 SELECT ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据", "cat": "select"})
    cases.append({"f": "simple", "sql": "SELECT ID, Name FROM 数据", "cat": "select"})
    cases.append({"f": "simple", "sql": "SELECT DISTINCT Active FROM 数据", "cat": "distinct"})
    cases.append({"f": "simple", "sql": "SELECT DISTINCT Tags FROM 数据", "cat": "distinct"})

    # ── WHERE 条件 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE ID = 1", "cat": "where"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price > 50", "cat": "where"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE ID IN (1,3,5)", "cat": "where_in"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price BETWEEN 50 AND 200", "cat": "where_between"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Name LIKE '%剑%'", "cat": "where_like"})

    # ── NULL 处理 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price IS NULL", "cat": "null"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price IS NOT NULL", "cat": "null"})

    # ── ORDER BY / LIMIT / OFFSET ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 ORDER BY Price DESC", "cat": "orderby"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 ORDER BY Price ASC", "cat": "orderby"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 LIMIT 2", "cat": "limit"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 LIMIT 2 OFFSET 1", "cat": "offset"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 ORDER BY ID DESC LIMIT 3", "cat": "orderby_limit"})

    # ── 聚合函数 ──
    cases.append({"f": "simple", "sql": "SELECT COUNT(*) FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT COUNT(Price) FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT SUM(Price) FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT AVG(Price) FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT MAX(Price), MIN(Price) FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT COUNT(DISTINCT Active) FROM 数据", "cat": "agg_distinct"})

    # ── GROUP BY / HAVING ──
    cases.append({"f": "simple", "sql": "SELECT Active, COUNT(*) FROM 数据 GROUP BY Active", "cat": "groupby"})
    cases.append({"f": "simple", "sql": "SELECT Tags, COUNT(*), AVG(Price) FROM 数据 GROUP BY Tags", "cat": "groupby"})
    cases.append({"f": "numbers", "sql": "SELECT grp, COUNT(*) FROM 数值 GROUP BY grp", "cat": "groupby"})
    cases.append({"f": "numbers", "sql": "SELECT grp, SUM(val) FROM 数值 GROUP BY grp HAVING SUM(val) > 20", "cat": "having"})

    # ── 数学表达式 ──
    cases.append({"f": "simple", "sql": "SELECT ID, Price * 2 AS doubled FROM 数据", "cat": "expr"})
    cases.append({"f": "numbers", "sql": "SELECT id, val + 10 AS shifted FROM 数值", "cat": "expr"})

    # ── CASE WHEN ──
    cases.append({"f": "simple", "sql": "SELECT Name, CASE WHEN Price > 100 THEN '贵' ELSE '便宜' END AS 等级 FROM 数据", "cat": "case_when"})

    # ── 字符串函数 ──
    cases.append({"f": "simple", "sql": "SELECT Name, LENGTH(Name) AS len FROM 数据", "cat": "string"})
    cases.append({"f": "simple", "sql": "SELECT UPPER(Active) AS ua FROM 数据", "cat": "string"})

    # ── 子查询 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE ID IN (SELECT ID FROM 数据 WHERE Active = '是')", "cat": "subquery"})
    cases.append({"f": "numbers", "sql": "SELECT * FROM 数值 WHERE val > (SELECT AVG(val) FROM 数值)", "cat": "subquery"})

    # ── 窗口函数 ──
    cases.append({"f": "numbers", "sql": "SELECT id, val, ROW_NUMBER() OVER (PARTITION BY grp ORDER BY val DESC) AS rn FROM 数值", "cat": "window"})
    cases.append({"f": "numbers", "sql": "SELECT id, val, RANK() OVER (PARTITION BY grp ORDER BY val DESC) AS rnk FROM 数值", "cat": "window"})
    cases.append({"f": "numbers", "sql": "SELECT id, val, DENSE_RANK() OVER (PARTITION BY grp ORDER BY val DESC) AS drnk FROM 数值", "cat": "window"})

    # ── 双行表头 ──
    cases.append({"f": "dual", "sql": "SELECT * FROM 技能配置", "cat": "dual_header"})
    # 双行表头: calibrator 把列名拍平成 "中文描述.英文名"（如 技能ID.skill_id），
    # 引擎则用英文名。两侧列名体系不同，只有 SELECT * 能在两侧都成立。
    # 因此双行表头类别只测裸 * 查询（按列位置比较数据值）。
    cases.append({"f": "dual", "sql": "SELECT * FROM 技能配置 LIMIT 5", "cat": "dual_header"})
    cases.append({"f": "dual", "sql": "SELECT COUNT(*) FROM 技能配置", "cat": "dual_header"})

    # ── 特殊字符 ──
    cases.append({"f": "special", "sql": "SELECT * FROM 特殊字符", "cat": "special_char"})
    cases.append({"f": "special", "sql": "SELECT * FROM 特殊字符 WHERE ID > 3", "cat": "special_char"})

    # ── JOIN（同文件多 sheet）──
    cases.append({"f": "join", "sql": "SELECT 技能.skill_id, 技能.skill_name, 掉落.item_name FROM 技能 JOIN 掉落 ON 技能.skill_id = 掉落.skill_ref", "cat": "join"})
    cases.append({"f": "join", "sql": "SELECT 技能.skill_name, 掉落.item_name, 掉落.qty FROM 技能 JOIN 掉落 ON 技能.skill_id = 掉落.skill_ref WHERE 掉落.qty >= 3", "cat": "join"})

    # ── 扩展: NOT IN / NOT LIKE ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE ID NOT IN (1,3)", "cat": "not_in"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Name NOT LIKE '%剑%'", "cat": "not_like"})

    # ── 扩展: 多列排序 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 ORDER BY Tags ASC, Price DESC", "cat": "orderby_multi"})
    cases.append({"f": "numbers", "sql": "SELECT * FROM 数值 ORDER BY grp ASC, val DESC", "cat": "orderby_multi"})

    # ── 扩展: 复杂条件 (AND/OR/NOT) ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price > 50 AND Active = '是'", "cat": "compound_where"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Tags = '武器' OR Tags = '防具'", "cat": "compound_where"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE NOT (Active = '是')", "cat": "compound_where"})

    # ── 扩展: 数学表达式 / 除法 ──
    cases.append({"f": "numbers", "sql": "SELECT id, val * 2 AS double_val FROM 数值", "cat": "expr"})
    cases.append({"f": "numbers", "sql": "SELECT id, val - 5 AS diff FROM 数值", "cat": "expr"})
    cases.append({"f": "numbers", "sql": "SELECT id, val / 2 AS half FROM 数值", "cat": "div"})
    cases.append({"f": "numbers", "sql": "SELECT id, val % 3 AS mod_val FROM 数值", "cat": "expr"})
    cases.append({"f": "numbers", "sql": "SELECT grp, SUM(val * 2) FROM 数值 GROUP BY grp", "cat": "expr_agg"})

    # ── 扩展: 字符串函数 ──
    cases.append({"f": "simple", "sql": "SELECT Name, LOWER(Name) AS lname FROM 数据", "cat": "string"})
    cases.append({"f": "simple", "sql": "SELECT Name, LENGTH(Tags) AS tlen FROM 数据", "cat": "string"})

    # ── 扩展: CASE WHEN 多分支 ──
    cases.append({"f": "numbers", "sql": "SELECT id, CASE WHEN val >= 20 THEN '高' WHEN val >= 10 THEN '中' ELSE '低' END AS grade FROM 数值", "cat": "case_when"})
    cases.append({"f": "simple", "sql": "SELECT Name, CASE WHEN Price IS NULL THEN '未知' ELSE '已知' END AS price_status FROM 数据", "cat": "case_null"})

    # ── 扩展: 聚合边界 ──
    cases.append({"f": "simple", "sql": "SELECT MAX(Price) - MIN(Price) AS range_val FROM 数据", "cat": "agg"})
    cases.append({"f": "simple", "sql": "SELECT SUM(Price) / COUNT(Price) AS manual_avg FROM 数据", "cat": "agg_expr"})
    cases.append({"f": "numbers", "sql": "SELECT grp, MIN(val), MAX(val), COUNT(*) FROM 数值 GROUP BY grp", "cat": "groupby_multi_agg"})

    # ── 扩展: 窗口函数更多场景 ──
    cases.append({"f": "numbers", "sql": "SELECT id, grp, val, SUM(val) OVER (PARTITION BY grp) AS grp_total FROM 数值", "cat": "window"})
    cases.append({"f": "numbers", "sql": "SELECT id, val, ROW_NUMBER() OVER (ORDER BY val) AS rn_all FROM 数值", "cat": "window"})

    # ── 扩展: 嵌套子查询 / EXISTS ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Price > (SELECT AVG(Price) FROM 数据)", "cat": "subquery"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE EXISTS (SELECT 1 FROM 数据 d2 WHERE d2.ID = 数据.ID AND d2.Price > 100)", "cat": "exists"})

    # ── 扩展: 空结果集 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE ID = 999", "cat": "empty_result"})
    cases.append({"f": "simple", "sql": "SELECT COUNT(*) FROM 数据 WHERE Price > 99999", "cat": "empty_result"})

    # ── 扩展: LIKE 变体 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Name LIKE '铁%'", "cat": "where_like"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 WHERE Tags LIKE '%具'", "cat": "where_like"})

    # ── 扩展: OFFSET 边界 ──
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 LIMIT 10 OFFSET 3", "cat": "offset"})
    cases.append({"f": "simple", "sql": "SELECT * FROM 数据 ORDER BY ID LIMIT 0", "cat": "limit_zero"})

    # ── 扩展: DISTINCT 多列 ──
    cases.append({"f": "simple", "sql": "SELECT DISTINCT Tags, Active FROM 数据", "cat": "distinct"})

    return cases


FIXTURE_BUILDERS = {
    "simple": _make_simple_wb,
    "numbers": _make_numbers_wb,
    "dual": _make_dual_header_wb,
    "special": _make_special_char_wb,
    "join": _make_join_wb,
}


# ============================================================
# 主流程
# ============================================================


def main() -> int:
    quick = "--quick" in sys.argv

    with tempfile.TemporaryDirectory(prefix="sqlbench_", ignore_cleanup_errors=True) as tmpdir:
        tmpdir_path = Path(tmpdir)

        # 1. 生成 fixture 文件
        fixtures: dict[str, str] = {}
        for key, builder in FIXTURE_BUILDERS.items():
            fname = f"{key}.xlsx"
            fixtures[key] = _save(builder(), tmpdir_path, fname)

        # 2. 导入 SQLite（每类一个 db，避免表名冲突）
        cal_dbs: dict[str, str] = {}
        for key, fpath in fixtures.items():
            db_name = f"bench_{key}"
            res = cmd_import(fpath, db_name)
            if not res.get("success"):
                print(f"WARNING: calibrator 导入失败 {key}: {res.get('message', '')}", file=sys.stderr)
            cal_dbs[key] = db_name

        # 3. 构建 + 过滤测试集
        all_cases = build_test_cases()
        if quick:
            all_cases = all_cases[:10]

        # 4. 跑差分测试
        total = len(all_cases)
        passed = 0
        failures: list[tuple[str, str, str]] = []
        cat_stats: dict[str, list[int]] = {}  # cat -> [passed, total]

        t0 = time.perf_counter()
        for i, tc in enumerate(all_cases):
            fkey = tc["f"]
            sql = tc["sql"]
            cat = tc["cat"]

            fpath = fixtures.get(fkey, "")
            db = cal_dbs.get(fkey, "")

            # 跑 ExcelMCP
            try:
                excel_res = execute_advanced_sql_query(fpath, sql)
            except Exception as e:
                excel_res = {"success": False, "data": [], "message": str(e)}

            # 跑 SQLite oracle
            try:
                sqlite_res = cmd_query(db, sql)
            except Exception as e:
                sqlite_res = {"success": False, "rows": [], "headers": [], "message": str(e)}

            ok = align_result(excel_res, sqlite_res)
            cat_stats.setdefault(cat, [0, 0])
            cat_stats[cat][1] += 1
            if ok:
                passed += 1
                cat_stats[cat][0] += 1
            else:
                emsg = excel_res.get("message", "")
                smsg = sqlite_res.get("message", "")
                note = f"excel={emsg[:60]}|sqlite={smsg[:60]}"
                failures.append((cat, sql[:80], note))

        elapsed_ms = (time.perf_counter() - t0) * 1000

        # 5. 计算指标
        accuracy = (passed / total * 100.0) if total > 0 else 0.0
        failed = total - passed

        # 6. 输出 METRIC 行
        print(f"METRIC accuracy={accuracy:.2f}")
        print(f"METRIC total_cases={total}")
        print(f"METRIC passed={passed}")
        print(f"METRIC failed={failed}")
        print(f"METRIC duration_ms={elapsed_ms:.1f}")

        # 7. 分类统计到 stderr（不影响 METRIC 解析）
        print("\n── 分类统计 ──", file=sys.stderr)
        for cat in sorted(cat_stats.keys()):
            p, t = cat_stats[cat]
            pct = (p / t * 100.0) if t > 0 else 0.0
            print(f"  {cat:20s} {p}/{t} ({pct:.0f}%)", file=sys.stderr)

        if failures:
            print("\n── 失败用例 ──", file=sys.stderr)
            for cat, sql, note in failures[:20]:
                print(f"  [{cat}] {sql}", file=sys.stderr)
                print(f"         {note}", file=sys.stderr)

        # 退出码：accuracy < 50% 视为异常（但仍输出指标供记录）
        return 0


if __name__ == "__main__":
    sys.exit(main())
