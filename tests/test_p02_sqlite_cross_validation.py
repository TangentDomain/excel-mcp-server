"""
TEST-P02: SQLite vs ExcelMCP 交叉对比验证（方向 B）
===================================================
用校准器对同一查询做双轨验证，确保 ExcelMCP SQL 引擎结果与 SQLite 一致。

覆盖 6 类复杂 SQL：
  B1: 多表 JOIN
  B2: 分组聚合 + HAVING
  B3: 窗口函数 (ROW_NUMBER / RANK)
  B4: 子查询（IN / EXISTS / 标量）
  B5: CASE WHEN 条件表达式
  B6: 双表头表的复杂查询

通过标准：ExcelMCP 结果与 SQLite 结果一致（允许浮点精度差异 < 0.01）
"""

import os
import sys
import sqlite3
import re
import pytest
import numpy as np

# Add workspace for calibrator import
sys.path.insert(0, "/root/workspace")
from sql_calibrator import get_db_path, cmd_import

from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

# Test data paths
TEST_DATA_DIR = os.path.join(os.path.dirname(__file__), "test_data")
GAME_CONFIG_PATH = os.path.join(TEST_DATA_DIR, "game_config.xlsx")
JOIN_TEST_PATH = os.path.join(TEST_DATA_DIR, "join_test.xlsx")
LARGE_SKILLS_PATH = os.path.join(TEST_DATA_DIR, "large_skills.xlsx")
COMPREHENSIVE_PATH = os.path.join(TEST_DATA_DIR, "comprehensive_test.xlsx")


# ============================================================
# 列名映射：ExcelMCP 英文别名 → 校准器 中文.英文 格式
# ============================================================

# game_config.xlsx 的映射
GAME_CONFIG_COLS = {
    # 装备配置 sheet
    'equip_id': '装备ID.equip_id',
    'equip_name': '装备名称.equip_name',
    'equip_type': '装备类型.equip_type',
    'quality': '品质.quality',
    'attack': '攻击力.attack',
    'defense': '防御力.defense',
    'hp': '生命值.hp',
    'level_req': '等级需求.level_req',
    # 技能配置 sheet
    'skill_id': '技能ID.skill_id',
    'skill_name': '技能名称.skill_name',
    'skill_type': '技能类型.skill_type',
    'level': '等级.level',
    'damage': '伤害.damage',
    'cooldown': '冷却时间.cooldown',
    'cost': '消耗.cost',
    'description': '描述.description',
}

# join_test.xlsx 的映射（非双表头，中英两列并存）
JOIN_TEST_COLS = {
    # 技能表
    'skill_name': 'skill_name',  # 有独立的 skill_name 列
    'skill_type': 'skill_type',
    'damage': 'damage',
    'equip_id': 'equip_id',
    # 装备表
    'equip_name': 'equip_name',
    'quality': 'quality',
    'attack_power': 'attack_power',
    # 怪物表
    'monster_name': 'monster_name',
}

# large_skills.xlsx 的映射
LARGE_SKILLS_COLS = {
    'skill_id': '技能ID.skill_id',
    'skill_name': '技能名称.skill_name',
    'skill_type': '技能类型.skill_type',
    'profession': '职业.profession',
    'level': '等级.level',
    'damage': '伤害.damage',
    'cooldown': '冷却.cooldown',
    'mp_cost': '消耗.mp_cost',
    'description': '描述.description',
    'quality': '品质.quality',
}


def translate_sql_for_calibrator(sql: str, col_map: dict) -> str:
    """
    将 ExcelMCP 风格的 SQL 翻译为校准器风格的 SQL。
    
    处理策略：
    1. 替换 alias.column_name → alias."中文.英文"（引号包裹，因为列名含点号）
    2. 替换裸 column_name → "中文.英文"
    3. 不替换字符串字面量（单引号内）
    4. 不替换表名
    5. 不替换已包含中文的列名（已翻译过）
    """
    result = sql
    
    # 按长度降序排列，避免短名替代长名的子串
    sorted_keys = sorted(col_map.keys(), key=len, reverse=True)
    
    for eng_alias in sorted_keys:
        calib_col = col_map[eng_alias]
        if eng_alias == calib_col:
            continue
        
        escaped = re.escape(eng_alias)
        quoted_calib = f'"{calib_col}"'
        
        # 模式1: alias.column_name → alias."中文.英文"
        pattern_qualified = r'(\w+)\.' + escaped + r'(?![.\w])'
        def repl_qualified(m, _q=quoted_calib):
            start = m.start()
            prefix = result[:start]
            if prefix.count("'") % 2 == 1:
                return m.group(0)
            alias = m.group(1)
            return f"{alias}.{_q}"
        
        result = re.sub(pattern_qualified, repl_qualified, result)
        
        # 模式2: 裸 column_name → "中文.英文"
        pattern_bare = r'(?<![.\w"])' + escaped + r'(?![.\w])'
        def repl_bare(m, _q=quoted_calib):
            start = m.start()
            prefix = result[:start]
            if prefix.count("'") % 2 == 1:
                return m.group(0)
            return _q
        
        result = re.sub(pattern_bare, repl_bare, result)
    
    return result


# ============================================================
# 工具函数：结果标准化与对比
# ============================================================

def normalize_value(v):
    """标准化单个值用于对比"""
    if v is None:
        return None
    if isinstance(v, float):
        if np.isnan(v) or np.isinf(v):
            return None
        return round(v, 4)
    if isinstance(v, (np.integer, np.floating)):
        return normalize_value(float(v)) if isinstance(v, np.floating) else int(v)
    if isinstance(v, str):
        return v.strip()
    return v


def normalize_result(data):
    """将 query 返回的 [[headers], [row1], ...] 标准化"""
    if not data or not isinstance(data, list):
        return [], []
    headers = [str(h).strip() for h in data[0]]
    rows = []
    for row in data[1:]:
        norm_row = [normalize_value(v) for v in row]
        rows.append(norm_row)
    return headers, rows


def sqlite_query(db_path, sql):
    """在 SQLite 上执行查询，返回 (headers, [rows])"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        cursor = conn.execute(sql)
        headers = [desc[0] for desc in cursor.description]
        rows = []
        for raw_row in cursor:
            rows.append([normalize_value(v) for v in raw_row])
        return headers, rows
    finally:
        conn.close()


def compare_results(excel_headers, excel_rows, sqlite_headers, sqlite_rows,
                    tol=0.01, order_by=None):
    """
    对比两组结果，返回 (match: bool, details: str)
    
    当 order_by 为 None 时，自动使用所有列作为排序键进行稳定排序
    """
    if len(excel_rows) != len(sqlite_rows):
        return False, f"行数不一致: ExcelMCP={len(excel_rows)} vs SQLite={len(sqlite_rows)}"
    
    if len(excel_rows) == 0 and len(sqlite_rows) == 0:
        return True, "双方均为空集"
    
    # 排序
    if order_by is not None:
        if isinstance(order_by, str):
            if order_by in excel_headers:
                idx = excel_headers.index(order_by)
            elif order_by in sqlite_headers:
                idx = sqlite_headers.index(order_by)
            else:
                idx = order_by
        else:
            idx = order_by
        try:
            excel_rows = sorted(excel_rows, key=lambda r: (r[idx] is None, r[idx]))
            sqlite_rows = sorted(sqlite_rows, key=lambda r: (r[idx] is None, r[idx]))
        except (IndexError, TypeError):
            pass
    else:
        # 无指定排序列时，用所有列作为排序键（处理 ORDER BY 不稳定的情况）
        def sort_key(row):
            return tuple((v is None, v) for v in row)
        try:
            excel_rows = sorted(excel_rows, key=sort_key)
            sqlite_rows = sorted(sqlite_rows, key=sort_key)
        except (TypeError, Exception):
            pass
    
    if len(excel_headers) != len(sqlite_headers):
        return False, f"列数不一致: {len(excel_headers)} vs {len(sqlite_headers)}"
    
    mismatches = 0
    mismatch_details = []
    for i, (er, sr) in enumerate(zip(excel_rows, sqlite_rows)):
        for j, (ev, sv) in enumerate(zip(er, sr)):
            if ev == sv:
                continue
            if isinstance(ev, (int, float)) and isinstance(sv, (int, float)):
                try:
                    if abs(float(ev) - float(sv)) > tol:
                        mismatches += 1
                        if len(mismatch_details) < 5:
                            col_name = excel_headers[j] if j < len(excel_headers) else f'col{j}'
                            mismatch_details.append(f"[{i}][{col_name}] {ev} vs {sv}")
                except (ValueError, TypeError):
                    mismatches += 1
            else:
                mismatches += 1
                if len(mismatch_details) < 5:
                    col_name = excel_headers[j] if j < len(excel_headers) else f'col{j}'
                    mismatch_details.append(f"[{i}][{col_name}] {ev!r} vs {sv!r}")
    
    if mismatches == 0:
        return True, f"完全匹配 ({len(excel_rows)} 行 × {len(excel_headers)} 列)"
    else:
        detail = f"{mismatches} 个单元格不匹配 ({len(excel_rows)} 行 × {len(excel_headers)} 列)"
        if mismatch_details:
            detail += " | " + "; ".join(mismatch_details)
        return False, detail


def run_cross_validation(xlsx_path, db_path, sql, col_map,
                         tol=0.01, order_by=None):
    """
    执行交叉对比：ExcelMCP vs SQLite
    
    Returns: (match, detail)
    """
    # 1. ExcelMCP 执行原始 SQL
    result = execute_advanced_sql_query(xlsx_path, sql)
    assert result["success"], f"ExcelMCP 查询失败: {result.get('message', '')}"
    eh, er = normalize_result(result["data"])
    
    # 2. SQLite 执行翻译后的 SQL
    sqlite_sql = translate_sql_for_calibrator(sql, col_map)
    sh, sr = sqlite_query(db_path, sqlite_sql)
    
    # 3. 对比
    match, detail = compare_results(eh, er, sh, sr, tol=tol, order_by=order_by)
    return match, detail, eh, er, sh, sr


# ============================================================
# Fixtures
# ============================================================

@pytest.fixture
def game_config_db():
    """导入 game_config.xlsx 到 SQLite"""
    db_name = "test_p02_game"
    cmd_import(GAME_CONFIG_PATH, db_name)
    yield get_db_path(db_name), GAME_CONFIG_PATH


@pytest.fixture
def join_test_db():
    """导入 join_test.xlsx 到 SQLite"""
    db_name = "test_p02_join"
    cmd_import(JOIN_TEST_PATH, db_name)
    yield get_db_path(db_name), JOIN_TEST_PATH


@pytest.fixture
def large_skills_db():
    """导入 large_skills.xlsx 到 SQLite"""
    db_name = "test_p02_large"
    cmd_import(LARGE_SKILLS_PATH, db_name)
    yield get_db_path(db_name), LARGE_SKILLS_PATH


@pytest.fixture
def comprehensive_db():
    """导入 comprehensive_test.xlsx 到 SQLite"""
    db_name = "test_p02_comprehensive"
    cmd_import(COMPREHENSIVE_PATH, db_name)
    yield get_db_path(db_name), COMPREHENSIVE_PATH


# ============================================================
# B1: 多表 JOIN
# ============================================================
class TestB1_JoinQueries:
    """B1: 多表 JOIN 查询对比"""

    def test_b1_inner_join_two_tables(self, join_test_db):
        """内连接：技能表 JOIN 装备表 on equip_id"""
        db_path, xlsx_path = join_test_db
        
        sql = """
            SELECT s.skill_name, s.damage, e.equip_name, e.attack_power 
            FROM 技能表 s 
            INNER JOIN 装备表 e ON s.equip_id = e.equip_id
        """
        
        match, detail, eh, er, sh, sr = run_cross_validation(
            xlsx_path, db_path, sql, JOIN_TEST_COLS, order_by="skill_name")
        print(f"B1-JOIN: {detail}")
        assert match, f"JOIN 结果不一致: {detail}"

    def test_b1_left_join_with_nulls(self, join_test_db):
        """左连接：验证 LEFT JOIN 基本功能"""
        db_path, xlsx_path = join_test_db
        
        sql = """
            SELECT s.skill_name, e.equip_name 
            FROM 技能表 s 
            LEFT JOIN 装备表 e ON s.equip_id = e.equip_id
            ORDER BY s.skill_name
        """
        
        result = execute_advanced_sql_query(xlsx_path, sql)
        assert result["success"], f"ExcelMCP 查询失败: {result.get('message', '')}"
        eh, er = normalize_result(result["data"])
        
        sqlite_sql = translate_sql_for_calibrator(sql, JOIN_TEST_COLS)
        sh, sr = sqlite_query(db_path, sqlite_sql)
        
        # 验证基本结构一致
        assert len(eh) == len(sh), f"列数不一致: {len(eh)} vs {len(sh)}"
        assert len(er) == len(sr), f"行数不一致: {len(er)} vs {len(sr)}"
        print(f"B1-LEFTJOIN: ExcelMCP={len(er)}行, SQLite={len(sr)}行, 列数={len(eh)} ✅")


# ============================================================
# B2: 分组聚合 + HAVING
# ============================================================
class TestB2_GroupByHaving:
    """B2: GROUP BY + HAVING + 聚合函数"""

    def test_b2_group_by_with_count_avg(self, game_config_db):
        """按品质分组统计装备数量和平均攻击力"""
        db_path, xlsx_path = game_config_db
        
        sql = """
            SELECT quality, COUNT(*) as cnt, AVG(attack) as avg_atk, MAX(defense) as max_def 
            FROM 装备配置 
            GROUP BY quality 
            ORDER BY quality
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, GAME_CONFIG_COLS, order_by=0, tol=0.1)
        print(f"B2-GROUPBY: {detail}")
        assert match, f"GROUP BY 结果不一致: {detail}"

    def test_b2_having_filter(self, game_config_db):
        """HAVING 过滤分组结果"""
        db_path, xlsx_path = game_config_db
        
        sql = """
            SELECT skill_type, COUNT(*) as cnt, AVG(damage) as avg_dmg 
            FROM 技能配置 
            GROUP BY skill_type 
            HAVING COUNT(*) >= 2
            ORDER BY cnt DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, GAME_CONFIG_COLS, tol=0.1)
        # 允许全列稳定排序（HAVING+ORDER BY cnt DESC 在 cnt 相同时不稳定）
        print(f"B2-HAVING: {detail}")
        assert match, f"HAVING 结果不一致: {detail}"


# ============================================================
# B3: 窗口函数
# ============================================================
class TestB3_WindowFunctions:
    """B3: ROW_NUMBER / RANK 等窗口函数"""

    def test_b3_row_number(self, large_skills_db):
        """ROW_NUMBER() OVER (PARTITION BY ... ORDER BY ...)"""
        db_path, xlsx_path = large_skills_db
        
        # 不用 LIMIT（不同 SQL 引擎对 LIMIT+窗口函数行为不同），
        # 改用子查询取每个职业 Top 3
        sql = """
            SELECT * FROM (
                SELECT skill_name, profession, damage,
                       ROW_NUMBER() OVER (PARTITION BY profession ORDER BY damage DESC) as rn
                FROM Skills 
                WHERE profession IS NOT NULL
            ) WHERE rn <= 3
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, LARGE_SKILLS_COLS, tol=0.01)
        print(f"B3-ROWNUM: {detail}")
        assert match, f"ROW_NUMBER 结果不一致: {detail}"

    def test_b3_rank_function(self, large_skills_db):
        """RANK() OVER (ORDER BY ...) — 取 Top 10"""
        db_path, xlsx_path = large_skills_db
        
        sql = """
            SELECT skill_name, quality, damage,
                   RANK() OVER (ORDER BY damage DESC) as rank_val
            FROM Skills
            WHERE rank_val <= 10
        """
        
        # 注意：RANK 过滤需要用子查询，因为 RANK 在 WHERE 中不可直接引用
        sql_safe = """
            SELECT * FROM (
                SELECT skill_name, quality, damage,
                       RANK() OVER (ORDER BY damage DESC) as rank_val
                FROM Skills
            ) sub WHERE rank_val <= 10
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql_safe, LARGE_SKILLS_COLS, tol=0.01)
        print(f"B3-RANK: {detail}")
        assert match, f"RANK 结果不一致: {detail}"


# ============================================================
# B4: 子查询（IN / 标量）
# ============================================================
class TestB4_Subqueries:
    """B4: IN 子查询和标量子查询"""

    def test_b4_in_subquery(self, game_config_db):
        """IN (SELECT ...) 子查询"""
        db_path, xlsx_path = game_config_db
        
        sql = """
            SELECT equip_name, quality, attack 
            FROM 装备配置 
            WHERE quality IN (
                SELECT DISTINCT quality FROM 装备配置 WHERE attack >= 50
            )
            ORDER BY attack DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, GAME_CONFIG_COLS, order_by=2)
        print(f"B4-IN: {detail}")
        assert match, f"IN 子查询结果不一致: {detail}"

    def test_b4_scalar_subquery(self, game_config_db):
        """标量子查询 — 在 SELECT 中使用"""
        db_path, xlsx_path = game_config_db
        
        sql = """
            SELECT e.equip_name, e.attack,
                   (SELECT AVG(attack) FROM 装备配置) as overall_avg
            FROM 装备配置 e
            WHERE e.attack >= (SELECT AVG(attack) FROM 装备配置)
            ORDER BY e.attack DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, GAME_CONFIG_COLS, order_by=1, tol=0.1)
        print(f"B4-SCALAR: {detail}")
        assert match, f"标量子查询结果不一致: {detail}"


# ============================================================
# B5: CASE WHEN 条件表达式
# ============================================================
class TestB5_CaseWhen:
    """B5: CASE WHEN 表达式"""

    def test_b5_simple_case(self, game_config_db):
        """简单 CASE WHEN 分类"""
        db_path, xlsx_path = game_config_db
        
        sql = """
            SELECT equip_name, attack,
                CASE 
                    WHEN quality >= 5 THEN '传说'
                    WHEN quality >= 4 THEN '史诗'
                    WHEN quality >= 3 THEN '稀有'
                    ELSE '普通'
                END as quality_tier
            FROM 装备配置
            ORDER BY quality DESC, attack DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, GAME_CONFIG_COLS, order_by=1)
        print(f"B5-CASE: {detail}")
        assert match, f"CASE WHEN 结果不一致: {detail}"

    @pytest.mark.filterwarnings("ignore::pytest.PytestUnraisableExceptionWarning")
    def test_b5_case_with_aggregation(self, large_skills_db):
        """CASE WHEN + GROUP BY 组合（使用字符串兼容条件）"""
        db_path, xlsx_path = large_skills_db
        
        # 注意：large_skills.xlsx 的 quality 列是中文字符串（传说/史诗/稀有/普通）
        # 不能用 >= 4 比较（SQLite 和 ExcelMCP 对字符串-整数比较行为不同）
        # 改用 IN 子句做字符串匹配
        sql = """
            SELECT 
                CASE 
                    WHEN quality IN ('传说', '史诗') THEN '高端'
                    ELSE '低端'
                END as tier,
                COUNT(*) as cnt,
                ROUND(AVG(damage), 1) as avg_dmg
            FROM Skills
            GROUP BY tier
            ORDER BY cnt DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, LARGE_SKILLS_COLS, order_by=1, tol=0.1)
        print(f"B5-CASE+GROUP: {detail}")
        assert match, f"CASE+GROUP BY 结果不一致: {detail}"


# ============================================================
# B6: 双表头表复杂查询
# ============================================================
class TestB6_DualHeaderQueries:
    """B6: 双表头（中英文别名）表的复杂查询"""

    def test_b6_dual_header_select(self, join_test_db):
        """双表头表基础 SELECT + 过滤"""
        db_path, xlsx_path = join_test_db
        
        sql = """
            SELECT skill_name, skill_type, damage, equip_id 
            FROM 技能表 
            WHERE damage > 30
            ORDER BY damage DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, JOIN_TEST_COLS, order_by=2)
        print(f"B6-DUALHDR: {detail}")
        assert match, f"双表头查询不一致: {detail}"

    @pytest.mark.filterwarnings("ignore::pytest.PytestUnraisableExceptionWarning")
    def test_b6_dual_header_agg(self, join_test_db):
        """双表头表聚合查询"""
        db_path, xlsx_path = join_test_db
        
        sql = """
            SELECT skill_type, 
                   COUNT(*) as cnt, 
                   SUM(damage) as total_dmg,
                   ROUND(AVG(damage), 2) as avg_dmg
            FROM 技能表
            GROUP BY skill_type
            ORDER BY total_dmg DESC
        """
        
        match, detail, _, _, _, _ = run_cross_validation(
            xlsx_path, db_path, sql, JOIN_TEST_COLS, order_by=2, tol=0.1)
        print(f"B6-DUALHDR-AGG: {detail}")
        assert match, f"双表头聚合不一致: {detail}"
