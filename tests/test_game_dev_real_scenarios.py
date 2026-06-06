"""真实游戏开发场景综合测试：赛季平衡性调整工作流。

场景覆盖：
1. SELECT + WHERE + ORDER BY: 筛选高伤害技能
2. GROUP BY + HAVING + 聚合: 按职业统计伤害分布
3. CTE + 窗口函数: 技能效率排名（伤害/冷却比）
4. UPDATE + CASE WHEN: 赛季平衡性批量调整
5. INSERT + 验证: 新赛季新增技能
6. DELETE + 验证: 移除退役装备
7. 跨表 JOIN: 玩家装备关联分析
8. 子查询: 找出超模组合
9. CASE WHEN + 别名: 生成平衡性报告
10. LIKE + 字符串函数: 模糊搜索与拼接
11. 完整生命周期: INSERT→UPDATE→DELETE→验证回滚
"""

import pandas as pd
import pytest

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

# ─── 测试数据 ───

SKILLS = {
    "技能ID": ["SK001", "SK002", "SK003", "SK004", "SK005", "SK006", "SK007", "SK008", "SK009", "SK010", "SK011", "SK012"],
    "技能名称": ["火球术", "冰冻术", "突袭", "火球术·强化", "冰风暴", "旋风斩", "暗影突袭", "致命一击", "治疗术", "圣光盾", "雷霆一击", "毒刃"],
    "技能类型": ["法师", "法师", "战士", "法师", "法师", "战士", "刺客", "刺客", "辅助", "辅助", "战士", "刺客"],
    "等级": [1, 1, 1, 5, 5, 5, 1, 5, 1, 5, 3, 3],
    "伤害": [100, 80, 120, 200, 180, 150, 90, 250, 0, 0, 160, 110],
    "冷却时间": [5, 8, 3, 8, 10, 6, 4, 12, 10, 15, 5, 3],
    "消耗": [30, 40, 20, 60, 80, 50, 25, 70, 50, 90, 45, 30],
}

EQUIP = {
    "装备ID": ["EQ001", "EQ002", "EQ003", "EQ004", "EQ005", "EQ006", "EQ007", "EQ008", "EQ009", "EQ010"],
    "装备名称": ["铁剑", "法杖", "精钢剑", "传说之剑", "皮甲", "锁子甲", "暗影匕首", "龙鳞甲", "学者之书", "废弃木棍"],
    "装备类型": ["武器", "武器", "武器", "武器", "防具", "防具", "武器", "防具", "武器", "武器"],
    "品质": [1, 2, 3, 5, 1, 2, 4, 5, 3, 0],
    "攻击力": [15, 20, 35, 80, 0, 0, 60, 0, 45, 2],
    "防御力": [0, 0, 0, 0, 10, 20, 0, 50, 10, 0],
    "生命值": [0, 50, 0, 100, 50, 100, 0, 300, 0, 0],
}

PLAYERS = {
    "玩家ID": ["P001", "P002", "P003", "P004", "P005"],
    "玩家名": ["阿尔萨斯", "吉安娜", "瓦尔登", "乌瑟尔", "新手小明"],
    "职业": ["战士", "法师", "刺客", "辅助", "战士"],
    "等级": [10, 8, 12, 6, 1],
    "金币": [5000, 3200, 8000, 2000, 100],
    "装备ID": ["EQ004", "EQ009", "EQ007", "EQ002", "EQ001"],
}


@pytest.fixture
def skills_file(tmp_path):
    p = str(tmp_path / "skills.xlsx")
    pd.DataFrame(SKILLS).to_excel(p, index=False, sheet_name="技能表")
    return p


@pytest.fixture
def equip_file(tmp_path):
    p = str(tmp_path / "equipment.xlsx")
    pd.DataFrame(EQUIP).to_excel(p, index=False, sheet_name="装备表")
    return p


@pytest.fixture
def player_file(tmp_path):
    p = str(tmp_path / "players.xlsx")
    pd.DataFrame(PLAYERS).to_excel(p, index=False, sheet_name="玩家表")
    return p


@pytest.fixture
def engine():
    return AdvancedSQLQueryEngine()


def _q(engine, file_path, sql):
    """Helper: execute SQL without headers, return data rows."""
    r = engine.execute_sql_query(file_path, sql, include_headers=False)
    assert r["success"], f"SQL failed: {r.get('message', '')}\nSQL: {sql}"
    return r["data"]


# ─── 场景 1: 筛选高伤害技能 ───


class TestScenario1HighDamageFilter:
    def test_filter_high_damage_ordered(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能名称, 伤害, 冷却时间 FROM 技能表 WHERE 伤害 > 100 ORDER BY 伤害 DESC")
        assert len(rows) == 7  # 暗影突袭90不满足>100
        assert rows[0][0] == "致命一击" and rows[0][1] == 250
        for i in range(len(rows) - 1):
            assert rows[i][1] >= rows[i + 1][1]

    def test_top3_dps(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能名称, ROUND(伤害 * 1.0 / 冷却时间, 1) AS DPS FROM 技能表 WHERE 伤害 > 0 ORDER BY DPS DESC LIMIT 3")
        assert len(rows) == 3
        assert rows[0][0] == "突袭"  # 120/3=40


# ─── 场景 2: 职业聚合统计 ───


class TestScenario2ClassAggregation:
    def test_group_by_class(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能类型, COUNT(*) AS cnt, ROUND(AVG(伤害), 1) AS avg_dmg, SUM(消耗) AS total_cost FROM 技能表 GROUP BY 技能类型 ORDER BY total_cost DESC")
        assert len(rows) == 4
        mage = [r for r in rows if r[0] == "法师"][0]
        assert mage[1] == 4  # 法师: SK001,SK002,SK004,SK005

    def test_having_filter(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能类型, AVG(伤害) AS avg_dmg FROM 技能表 WHERE 伤害 > 0 GROUP BY 技能类型 HAVING avg_dmg > 100")
        assert len(rows) >= 2


# ─── 场景 3: CTE + 窗口函数 ───


class TestScenario3CTEWindowRanking:
    def test_dps_ranking(self, engine, skills_file):
        rows = _q(
            engine,
            skills_file,
            "WITH DPS AS ("
            "  SELECT 技能名称, 技能类型, 伤害, 冷却时间, "
            "  ROUND(伤害 * 1.0 / 冷却时间, 2) AS DPS "
            "  FROM 技能表 WHERE 伤害 > 0"
            ") SELECT 技能名称, 技能类型, DPS, "
            "RANK() OVER (ORDER BY DPS DESC) AS 排名 FROM DPS",
        )
        assert len(rows) == 10
        top1 = [r for r in rows if int(r[3]) == 1]
        assert len(top1) == 1 and top1[0][0] == "突袭"

    def test_row_number_partition(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能名称, 技能类型, 伤害, ROW_NUMBER() OVER (PARTITION BY 技能类型 ORDER BY 伤害 DESC) AS rank FROM 技能表 WHERE 伤害 > 0")
        from collections import defaultdict

        by_class = defaultdict(list)
        for row in rows:
            by_class[row[1]].append(int(row[3]))
        for cls, ranks in by_class.items():
            assert sorted(ranks) == list(range(1, len(ranks) + 1))


# ─── 场景 4: 赛季平衡性 UPDATE ───


class TestScenario4BalancePatch:
    def test_case_when_balance(self, engine, skills_file):
        orig = {r[0]: r[2] for r in _q(engine, skills_file, "SELECT 技能ID, 技能类型, 伤害 FROM 技能表 ORDER BY 技能ID")}

        engine.execute_update_query(
            skills_file, "UPDATE 技能表 SET 伤害 = CASE WHEN 技能类型 = '法师' THEN ROUND(伤害 * 0.85) WHEN 技能类型 = '战士' THEN ROUND(伤害 * 1.10) ELSE 伤害 END WHERE 伤害 > 0"
        )

        after = {r[0]: (r[1], r[2]) for r in _q(engine, skills_file, "SELECT 技能ID, 技能类型, 伤害 FROM 技能表 ORDER BY 技能ID")}
        for sid, (stype, dmg) in after.items():
            o = orig[sid]
            if o == 0:
                assert dmg == 0
            elif stype == "法师":
                assert dmg == round(o * 0.85), f"{sid}: {o}->{dmg}"
            elif stype == "战士":
                assert dmg == round(o * 1.10), f"{sid}: {o}->{dmg}"
            else:
                assert dmg == o

    def test_auxiliary_unchanged(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 伤害 FROM 技能表 WHERE 技能类型 = '辅助'")
        for row in rows:
            assert row[0] == 0


# ─── 场景 5: 新增技能 ───


class TestScenario5NewSkillInsert:
    def test_insert_and_verify(self, engine, skills_file):
        n_before = _q(engine, skills_file, "SELECT COUNT(*) FROM 技能表")[0][0]

        engine.execute_insert_query(
            skills_file, "INSERT INTO 技能表 (技能ID, 技能名称, 技能类型, 等级, 伤害, 冷却时间, 消耗) VALUES ('SK013', '陨石术', '法师', 8, 300, 15, 100), ('SK014', '反击盾', '战士', 3, 50, 8, 35)"
        )

        n_after = _q(engine, skills_file, "SELECT COUNT(*) FROM 技能表")[0][0]
        assert n_after == n_before + 2

        meteor = _q(engine, skills_file, "SELECT 技能名称, 伤害, 冷却时间 FROM 技能表 WHERE 技能ID = 'SK013'")
        assert meteor[0] == ["陨石术", 300, 15]


# ─── 场景 6: 删除退役装备 ───


class TestScenario6RemoveDeprecated:
    def test_delete_and_verify(self, engine, equip_file):
        before = _q(engine, equip_file, "SELECT 装备名称 FROM 装备表 WHERE 品质 = 0")
        assert len(before) >= 1
        assert "废弃木棍" in [r[0] for r in before]

        engine.execute_delete_query(equip_file, "DELETE FROM 装备表 WHERE 品质 = 0")

        remaining = _q(engine, equip_file, "SELECT COUNT(*) FROM 装备表 WHERE 品质 = 0")
        assert remaining[0][0] == 0

        total = _q(engine, equip_file, "SELECT COUNT(*) FROM 装备表")
        assert total[0][0] == 9


# ─── 场景 7: 跨表 JOIN ───


class TestScenario7CrossTableJoin:
    def test_player_equipment(self, engine, player_file, equip_file):
        """跨文件 JOIN 需要用 表名@'路径' 语法。"""
        import os

        equip_dir = os.path.dirname(equip_file)
        equip_name = os.path.basename(equip_file)
        rows = _q(engine, player_file, f"SELECT 玩家名, 职业, 装备名称, 攻击力 FROM 玩家表 LEFT JOIN 装备表@'{equip_dir}/{equip_name}' ON 玩家表.装备ID = 装备表.装备ID ORDER BY 攻击力 DESC")
        names = [r[0] for r in rows]
        assert "阿尔萨斯" in names
        # 阿尔萨斯拿传说之剑(攻击80)应该排第一
        top = [r for r in rows if r[0] == "阿尔萨斯"][0]
        assert top[2] == "传说之剑"

    def test_overpowered_combo(self, engine, player_file, equip_file):
        """跨表JOIN + 子查询：找出攻击力超过平均值的玩家。"""
        import os

        equip_dir = os.path.dirname(equip_file)
        equip_name = os.path.basename(equip_file)
        rows = _q(
            engine,
            player_file,
            f"SELECT 玩家名, 装备名称, 攻击力 "
            f"FROM 玩家表 "
            f"LEFT JOIN 装备表@'{equip_dir}/{equip_name}' ON 玩家表.装备ID = 装备表.装备ID "
            f"WHERE 攻击力 > (SELECT AVG(攻击力) FROM 装备表@'{equip_dir}/{equip_name}') "
            f"ORDER BY 攻击力 DESC",
        )
        names = [r[0] for r in rows]
        # 攻击力: 传说之剑80, 暗影匕首60, 精钢剑35, 学者之书45, 铁剑15, 法杖20 → avg≈42.5
        assert "阿尔萨斯" in names  # 传说之剑80 > avg
        assert "瓦尔登" in names  # 暗影匕首60 > avg
        assert "新手小明" not in names  # 铁剑15 < avg


# ─── 场景 8: CASE WHEN 平衡性报告 ───


class TestScenario8BalanceReport:
    def test_tier_report(self, engine, skills_file):
        rows = _q(
            engine,
            skills_file,
            "SELECT 技能名称, 技能类型, 伤害, 冷却时间, "
            "CASE "
            "  WHEN 伤害 * 1.0 / 冷却时间 >= 30 THEN 'T0' "
            "  WHEN 伤害 * 1.0 / 冷却时间 >= 20 THEN 'T1' "
            "  WHEN 伤害 * 1.0 / 冷却时间 >= 10 THEN 'T2' "
            "  ELSE 'T3' "
            "END AS tier "
            "FROM 技能表 WHERE 伤害 > 0 ORDER BY 伤害 * 1.0 / 冷却时间 DESC",
        )
        assert len(rows) == 10
        assert rows[0][4] == "T0"  # 突袭 40 DPS
        valid = {"T0", "T1", "T2", "T3"}
        for row in rows:
            assert row[4] in valid

    def test_equipment_quality_report(self, engine, equip_file):
        rows = _q(
            engine,
            equip_file,
            "SELECT "
            "CASE "
            "  WHEN 品质 >= 5 THEN '传说' "
            "  WHEN 品质 >= 4 THEN '史诗' "
            "  WHEN 品质 >= 3 THEN '稀有' "
            "  WHEN 品质 >= 2 THEN '优秀' "
            "  ELSE '普通' "
            "END AS quality, "
            "COUNT(*) AS cnt, ROUND(AVG(攻击力 + 防御力), 1) AS avg_power "
            "FROM 装备表 GROUP BY quality ORDER BY avg_power DESC",
        )
        assert len(rows) >= 4
        legendary = [r for r in rows if r[0] == "传说"]
        assert len(legendary) == 1
        assert legendary[0][1] == 2  # 传说之剑+龙鳞甲


# ─── 场景 9: LIKE + 字符串函数 ───


class TestScenario9StringOps:
    def test_like_search(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能名称 FROM 技能表 WHERE 技能名称 LIKE '%火%' OR 技能名称 LIKE '%冰%' ORDER BY 技能名称")
        names = [r[0] for r in rows]
        assert "火球术" in names
        assert "冰冻术" in names

    def test_concat_label(self, engine, skills_file):
        rows = _q(engine, skills_file, "SELECT 技能名称 || '(' || 技能类型 || ')' AS label FROM 技能表 WHERE 等级 >= 5 ORDER BY 技能名称")
        assert len(rows) == 5
        assert "致命一击(刺客)" in [r[0] for r in rows]


# ─── 场景 10: 完整生命周期 ───


class TestScenario10Lifecycle:
    def test_insert_update_delete_rollback(self, engine, skills_file):
        engine.execute_insert_query(skills_file, "INSERT INTO 技能表 (技能ID, 技能名称, 技能类型, 等级, 伤害, 冷却时间, 消耗) VALUES ('SK999', '测试技能', '测试', 1, 999, 1, 1)")
        assert _q(engine, skills_file, "SELECT 伤害 FROM 技能表 WHERE 技能ID = 'SK999'")[0][0] == 999

        engine.execute_update_query(skills_file, "UPDATE 技能表 SET 伤害 = 500 WHERE 技能ID = 'SK999'")
        assert _q(engine, skills_file, "SELECT 伤害 FROM 技能表 WHERE 技能ID = 'SK999'")[0][0] == 500

        engine.execute_delete_query(skills_file, "DELETE FROM 技能表 WHERE 技能ID = 'SK999'")
        assert _q(engine, skills_file, "SELECT COUNT(*) FROM 技能表 WHERE 技能ID = 'SK999'")[0][0] == 0

    def test_multi_column_update(self, engine, equip_file):
        engine.execute_update_query(equip_file, "UPDATE 装备表 SET 攻击力 = 攻击力 * 2, 防御力 = 防御力 * 2 WHERE 品质 >= 4")
        rows = _q(engine, equip_file, "SELECT 装备名称, 攻击力, 防御力 FROM 装备表 WHERE 品质 >= 4 ORDER BY 攻击力 DESC")
        names = [r[0] for r in rows]
        assert "传说之剑" in names  # 80*2=160
        assert "暗影匕首" in names  # 60*2=120
