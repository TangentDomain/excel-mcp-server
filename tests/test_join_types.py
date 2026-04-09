"""测试 RIGHT JOIN, FULL JOIN, CROSS JOIN 支持"""
import os
import sys
import pytest
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

# xdist isolation: run all join type tests in the same worker to avoid file I/O races
pytestmark = pytest.mark.xdist_group("join_types")


def _data_rows(result):
    """从查询结果中提取数据行（跳过header行）"""
    data = result.get('data', [])
    if not data:
        return []
    return data[1:] if len(data) > 1 else []


# ============ 测试数据 fixture ============

@pytest.fixture
def engine():
    """创建SQL引擎实例"""
    return AdvancedSQLQueryEngine()


@pytest.fixture
def join_fixtures(tmp_path):
    """创建JOIN测试用的Excel文件
    
    技能表:
    | skill_id | skill_name | type   | damage |
    |----------|------------|--------|--------|
    | 1        | 火球术     | 法术   | 200    |
    | 2        | 斩击       | 物理   | 150    |
    | 3        | 治疗术     | 法术   | 0      |
    | 4        | 冰冻术     | 法术   | 180    |
    
    技能解锁表:
    | skill_id | level_req | cost |
    |----------|-----------|------|
    | 1        | 5         | 100  |
    | 2        | 1         | 50   |
    | 5        | 10        | 200  |
    """
    import uuid
    file_path = str(tmp_path / f"join_test_{uuid.uuid4().hex[:8]}.xlsx")
    
    skills = pd.DataFrame({
        'skill_id': [1, 2, 3, 4],
        'skill_name': ['火球术', '斩击', '治疗术', '冰冻术'],
        'type': ['法术', '物理', '法术', '法术'],
        'damage': [200, 150, 0, 180]
    })
    
    unlocks = pd.DataFrame({
        'skill_id': [1, 2, 5],
        'level_req': [5, 1, 10],
        'cost': [100, 50, 200]
    })
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        skills.to_excel(writer, sheet_name='技能表', index=False)
        unlocks.to_excel(writer, sheet_name='解锁表', index=False)
    
    return file_path


# ============ RIGHT JOIN 测试 ============

class TestRightJoin:
    """RIGHT JOIN: 保留右表所有行，左表无匹配则为NULL"""
    
    def test_basic_right_join(self, engine, join_fixtures):
        """基本RIGHT JOIN — 右表全部保留"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_id, a.skill_name, b.level_req, b.cost FROM 技能表 a RIGHT JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 右表有3行(1,2,5)，RIGHT JOIN保留全部
        assert len(rows) == 3
    
    def test_right_join_unmatched_null(self, engine, join_fixtures):
        """RIGHT JOIN — 无匹配行左表列为空"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.cost FROM 技能表 a RIGHT JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 3行: 匹配(火球术,100), (斩击,50), 不匹配('',200)
        assert len(rows) == 3
        # 最后一行cost=200, skill_name为空
        assert rows[-1][1] == 200
        assert rows[-1][0] in (None, '', 'None')
    
    def test_right_join_from_right_table(self, engine, join_fixtures):
        """RIGHT JOIN以右表为FROM — 等效于LEFT JOIN"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT b.skill_id, b.cost, a.skill_name FROM 解锁表 b RIGHT JOIN 技能表 a ON b.skill_id = a.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # FROM 解锁表 RIGHT JOIN 技能表 = 保留技能表全部4行
        assert len(rows) == 4
    
    def test_right_join_with_where(self, engine, join_fixtures):
        """RIGHT JOIN + WHERE过滤"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.level_req, b.cost FROM 技能表 a RIGHT JOIN 解锁表 b ON a.skill_id = b.skill_id WHERE b.level_req > 3"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # level_req > 3: skill_id=1(level 5) 和 skill_id=5(level 10)
        assert len(rows) == 2
    
    def test_right_join_with_aggregation(self, engine, join_fixtures):
        """RIGHT JOIN + GROUP BY聚合"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.type, COUNT(*) as cnt FROM 技能表 a RIGHT JOIN 解锁表 b ON a.skill_id = b.skill_id GROUP BY a.type"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) >= 1


# ============ FULL JOIN 测试 ============

class TestFullJoin:
    """FULL JOIN: 保留两表所有行，无匹配则为NULL"""
    
    def test_basic_full_join(self, engine, join_fixtures):
        """基本FULL JOIN — 两表所有行都保留"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_id, a.skill_name, b.level_req FROM 技能表 a FULL JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 技能表4行 + 解锁表独有1行(skill_id=5) - 匹配行(1,2)不重复 = 5行
        assert len(rows) == 5
    
    def test_full_join_left_only(self, engine, join_fixtures):
        """FULL JOIN — 验证左表独有行(技能3,4在解锁表中不存在)"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_id, a.skill_name, b.level_req FROM 技能表 a FULL JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 技能表独有: skill_id=3(治疗术), skill_id=4(冰冻术) → level_req为空
        left_only = [r for r in rows if r[0] in (3, 4) and r[2] in (None, '', 'None')]
        assert len(left_only) == 2
    
    def test_full_join_right_only(self, engine, join_fixtures):
        """FULL JOIN — 验证右表独有行(skill_id=5在技能表中不存在)"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_id, a.skill_name, b.level_req FROM 技能表 a FULL JOIN 解锁表 b ON a.skill_id = b.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 解锁表独有: skill_id=5 → 左表列为空，level_req=10
        right_only = [r for r in rows if r[2] == 10]
        assert len(right_only) == 1
    
    def test_full_join_with_order(self, engine, join_fixtures):
        """FULL JOIN + ORDER BY"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_id, a.skill_name, b.level_req FROM 技能表 a FULL JOIN 解锁表 b ON a.skill_id = b.skill_id ORDER BY a.skill_id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 5


# ============ CROSS JOIN 测试 ============

class TestCrossJoin:
    """CROSS JOIN: 笛卡尔积"""
    
    def test_basic_cross_join(self, engine, join_fixtures):
        """基本CROSS JOIN — 笛卡尔积 4×3=12"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.cost FROM 技能表 a CROSS JOIN 解锁表 b"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 12
    
    def test_cross_join_with_where(self, engine, join_fixtures):
        """CROSS JOIN + WHERE（过滤笛卡尔积结果）"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.cost FROM 技能表 a CROSS JOIN 解锁表 b WHERE a.damage = 200"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # damage=200只有火球术, CROSS JOIN解锁表3行 = 3行
        assert len(rows) == 3
    
    def test_cross_join_with_limit(self, engine, join_fixtures):
        """CROSS JOIN + LIMIT"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.cost FROM 技能表 a CROSS JOIN 解锁表 b LIMIT 5"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 5
    
    def test_cross_join_no_on_clause(self, engine, join_fixtures):
        """CROSS JOIN不需要ON条件"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.skill_name, b.cost FROM 技能表 a CROSS JOIN 解锁表 b"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 12


# ============ 边界条件测试 ============

class TestJoinEdgeCases:
    """JOIN边界条件"""
    
    def test_no_match_right_join(self, engine, tmp_path):
        """RIGHT JOIN无匹配行"""
        file_path = str(tmp_path / "no_match.xlsx")
        
        skills = pd.DataFrame({'id': [1, 2], 'name': ['A', 'B']})
        other = pd.DataFrame({'id': [10, 20], 'val': ['X', 'Y']})
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            skills.to_excel(writer, sheet_name='技能', index=False)
            other.to_excel(writer, sheet_name='数据', index=False)
        
        result = engine.execute_sql_query(
            file_path,
            "SELECT a.name, b.val FROM 技能 a RIGHT JOIN 数据 b ON a.id = b.id"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 无匹配 RIGHT JOIN → 右表2行保留，左表列为空
        assert len(rows) == 2
        assert rows[0][0] in (None, '', 'None')  # name is empty
        assert rows[0][1] == 'X'
    
    def test_single_row_cross_join(self, engine, tmp_path):
        """单行表CROSS JOIN"""
        file_path = str(tmp_path / "single_row.xlsx")
        
        a = pd.DataFrame({'x': [1]})
        b = pd.DataFrame({'y': [10, 20, 30]})
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            a.to_excel(writer, sheet_name='A', index=False)
            b.to_excel(writer, sheet_name='B', index=False)
        
        result = engine.execute_sql_query(
            file_path,
            "SELECT * FROM A a CROSS JOIN B b"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 3
    
    def test_cross_join_large_cartesian(self, engine, tmp_path):
        """较大表的CROSS JOIN笛卡尔积验证"""
        file_path = str(tmp_path / "cross_large.xlsx")
        
        left = pd.DataFrame({'a': range(1, 6)})   # 5行
        right = pd.DataFrame({'b': range(1, 4)})  # 3行
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            left.to_excel(writer, sheet_name='L', index=False)
            right.to_excel(writer, sheet_name='R', index=False)
        
        result = engine.execute_sql_query(
            file_path,
            "SELECT * FROM L a CROSS JOIN R b"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        assert len(rows) == 15  # 5 × 3
    
    def test_cross_join_with_aggregation(self, engine, join_fixtures):
        """CROSS JOIN + GROUP BY聚合"""
        result = engine.execute_sql_query(
            join_fixtures,
            "SELECT a.type, SUM(b.cost) as total_cost FROM 技能表 a CROSS JOIN 解锁表 b GROUP BY a.type"
        )
        assert result['success'] is True
        rows = _data_rows(result)
        # 应有法术和物理两个分组
        assert len(rows) >= 1
