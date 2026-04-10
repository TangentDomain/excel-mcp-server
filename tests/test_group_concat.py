"""
GROUP_CONCAT 聚合函数测试 - REQ-EXCEL-015
"""
import os
import pytest
import pandas as pd


@pytest.fixture
def dept_skills(tmp_path):
    """部门技能表 - 用于 GROUP_CONCAT 测试"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '技能配置'
    ws.append(['部门', 'dept', '技能名称', 'skill_name'])
    ws.append(['法师', 'mage', '火球术', 'fireball'])
    ws.append(['法师', 'mage', '冰冻术', 'ice'])
    ws.append(['法师', 'mage', '火墙', 'firewall'])
    ws.append(['战士', 'warrior', '斩击', 'slash'])
    ws.append(['战士', 'warrior', '旋风斩', 'whirlwind'])
    ws.append(['牧师', 'priest', '治疗术', 'heal'])
    ws.append(['牧师', 'priest', '圣光术', 'holy'])
    path = str(tmp_path / 'group_concat_test.xlsx')
    wb.save(path)
    return path


def _get_rows(result):
    """提取数据行（返回dict列表，跳过表头行和TOTAL行）"""
    if not result.get('success'):
        return []
    data = result.get('data', [])
    if not data or not isinstance(data[0], list):
        return []
    headers = data[0]
    rows = []
    for row in data[1:]:
        # 跳过 TOTAL 汇总行（第一列为 'TOTAL'）
        if row and row[0] == 'TOTAL':
            continue
        rows.append(dict(zip(headers, row)))
    return rows


class TestGroupConcat:
    """GROUP_CONCAT() 测试"""

    def test_group_concat_basic(self, dept_skills):
        """基本 GROUP_CONCAT: 按部门拼接技能名"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            dept_skills,
            "SELECT 部门, GROUP_CONCAT(技能名称) as 技能列表 FROM 技能配置 GROUP BY 部门"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 3

        # 验证结果
        dept_skills_map = {r['部门']: r['技能列表'] for r in rows}
        # 法师组: 火球术,冰冻术,火墙
        assert '法师' in dept_skills_map
        mage_skills = dept_skills_map['法师']
        assert '火球术' in mage_skills
        assert '冰冻术' in mage_skills
        assert '火墙' in mage_skills

        # 战士组: 斩击,旋风斩
        assert '战士' in dept_skills_map
        warrior_skills = dept_skills_map['战士']
        assert '斩击' in warrior_skills
        assert '旋风斩' in warrior_skills

        # 牧师组: 治疗术,圣光术
        assert '牧师' in dept_skills_map
        priest_skills = dept_skills_map['牧师']
        assert '治疗术' in priest_skills
        assert '圣光术' in priest_skills

    def test_group_concat_with_separator(self, dept_skills):
        """GROUP_CONCAT 自定义分隔符"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            dept_skills,
            "SELECT 部门, GROUP_CONCAT(skill_name, '|') as skills FROM 技能配置 GROUP BY 部门"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 3

        # 验证使用 | 分隔符
        dept_skills_map = {r['部门']: r['skills'] for r in rows}
        mage_skills = dept_skills_map['法师']
        assert '|' in mage_skills
        assert mage_skills == 'fireball|ice|firewall' or mage_skills == 'ice|fireball|firewall' or mage_skills == 'firewall|ice|fireball'

    def test_group_concat_with_count(self, dept_skills):
        """GROUP_CONCAT 与其他聚合函数组合"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            dept_skills,
            "SELECT 部门, COUNT(*) as cnt, GROUP_CONCAT(技能名称) as skills FROM 技能配置 GROUP BY 部门"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 3

        # 验证计数正确
        dept_count_map = {r['部门']: r['cnt'] for r in rows}
        assert dept_count_map['法师'] == 3
        assert dept_count_map['战士'] == 2
        assert dept_count_map['牧师'] == 2

    def test_group_concat_auto_alias(self, dept_skills):
        """GROUP_CONCAT 无别名时自动生成列名"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            dept_skills,
            "SELECT 部门, GROUP_CONCAT(skill_name) FROM 技能配置 GROUP BY 部门"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        assert len(rows) == 3
        # 自动生成的列名应该包含 groupconcat
        headers = result['data'][0]
        assert any('groupconcat' in h.lower() or 'skill_name' in h.lower() for h in headers)

    @pytest.mark.xfail(reason="GROUP_CONCAT + HAVING 组合场景待完善")
    def test_group_concat_having(self, dept_skills):
        """GROUP_CONCAT 与 HAVING 子句"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        result = execute_advanced_sql_query(
            dept_skills,
            "SELECT 部门, GROUP_CONCAT(技能名称) as skills FROM 技能配置 GROUP BY 部门 HAVING COUNT(*) >= 2"
        )
        assert result['success'] is True, f"Query failed: {result.get('message')}"
        rows = _get_rows(result)
        # 所有部门都有 >= 2 个技能，所以应该返回3行
        assert len(rows) == 3
