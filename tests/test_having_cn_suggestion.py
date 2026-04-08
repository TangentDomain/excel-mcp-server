"""HAVING空结果建议 - 中文列名/无别名场景精确匹配

Regression test for round 72 pre-existing limitation:
- HAVING AVG(伤害) > 99999 无法匹配到 avg_dmg 列（中文不匹配[a-zA-Z_]+正则）
- 修复: _having_agg_alias_map优先匹配 + 无别名聚合自动注册
"""
import pytest
import sys
sys.path.insert(0, 'src')

from excel_mcp.api.advanced_sql_query import execute_advanced_sql_query

GAME_CONFIG = 'tests/test_data/game_config.xlsx'


class TestHavingChineseColumnSuggestion:
    """HAVING空结果建议正确识别中文列名对应的聚合别名列"""

    def test_having_cn_with_explicit_alias_gt(self):
        """HAVING AVG(伤害) > 99999 (显式别名 avg_dmg) → 精确显示最大值"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, AVG(伤害) as avg_dmg FROM 技能配置 "
            "GROUP BY 技能类型 HAVING AVG(伤害) > 99999"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'avg_dmg' in suggestion, f"应包含列名avg_dmg, 实际: {suggestion}"
        assert '最大值为' in suggestion, f"应显示最大值, 实际: {suggestion}"
        assert '99999' in suggestion, f"应显示阈值99999, 实际: {suggestion}"

    def test_having_cn_with_explicit_alias_lt(self):
        """HAVING COUNT(*) < 1 (显式别名 cnt) → 精确显示最小值"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, COUNT(*) as cnt FROM 技能配置 "
            "GROUP BY 技能类型 HAVING COUNT(*) < 1"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'cnt' in suggestion
        assert '最小值为' in suggestion

    def test_having_cn_with_explicit_alias_sum_gte(self):
        """HAVING SUM(伤害) >= 99999 (显式别名 total) → 精确显示最大值"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, SUM(伤害) as total FROM 技能配置 "
            "GROUP BY 技能类型 HAVING SUM(伤害) >= 99999"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'total' in suggestion
        assert '最大值为' in suggestion

    def test_having_cn_without_alias(self):
        """HAVING AVG(伤害) > 99999 (无别名) → 自动生成别名 avg_damage 仍能匹配"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, AVG(伤害) FROM 技能配置 "
            "GROUP BY 技能类型 HAVING AVG(伤害) > 99999"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'avg_damage' in suggestion, f"应匹配自动别名avg_damage, 实际: {suggestion}"
        assert '最大值为' in suggestion

    def test_having_cn_count_star_no_alias(self):
        """HAVING COUNT(*) > 999 (无别名) → 自动生成别名 count_star"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, COUNT(*) FROM 技能配置 "
            "GROUP BY 技能类型 HAVING COUNT(*) > 999"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'count_star' in suggestion
        assert '最大值为' in suggestion

    def test_having_no_empty_result_no_suggestion(self):
        """HAVING有结果时不生成空结果建议"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT 技能类型, AVG(伤害) as avg_dmg FROM 技能配置 "
            "GROUP BY 技能类型 HAVING AVG(伤害) > 50"
        )
        assert result['success'] is True
        # Should have results (法师 avg damage > 50)
        assert len(result.get('data', [])) > 0
        suggestion = result['query_info'].get('suggestion', '')
        assert 'HAVING' not in suggestion

    def test_having_en_column_still_works(self):
        """HAVING英文列名仍正常（回归测试）"""
        result = execute_advanced_sql_query(
            GAME_CONFIG,
            "SELECT skill_type, AVG(damage) as avg_dmg FROM 技能配置 "
            "GROUP BY skill_type HAVING AVG(damage) > 99999"
        )
        assert result['success'] is True
        suggestion = result['query_info'].get('suggestion', '')
        assert 'avg_dmg' in suggestion
        assert '最大值为' in suggestion
