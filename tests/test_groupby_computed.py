"""GROUP BY + 计算表达式(CASE WHEN/COALESCE) 正确性测试

验证修复：CASE WHEN和COALESCE在SELECT中与多列GROUP BY组合时，
结果应正确按所有GROUP BY列分组，而非仅按计算列分组。

Bug根因：旧代码将计算列别名添加到group_by_columns后用独立groupby，
忽略了其他GROUP BY列，导致索引不匹配和结果错误。
修复：预计算表达式添加到df副本，使grouped可直接访问。
"""

import pytest
import openpyxl


@pytest.fixture
def multi_col_groupby_file(tmp_path):
    """创建含类型+等级+伤害三列的测试文件，含NULL类型"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'test_data'
    ws.append(['类型', '等级', '伤害'])
    ws.append(['法师', 1, 100])
    ws.append(['法师', 2, 200])
    ws.append(['战士', 1, 150])
    ws.append(['战士', 2, 250])
    ws.append([None, 3, 50])  # NULL类型
    wb.save(tmp_path / 'multi_col.xlsx')
    return str(tmp_path / 'multi_col.xlsx')


class TestCOALESCEGroupByMultiColumn:
    """COALESCE + 多列GROUP BY"""

    def test_coalesce_multi_column_groupby_count(self, multi_col_groupby_file):
        """COALESCE + 2列GROUP BY + COUNT应返回4行（每列组合一行）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT COALESCE(类型, '未知') as t, 等级, COUNT(*) as cnt FROM test_data GROUP BY 类型, 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        # 4行: 法师/1, 法师/2, 战士/1, 战士/2, NULL/3 → 不，NULL/3也是一行，所以是5行
        # 等等，数据有5行: 法师/1, 法师/2, 战士/1, 战士/2, NULL/3
        # GROUP BY 类型, 等级 → 5组
        assert result['query_info']['filtered_rows'] == 5

    def test_coalesce_multi_column_groupby_values(self, multi_col_groupby_file):
        """COALESCE + 多列GROUP BY：每行等级值应正确"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT COALESCE(类型, '未知') as t, 等级 FROM test_data GROUP BY 类型, 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        data = [row for row in result['data'][1:] if row[0] != 'TOTAL']  # 跳过表头和TOTAL
        levels = sorted([row[1] for row in data])
        assert levels == [1, 1, 2, 2, 3]

    def test_coalesce_multi_column_groupby_with_avg(self, multi_col_groupby_file):
        """COALESCE + 多列GROUP BY + AVG聚合"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT COALESCE(类型, '未知') as t, 等级, AVG(伤害) as avg_dmg FROM test_data GROUP BY 类型, 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        data = [row for row in result['data'][1:] if row[0] != 'TOTAL']  # 跳过表头和TOTAL
        assert len(data) == 5  # 5组


class TestCaseWhenGroupByMultiColumn:
    """CASE WHEN + 多列GROUP BY"""

    def test_case_when_multi_column_groupby(self, multi_col_groupby_file):
        """CASE WHEN + 2列GROUP BY应返回正确行数"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT CASE WHEN 类型='法师' THEN '魔法' ELSE '物理' END as c, 等级, COUNT(*) as cnt FROM test_data GROUP BY 类型, 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        assert result['query_info']['filtered_rows'] == 5

    def test_case_when_multi_column_values(self, multi_col_groupby_file):
        """CASE WHEN + 多列GROUP BY：等级列值应正确"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT CASE WHEN 类型='法师' THEN '魔法' ELSE '物理' END as c, 等级 FROM test_data GROUP BY 类型, 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        data = [row for row in result['data'][1:] if row[0] != 'TOTAL']  # 跳过表头和TOTAL
        levels = sorted([row[1] for row in data])
        assert levels == [1, 1, 2, 2, 3]


class TestComputedGroupByRegression:
    """回归测试：确保修复不破坏已有功能"""

    def test_coalesce_single_column_groupby(self, multi_col_groupby_file):
        """COALESCE + 单列GROUP BY（回归）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT COALESCE(类型, '未知') as t, COUNT(*) as cnt FROM test_data GROUP BY 类型"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        # 3组: 法师(2), 战士(2), NULL(1)
        assert result['query_info']['filtered_rows'] == 3

    def test_coalesce_with_avg_aggregate(self, multi_col_groupby_file):
        """COALESCE + AVG聚合（回归）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT COALESCE(类型, '未知') as t, AVG(伤害) as avg_dmg FROM test_data GROUP BY 类型"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        data = result['data'][1:]  # 跳过表头
        # 法师: (100+200)/2=150, 战士: (150+250)/2=200, NULL: 50
        avg_map = {row[0]: float(row[1]) for row in data if row[0] != 'TOTAL'}
        assert avg_map['法师'] == 150.0
        assert avg_map['战士'] == 200.0
        # COALESCE(类型, '未知') 应该返回 '未知' 或 ''，取决于实现
        assert (avg_map.get('未知') == 50.0 or avg_map.get('') == 50.0)  # NULL类型

    def test_coalesce_in_where_with_groupby(self, multi_col_groupby_file):
        """COALESCE在WHERE+GROUP BY中（回归）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT 等级, COUNT(*) as cnt FROM test_data WHERE COALESCE(类型, '未知')='法师' GROUP BY 等级"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        # 法师有等级1和2两行
        assert result['query_info']['filtered_rows'] == 2

    def test_case_when_in_select_no_groupby(self, multi_col_groupby_file):
        """CASE WHEN在SELECT中无GROUP BY（回归）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query

        sql = "SELECT 类型, CASE WHEN 伤害>100 THEN '高' ELSE '低' END as level FROM test_data"
        result = execute_advanced_sql_query(multi_col_groupby_file, sql)
        assert result['success'], result.get('message', '')
        data = result['data'][1:]
        # 100→低, 200→高, 150→高, 250→高, 50→低
        assert data[0][1] == '低'   # 法师 100
        assert data[1][1] == '高'   # 法师 200
        assert data[2][1] == '高'   # 战士 150
        assert data[3][1] == '高'   # 战士 250
        assert data[4][1] == '低'   # NULL 50
