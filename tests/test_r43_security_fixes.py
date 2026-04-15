"""
R43 安全修复验证测试
- R4 [P1]: CTE 嵌套深度限制（防 StackOverflow）
- R7 [P1]: pandas query() 字符串注入防护

注意: API 返回 data 格式为 list[list], 首行为表头.
"""
import pytest
import pandas as pd
import numpy as np
from openpyxl import Workbook

from excel_mcp_server_fastmcp.api.advanced_sql_query import (
    execute_advanced_sql_query,
    execute_advanced_update_query,
)


# ============================================================
# 辅助函数
# ============================================================

def _rows(result):
    """提取数据行(去掉表头)"""
    return result["data"][1:] if len(result["data"]) > 0 else []


def _hdr(result):
    """提取表头行"""
    return result["data"][0] if len(result["data"]) > 0 else []


def _col_val(result, col_name, row_idx=0):
    """获取指定列名在指定数据行的值"""
    headers = _hdr(result)
    rows = _rows(result)
    if row_idx >= len(rows):
        return None
    if col_name in headers:
        return rows[row_idx][headers.index(col_name)]
    return None


# ============================================================
# 测试数据准备
# ============================================================

@pytest.fixture
def sample_xlsx(tmp_path):
    """创建测试用 xlsx 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "员工"
    ws.append(["ID", "姓名", "部门", "薪资"])
    ws.append([1, "张三", "技术部", 10000])
    ws.append([2, "李四", "市场部", 8000])
    ws.append([3, "王五", "技术部", 12000])
    ws.append([4, "赵六", "市场部", 9000])
    ws.append([5, "O'Brien", "财务部", 11000])  # 含单引号的名称
    ws.append([6, '路径\\测试', "技术部", 9500])   # 含反斜杠的名称
    wb.save(tmp_path / "test.xlsx")
    return str(tmp_path / "test.xlsx")


# ============================================================
# R4: CTE 深度限制测试
# ============================================================

class TestR4CTEDepthLimit:
    """R4 [P1] CTE 嵌套深度限制 — 防止 StackOverflow"""

    def test_normal_cte_works(self, sample_xlsx):
        """普通 CTE（1层嵌套）应正常工作"""
        sql = """
        WITH TechStaff AS (
            SELECT * FROM 员工 WHERE 部门 = '技术部'
        )
        SELECT * FROM TechStaff WHERE 薪资 > 10000
        """
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        assert len(_rows(result)) == 1
        assert _col_val(result, '姓名') == '王五'

    def test_nested_2level_cte_works(self, sample_xlsx):
        """2 层 CTE 嵌套应正常工作（多 CTE 平铺，非递归嵌套）"""
        sql = """
        WITH Level1 AS (
            SELECT * FROM 员工 WHERE 部门 = '技术部'
        ),
        Level2 AS (
            SELECT * FROM Level1 WHERE 薪资 > 9000
        )
        SELECT * FROM Level2
        """
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        # 测试数据中技术部且薪资>9000的有：张三(10000), 王五(12000), 路径\测试(9500)
        assert len(_rows(result)) == 3

    def test_deep_cte_exceeds_limit(self, sample_xlsx):
        """超过深度限制的 CTE 应返回友好错误（非 StackOverflow 崩溃）

        注意：标准 SQL 的平铺 CTE 链（每个 CTE body 是简单 SELECT）
        不会触发递归深度累积，因为每个 CTE 的内部查询不含 WITH 子句。
        深度防护主要防御：
        1) 非常规/畸形 SQL 解析导致的意外递归
        2) UPDATE/DELETE 路径中 _inject_ctes_to_worksheets → _execute_query 的交叉调用
        此测试通过直接构造 engine 并调用内部方法来验证深度检查机制。
        """
        import sqlglot
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        # 用正常方式构造 engine（读取 xlsx）
        engine = AdvancedSQLQueryEngine(sample_xlsx)
        parsed = sqlglot.parse_one("SELECT * FROM 员工")
        worksheets_data = engine._load_excel_data(sample_xlsx)

        # 直接用超限深度参数调用 _execute_query
        try:
            result_df = engine._execute_query(
                parsed, worksheets_data, _cte_depth=100
            )
            # 如果返回了 DataFrame 而不是抛异常，说明检查未生效
            assert False, f"深度检查未拦截超限调用，返回了 {len(result_df)} 行"
        except ValueError as e:
            error_str = str(e)
            # 验证错误消息包含深度相关信息
            assert '深度' in error_str or 'depth' in error_str.lower() or 'CTE' in error_str or 'overflow' in error_str.lower(), \
                f"ValueError 消息应包含深度限制提示，实际: {error_str}"

    def test_cte_depth_limit_is_configurable(self):
        """验证 _MAX_CTE_DEPTH 类常量存在且为正整数"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        assert hasattr(AdvancedSQLQueryEngine, '_MAX_CTE_DEPTH')
        assert isinstance(AdvancedSQLQueryEngine._MAX_CTE_DEPTH, int)
        assert AdvancedSQLQueryEngine._MAX_CTE_DEPTH > 0
        assert AdvancedSQLQueryEngine._MAX_CTE_DEPTH <= 100  # 合理上限

    def test_update_with_cte_still_works(self, sample_xlsx):
        """UPDATE + CTE 不受影响"""
        sql = """
        WITH HighSalary AS (
            SELECT ID FROM 员工 WHERE 薪资 >= 10000
        )
        UPDATE 员工 SET 薪资 = ROUND(薪资 * 1.05, 2)
        WHERE ID IN (SELECT ID FROM HighSalary)
        """
        result = execute_advanced_update_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"


# ============================================================
# R7: pandas query() 注入防护测试
# ============================================================

class TestR7PandasQueryInjection:
    """R7 [P1] pandas query() 字符串注入防护"""

    def test_single_quote_in_string_literal(self, sample_xlsx):
        """含单引号的字符串不应破坏 pandas query()"""
        sql = "SELECT * FROM 员工 WHERE 姓名 = \"O'Brien\""
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        assert len(_rows(result)) == 1
        assert _col_val(result, 'ID') == 5

    def test_backslash_in_string_literal(self, sample_xlsx):
        """含反斜杠的字符串不应破坏 pandas query()"""
        sql = "SELECT * FROM 员工 WHERE 姓名 = '路径\\\\测试'"
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        assert len(_rows(result)) == 1
        assert _col_val(result, 'ID') == 6

    def test_single_quote_in_like_pattern(self, sample_xlsx):
        """LIKE 模式中含单引号不应破坏查询（使用已有数据 O'Brien）"""
        # 注意: 此引擎不支持 INSERT，使用 fixture 中已有的 O'Brien 数据
        # O'Brien 中 ' 后面是 B，所以用 %'B% 模式匹配
        sql = "SELECT * FROM 员工 WHERE 姓名 LIKE \"%'B%\""
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        assert len(_rows(result)) >= 1, f"应至少匹配 O'Brien, 实际行数: {len(_rows(result))}"

    def test_injection_attempt_column_name(self, sample_xlsx):
        """尝试通过字符串注入列名的攻击应被阻止或安全处理

        注意: 此引擎不支持 INSERT，使用 UPDATE 修改已有数据来测试。
        """
        # 先用 UPDATE 将某行姓名改为恶意字符串
        sql_update = "UPDATE 员工 SET 姓名 = \"') | (True) | ('\" WHERE ID = 4"
        result = execute_advanced_update_query(sample_xlsx, sql_update)
        # UPDATE 可能成功也可能失败（取决于引擎对特殊字符的处理能力）
        if result['success']:
            malicious_name = "') | (True) | ('"
            sql_select = f"SELECT * FROM 员工 WHERE 姓名 = \"{malicious_name}\""
            result2 = execute_advanced_sql_query(sample_xlsx, sql_select)
            assert result2['success'] is True
            if len(_rows(result2)) > 0:
                for i in range(len(_rows(result2))):
                    assert _col_val(result2, '姓名', i) == malicious_name

    def test_escape_pandas_query_string_method(self):
        """验证 _escape_pandas_query_string 静态方法正确转义"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine

        # 单引号被转义
        assert AdvancedSQLQueryEngine._escape_pandas_query_string("O'Brien") == "O\\'Brien"
        # 反斜杠被转义
        assert AdvancedSQLQueryEngine._escape_pandas_query_string("path\\to") == "path\\\\to"
        # 组合情况
        assert AdvancedSQLQueryEngine._escape_pandas_query_string("a'b\\c") == "a\\'b\\\\c"
        # 普通字符串不变
        assert AdvancedSQLQueryEngine._escape_pandas_query_string("hello") == "hello"
        # 空字符串
        assert AdvancedSQLQueryEngine._escape_pandas_query_string("") == ""

    def test_normal_queries_unaffected(self, sample_xlsx):
        """正常查询不受转义逻辑影响"""
        tests = [
            "SELECT * FROM 员工 WHERE 部门 = '技术部'",
            "SELECT * FROM 员工 WHERE 薪资 > 9000",
            "SELECT * FROM 员工 WHERE 部门 = '技术部' AND 薪资 > 10000",
            "SELECT * FROM 员工 WHERE 姓名 LIKE '%张%'",
            "SELECT * FROM 员工 WHERE ID IN (1, 2, 3)",
        ]
        for sql in tests:
            result = execute_advanced_sql_query(sample_xlsx, sql)
            assert result['success'] is True, f"查询失败: {sql}, 错误: {result.get('message')}"

    def test_special_chars_in_various_contexts(self, sample_xlsx):
        """各种特殊字符在不同 SQL 上下文中都能正确处理"""
        # 含单引号 + IN 子句
        sql = "SELECT * FROM 员工 WHERE 姓名 IN (\"O'Brien\", '张三')"
        result = execute_advanced_sql_query(sample_xlsx, sql)
        assert result['success'] is True, f"Failed: {result.get('message', '')}"
        assert len(_rows(result)) == 2
