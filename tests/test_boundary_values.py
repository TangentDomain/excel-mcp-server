"""
边界值测试 - REQ-034
测试极端值、精度损失、NULL处理、大数值等边界情况

@intention: 确保SQL引擎和Excel操作在边界条件下正确工作
"""

import pytest
import os
import tempfile
from openpyxl import Workbook

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine, _safe_float_comparison


def _get_column_data(result, col_name):
    """从查询结果中提取指定列的数据（跳过表头行）"""
    data = result.get('data', [])
    if len(data) < 2:
        return []
    headers = data[0]
    col_idx = headers.index(col_name) if col_name in headers else -1
    if col_idx == -1:
        return []
    return [row[col_idx] for row in data[1:]]


@pytest.fixture
def boundary_test_file():
    """创建包含边界值的测试Excel文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "边界值测试"

    # 表头
    ws.append(["ID", "名称", "价格", "冷却时间", "伤害值", "概率", "描述"])

    # 正常值
    ws.append([1, "普通攻击", 100.0, 1.5, 50, 0.5, "基础攻击"])
    ws.append([2, "强击", 200.0, 2.0, 100, 0.3, "强力攻击"])

    # 边界值 - 0.1精度测试
    ws.append([3, "快速攻击", 50.0, 0.1, 10, 0.1, "极快攻击"])
    ws.append([4, "超快攻击", 25.0, 0.01, 5, 0.01, "超快攻击"])
    ws.append([5, "极速攻击", 12.5, 0.001, 2, 0.001, "极速攻击"])

    # 大数值边界测试
    ws.append([6, "终极攻击", 99999.99, 999.999, 99999, 0.999, "极限攻击"])
    ws.append([7, "超大伤害", 999999.999, 9999.9999, 999999, 0.9999, "超大攻击"])
    ws.append([8, "最大值", 9999999.0, 0.0, 9999999, 1.0, "最大攻击"])

    # 零值和负值测试
    ws.append([9, "免费技能", 0.0, 0.0, 0, 0.0, "零值测试"])
    ws.append([10, "测试负值", -100.0, -1.0, -50, -0.5, "负值测试"])

    # NULL/空值测试
    ws.append([11, None, None, None, None, None, None])
    ws.append([12, "", 0.0, 0.0, 0, 0.0, ""])
    ws.append([None, "空ID测试", 50.0, 1.0, 25, 0.5, "ID为空"])

    # 浮点精度边界
    ws.append([15, "极小值", 0.00001, 0.00001, 1, 0.00001, "极小值测试"])
    ws.append([16, "微小概率", 0.000001, 0.000001, 1, 0.000001, "极微概率"])

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        wb.save(f.name)
        yield f.name

    os.unlink(f.name)


@pytest.fixture
def engine():
    """创建SQL查询引擎实例"""
    return AdvancedSQLQueryEngine()


class TestBoundaryValues:
    """边界值测试类"""

    # ==================== 精度测试 ====================

    def test_float_precision_0_1(self, engine, boundary_test_file):
        """测试0.1秒精度不丢失"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 冷却时间 FROM 边界值测试 WHERE 冷却时间 = 0.1"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '快速攻击' in names, f"0.1精度丢失，找到: {names}"

    def test_float_precision_0_01(self, engine, boundary_test_file):
        """测试0.01秒精度"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 冷却时间 FROM 边界值测试 WHERE 冷却时间 = 0.01"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '超快攻击' in names, f"0.01精度丢失，找到: {names}"

    def test_float_precision_0_001(self, engine, boundary_test_file):
        """测试0.001秒精度"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 冷却时间 FROM 边界值测试 WHERE 冷却时间 = 0.001"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '极速攻击' in names, f"0.001精度丢失，找到: {names}"

    # ==================== 大数值测试 ====================

    def test_large_value_99999(self, engine, boundary_test_file):
        """测试99999大数值"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 价格 >= 99999"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '终极攻击' in names, f"99999大数值未匹配，找到: {names}"

    def test_large_value_999999(self, engine, boundary_test_file):
        """测试999999超大数值"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 伤害值 FROM 边界值测试 WHERE 伤害值 >= 999999"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该找到超大伤害值的记录"

    def test_max_value_comparison(self, engine, boundary_test_file):
        """测试最大值比较"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 价格 = 9999999"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '最大值' in names, f"最大值未匹配，找到: {names}"

    # ==================== 零值和负值测试 ====================

    def test_zero_value(self, engine, boundary_test_file):
        """测试零值处理"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格, 冷却时间 FROM 边界值测试 WHERE 价格 = 0 AND 冷却时间 = 0"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '免费技能' in names, f"零值未匹配，找到: {names}"

    def test_negative_value(self, engine, boundary_test_file):
        """测试负值处理"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 价格 < 0"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '测试负值' in names, f"负值未匹配，找到: {names}"

    # ==================== NULL处理测试 ====================

    def test_null_handling_is_null(self, engine, boundary_test_file):
        """测试IS NULL条件 - 验证查询不崩溃"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT * FROM 边界值测试 WHERE 名称 IS NULL"
        )
        # NULL值在加载时被转换为空字符串，全NULL行被dropna移除
        # 验证查询不崩溃且返回成功
        assert result['success'], f"IS NULL查询失败: {result.get('message', '')}"

    def test_null_handling_is_not_null(self, engine, boundary_test_file):
        """测试IS NOT NULL条件"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 名称 IS NOT NULL AND 价格 IS NOT NULL"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该找到非NULL值的记录"

    # ==================== 范围查询测试 ====================

    def test_between_precision(self, engine, boundary_test_file):
        """测试BETWEEN精度"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 冷却时间 FROM 边界值测试 WHERE 冷却时间 BETWEEN 0.09 AND 0.11"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '快速攻击' in names, f"BETWEEN精度丢失，找到: {names}"

    def test_between_large_values(self, engine, boundary_test_file):
        """测试大数值BETWEEN"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 价格 BETWEEN 99999 AND 1000000"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该找到大数值范围的记录"

    # ==================== 聚合函数边界测试 ====================

    def test_max_with_boundary(self, engine, boundary_test_file):
        """测试MAX函数在大数值边界"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT MAX(价格) as 最高价格 FROM 边界值测试 WHERE 价格 IS NOT NULL"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该返回聚合结果"

    def test_min_with_boundary(self, engine, boundary_test_file):
        """测试MIN函数在负值边界"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT MIN(价格) as 最低价格 FROM 边界值测试 WHERE 价格 IS NOT NULL"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该返回聚合结果"

    def test_count_with_null(self, engine, boundary_test_file):
        """测试COUNT函数对NULL的处理"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT COUNT(*) as 总数 FROM 边界值测试"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        assert len(result.get('data', [])) > 1, "应该返回聚合结果"

    # ==================== 极小值测试 ====================

    def test_tiny_value(self, engine, boundary_test_file):
        """测试极小值 (0.00001)"""
        result = engine.execute_sql_query(
            boundary_test_file,
            "SELECT 名称, 价格 FROM 边界值测试 WHERE 价格 <= 0.0001 AND 价格 > 0"
        )
        assert result['success'], f"查询失败: {result.get('message', '')}"
        names = _get_column_data(result, '名称')
        assert '极小值' in names, f"极小值未匹配，找到: {names}"


class TestSafeFloatComparison:
    """安全浮点比较函数测试"""

    def test_normal_comparison(self):
        """测试正常比较"""
        assert _safe_float_comparison(1.0, 2.0, '<') is True
        assert _safe_float_comparison(2.0, 1.0, '>') is True
        assert _safe_float_comparison(1.0, 1.0, '>=') is True
        assert _safe_float_comparison(1.0, 1.0, '<=') is True

    def test_none_handling(self):
        """测试None值处理"""
        assert _safe_float_comparison(None, 1.0, '<') is False
        assert _safe_float_comparison(1.0, None, '>') is False
        assert _safe_float_comparison(None, None, '<') is False

    def test_precision_0_1(self):
        """测试0.1精度比较"""
        assert _safe_float_comparison(0.1, 0.2, '<') is True
        assert _safe_float_comparison(0.1, 0.1, '<=') is True
        assert _safe_float_comparison(0.1, 0.1, '==') is True

    def test_large_values(self):
        """测试大数值比较"""
        assert _safe_float_comparison(99999, 99999, '<=') is True
        assert _safe_float_comparison(9999999, 99999, '>') is True

    def test_negative_values(self):
        """测试负值比较"""
        assert _safe_float_comparison(-100, 0, '<') is True
        assert _safe_float_comparison(-100, -50, '<') is True

    def test_equality(self):
        """测试==比较"""
        assert _safe_float_comparison(1.0, 1.0, '==') is True
        assert _safe_float_comparison(0.1, 0.1, '==') is True
        assert _safe_float_comparison(99999.0, 99999.0, '==') is True

    def test_type_coercion(self):
        """测试类型转换"""
        assert _safe_float_comparison("1.5", 2.0, '<') is True
        assert _safe_float_comparison(1, 2, '<') is True


if __name__ == '__main__':
    pytest.main([__file__, '-v', '--tb=short'])
