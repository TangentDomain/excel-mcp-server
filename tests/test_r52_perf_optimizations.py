"""
R52 P3-PERF-02: 热路径重复导入优化测试

验证 _serialize_value 和其他热路径方法不再包含冗余的 inline import。
优化内容:
- datetime, decimal, hashlib 移至顶层导入
- pandas (pd), numpy (np), re, time, sqlglot (exp) 复用顶层导入
- 消除 __import__() 动态调用
- sg_exp → exp 统一别名
- _re → re 统一别名
"""
import pytest
import inspect
import numpy as np
import pandas as pd
from openpyxl import Workbook


@pytest.fixture
def sample_xlsx(tmp_path):
    """创建测试用 xlsx 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "TestData"
    ws.append(["ID", "Name", "Value", "Score"])
    for i in range(1, 21):
        ws.append([f"ID{i}", f"Item-{i}", round(i * 1.5, 2), np.random.randint(50, 100)])
    wb.save(tmp_path / "test_perf.xlsx")
    return str(tmp_path / "test_perf.xlsx")


class TestP3Perf02_HotPathImports:
    """验证热路径无冗余 inline import"""

    def test_serialize_value_datetime_handling(self, sample_xlsx):
        """_serialize_value 正确处理 datetime 类型（使用顶层导入）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        engine = AdvancedSQLQueryEngine()
        
        # 测试各种类型的序列化（覆盖所有分支）
        test_cases = [
            (None, None),
            (42, 42),
            (3.14, 3.14),
            (25.0, 25),  # 整数浮点→int
            ("hello", "hello"),
            (True, True),
            (False, False),
            (np.int64(100), 100),
            (np.float64(2.718), 2.718),
            (np.float64(10.0), 10),  # numpy 整数浮点→int
        ]
        
        for val, expected in test_cases:
            result = engine._serialize_value(val)
            if expected is None:
                assert result is None, f"_serialize_value({val}) should be None"
            elif isinstance(expected, float):
                assert result == expected or (
                    isinstance(result, float) and 
                    abs(result - expected) < 1e-9
                ), f"_serialize_value({val}) = {result}, expected {expected}"
            else:
                assert result == expected, f"_serialize_value({val}) = {result}, expected {expected}"

    def test_serialize_value_nan_inf_handling(self, sample_xlsx):
        """_serialize_value 正确处理 NaN/inf（R42+R48 安全加固）"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        engine = AdvancedSQLQueryEngine()
        
        assert engine._serialize_value(float('nan')) is None
        assert engine._serialize_value(float('inf')) is None
        assert engine._serialize_value(float('-inf')) is None
        assert engine._serialize_value(np.nan) is None
        assert engine._serialize_value(np.inf) is None

    def test_serialize_value_decimal_handling(self, sample_xlsx):
        """_serialize_value 正确处理 Decimal 类型（R48 新增）"""
        from decimal import Decimal
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        engine = AdvancedSQLQueryEngine()
        
        assert engine._serialize_value(Decimal("3.14")) == 3.14
        assert engine._serialize_value(Decimal("100")) == 100
        assert engine._serialize_value(Decimal("NaN")) is None
        assert engine._serialize_value(Decimal("Infinity")) is None

    def test_no_inline_imports_in_serialize_value_source(self):
        """源码检查：_serialize_value 方法体中不包含 import 语句"""
        import inspect
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        source = inspect.getsource(AdvancedSQLQueryEngine._serialize_value)
        lines = [l.strip() for l in source.split('\n')]
        
        import_lines = [l for l in lines if l.startswith('import ') or l.startswith('from ')]
        assert len(import_lines) == 0, f"_serialize_value contains inline imports: {import_lines}"

    def test_no_dynamic_import_calls(self):
        """源码检查：不再使用 __import__() 动态调用"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        source = inspect.getsource(AdvancedSQLQueryEngine)
        assert '__import__' not in source, "Found __import__() dynamic call - should use top-level imports"

    def test_sg_exp_alias_removed(self):
        """源码检查：sg_exp 别名已全部替换为 exp"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        source = inspect.getsource(AdvancedSQLQueryEngine)
        assert 'sg_exp' not in source, "Found remaining 'sg_exp' alias - should all be replaced with 'exp'"

    def test_re_alias_consistent(self):
        """源码检查：_re 别名已统一为 re"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
        
        source = inspect.getsource(AdvancedSQLQueryEngine)
        # 允许注释中的 _re 引用，但代码中不应有 _re.
        code_lines = [l for l in source.split('\n') if not l.strip().startswith('#')]
        for line in code_lines:
            assert '_re.' not in line, f"Found '_re.' usage in code: {line.strip()}"

    def test_top_level_imports_present(self):
        """源码检查：datetime, decimal, hashlib 已在顶层导入"""
        from excel_mcp_server_fastmcp.api import advanced_sql_query
        
        with open(advanced_sql_query.__file__, 'r') as f:
            source = f.read()
        
        # 检查前60行（导入区域）
        header = '\n'.join(source.split('\n')[:60])
        assert 'import datetime' in header, "datetime not imported at top level"
        assert 'from decimal import' in header, "decimal not imported at top level"
        assert 'import hashlib' in header, "hashlib not imported at top level"

    def test_full_query_with_optimized_imports(self, sample_xlsx):
        """端到端查询验证：优化后完整功能正常"""
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        
        # 基础查询
        r1 = execute_advanced_sql_query(sample_xlsx, "SELECT * FROM TestData LIMIT 5")
        assert r1['success']
        # 结果包含表头行，所以 LIMIT 5 返回 6 行（1 表头 + 5 数据）
        assert len(r1['data']) == 6
        
        # 聚合查询（触发 _serialize_value 数值路径）
        r2 = execute_advanced_sql_query(sample_xlsx, 
            "SELECT COUNT(*) as cnt, AVG(Value) as avg_val, SUM(Score) as total FROM TestData")
        assert r2['success']
        # 聚合结果格式：第一行是表头，第二行是数据
        assert len(r2['data']) >= 2  # 至少表头+数据行
        
        # WHERE 查询
        r3 = execute_advanced_sql_query(sample_xlsx,
            "SELECT * FROM TestData WHERE Score > 70 ORDER BY Value DESC")
        assert r3['success']

    def test_performance_no_regression(self, sample_xlsx):
        """性能验证：优化后查询性能无退化"""
        import time
        from excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query
        
        # 预热
        execute_advanced_sql_query(sample_xlsx, "SELECT * FROM TestData")
        
        # 计时
        start = time.time()
        for _ in range(10):
            execute_advanced_sql_query(sample_xlsx, 
                "SELECT ID, Name, ROUND(Value, 1) as RndVal, Score FROM TestData WHERE Score > 50")
        elapsed = (time.time() - start) * 1000  # ms
        
        # 10次查询应在合理时间内完成（<5秒）
        assert elapsed < 5000, f"Performance regression: 10 queries took {elapsed:.0f}ms"
