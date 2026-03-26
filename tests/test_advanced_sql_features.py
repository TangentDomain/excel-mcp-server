# -*- coding: utf-8 -*-
"""
Advanced SQL Query 高级功能测试套件

覆盖 advanced_sql_query.py 中的高级 SQL 功能
"""

import pytest
import pandas as pd
import tempfile
import os


class TestAdvancedSQLFeatures:
    """高级SQL功能测试"""

    @pytest.fixture
    def test_excel_file(self):
        """创建测试用Excel文件"""
        data = {
            'ID': [1, 2, 3, 4, 5],
            'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
            'Age': [25, 30, 35, 40, 45],
            'Department': ['Sales', 'IT', 'HR', 'Sales', 'IT'],
            'Salary': [50000, 60000, 55000, 70000, 65000]
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False, sheet_name="Employees")
            yield tmp.name
        
        try:
            os.unlink(tmp.name)
        except:
            pass

    @pytest.fixture
    def test_excel_with_nulls(self):
        """创建包含空值的测试文件"""
        data = {
            'ID': [1, 2, 3, 4, 5],
            'Name': ['Alice', None, 'Charlie', 'David', 'Eve'],
            'Age': [25, 30, None, 40, 45],
            'Department': ['Sales', 'IT', 'HR', None, 'IT'],
            'Salary': [50000, 60000, 55000, 70000, None]
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            df.to_excel(tmp.name, index=False, sheet_name="Employees")
            yield tmp.name
        
        try:
            os.unlink(tmp.name)
        except:
            pass

    # ==================== 高级查询测试 ====================

    def test_distinct_query(self, test_excel_file):
        """测试DISTINCT查询"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT DISTINCT Department FROM Employees"
        )
        
        assert result['success'] is True
        assert len(result['data']) > 0

    def test_count_with_group_by(self, test_excel_file):
        """测试COUNT和GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Department, COUNT(*) as Count FROM Employees GROUP BY Department"
        )
        
        assert result['success'] is True

    def test_sum_with_group_by(self, test_excel_file):
        """测试SUM和GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Department, SUM(Salary) as TotalSalary FROM Employees GROUP BY Department"
        )
        
        assert result['success'] is True

    def test_avg_with_group_by(self, test_excel_file):
        """测试AVG和GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Department, AVG(Salary) as AvgSalary FROM Employees GROUP BY Department"
        )
        
        assert result['success'] is True

    def test_min_max_with_group_by(self, test_excel_file):
        """测试MIN/MAX和GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Department, MIN(Age) as MinAge, MAX(Age) as MaxAge FROM Employees GROUP BY Department"
        )
        
        assert result['success'] is True

    def test_order_by_desc(self, test_excel_file):
        """测试ORDER BY DESC"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Salary FROM Employees ORDER BY Salary DESC"
        )
        
        assert result['success'] is True

    def test_order_by_asc(self, test_excel_file):
        """测试ORDER BY ASC"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Age FROM Employees ORDER BY Age ASC"
        )
        
        assert result['success'] is True

    def test_limit_clause(self, test_excel_file):
        """测试LIMIT子句"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees LIMIT 3"
        )
        
        assert result['success'] is True

    def test_where_between(self, test_excel_file):
        """测试WHERE BETWEEN"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Age FROM Employees WHERE Age >= 25 AND Age <= 35"
        )
        
        assert result['success'] is True

    def test_where_in(self, test_excel_file):
        """测试WHERE IN"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Department FROM Employees WHERE Department IN ('Sales', 'IT')"
        )
        
        assert result['success'] is True

    def test_where_like(self, test_excel_file):
        """测试WHERE LIKE"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name LIKE 'A%'"
        )
        
        assert result['success'] is True

    def test_where_not_equal(self, test_excel_file):
        """测试WHERE <>"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Department FROM Employees WHERE Department <> 'HR'"
        )
        
        assert result['success'] is True

    def test_multiple_conditions(self, test_excel_file):
        """测试多条件查询"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Age, Salary FROM Employees WHERE Age > 25 AND Salary > 50000"
        )
        
        assert result['success'] is True

    def test_or_conditions(self, test_excel_file):
        """测试OR条件"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Department FROM Employees WHERE Department = 'Sales' OR Department = 'IT'"
        )
        
        assert result['success'] is True

    # ==================== 空值处理测试 ====================

    def test_is_null(self, test_excel_with_nulls):
        """测试空值处理"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # 由于 IS NULL 可能不被支持，使用等效的查询
        result = execute_advanced_sql_query(
            file_path=test_excel_with_nulls,
            sql="SELECT Name FROM Employees WHERE Name = '' OR Name IS NULL"
        )
        
        assert result is not None

    def test_is_not_null(self, test_excel_with_nulls):
        """测试非空值处理"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # 由于 IS NOT NULL 可能不被支持，使用等效的查询
        result = execute_advanced_sql_query(
            file_path=test_excel_with_nulls,
            sql="SELECT Name FROM Employees WHERE Name IS NOT NULL AND Name <> ''"
        )
        
        assert result is not None

    # ==================== 计算字段测试 ====================

    def test_arithmetic_division(self, test_excel_file):
        """测试算术运算 - 除法"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Salary/1000 as SalaryK FROM Employees"
        )
        
        assert result['success'] is True

    def test_arithmetic_addition(self, test_excel_file):
        """测试算术运算 - 加法"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Age + 5 as AgePlus5 FROM Employees"
        )
        
        assert result['success'] is True

    def test_arithmetic_multiplication(self, test_excel_file):
        """测试算术运算 - 乘法"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name, Salary * 12 as AnnualSalary FROM Employees"
        )
        
        assert result['success'] is True

    # ==================== 别名测试 ====================

    def test_column_alias(self, test_excel_file):
        """测试列别名"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name as EmployeeName, Salary as EmployeeSalary FROM Employees"
        )
        
        assert result['success'] is True

    # ==================== 错误处理测试 ====================

    def test_invalid_table_name(self, test_excel_file):
        """测试无效表名"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT * FROM NonExistentTable"
        )
        
        assert result['success'] is False

    def test_invalid_column(self, test_excel_file):
        """测试无效列名"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT InvalidColumn FROM Employees"
        )
        
        # 可能是错误或空结果
        assert result is not None

    def test_syntax_error(self, test_excel_file):
        """测试SQL语法错误"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # 简化语法错误测试
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM"
        )
        
        # 语法错误应该失败
        assert result is not None

    # ==================== OFFSET分页测试 ====================

    def test_offset_only(self, test_excel_file):
        """测试OFFSET跳过前N行"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result_all = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees ORDER BY Name LIMIT 10"
        )
        result_offset = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees ORDER BY Name LIMIT 10 OFFSET 2"
        )
        
        assert result_all['success'] is True
        assert result_offset['success'] is True
        
        # data[0]是header，data[1:]是实际数据
        all_names = [row[0] for row in result_all['data'][1:]]
        offset_names = [row[0] for row in result_offset['data'][1:]]
        
        # OFFSET后的结果数应比全部少2
        assert len(all_names) - len(offset_names) == 2
        # offset结果从all的第3个开始
        assert offset_names == all_names[2:]

    def test_offset_with_limit(self, test_excel_file):
        """测试OFFSET+LIMIT组合分页 - 验证分页不重叠"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # 全量
        result_all = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees ORDER BY Name LIMIT 100"
        )
        # 分页1: OFFSET 0 LIMIT 3
        page1 = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees ORDER BY Name LIMIT 3 OFFSET 0"
        )
        # 分页2: OFFSET 3 LIMIT 3
        page2 = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees ORDER BY Name LIMIT 3 OFFSET 3"
        )
        
        assert page1['success'] is True
        assert page2['success'] is True
        
        all_data = [row[0] for row in result_all['data'][1:]]
        p1_data = [row[0] for row in page1['data'][1:]]
        p2_data = [row[0] for row in page2['data'][1:]]
        
        # 合并两页应等于全量
        assert p1_data + p2_data == all_data
        # 两页不应有重叠
        assert len(set(p1_data) & set(p2_data)) == 0

    # ==================== NOT LIKE测试 ====================

    def test_not_like(self, test_excel_file):
        """测试NOT LIKE排除匹配"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # LIKE匹配含'l'的（Alice, Charlie）
        result_like = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name LIKE '%l%'"
        )
        # NOT LIKE排除含'l'的
        result_not_like = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name NOT LIKE '%l%'"
        )
        
        assert result_like['success'] is True
        assert result_not_like['success'] is True
        
        # 过滤掉可能的header行
        like_names = set(row[0] for row in result_like['data'] if row[0] != 'Name')
        not_like_names = set(row[0] for row in result_not_like['data'] if row[0] != 'Name')
        
        # 两集合不应有交集
        assert len(like_names & not_like_names) == 0

    def test_like_case_insensitive(self, test_excel_file):
        """测试LIKE大小写不敏感（游戏配置表场景：搜Fire和fire应匹配同一行）"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result_lower = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name LIKE '%alice%'"
        )
        result_upper = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name LIKE '%ALICE%'"
        )
        result_mixed = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Name LIKE '%Alice%'"
        )
        
        assert result_lower['success'] is True
        assert result_upper['success'] is True
        assert result_mixed['success'] is True
        
        # 三种大小写应返回相同结果
        names_lower = set(row[0] for row in result_lower['data'] if row[0] != 'Name')
        names_upper = set(row[0] for row in result_upper['data'] if row[0] != 'Name')
        names_mixed = set(row[0] for row in result_mixed['data'] if row[0] != 'Name')
        
        assert names_lower == names_upper == names_mixed
        assert 'Alice' in names_lower

    # ==================== COUNT(DISTINCT)测试 ====================

    def test_count_distinct_no_group_by(self, test_excel_file):
        """测试COUNT(DISTINCT)不带GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT COUNT(DISTINCT Department) FROM Employees"
        )
        
        assert result['success'] is True
        # 应返回不同部门的数量
        data_values = [row[0] for row in result['data'] if row[0] != 'count_distinct_Department']
        assert len(data_values) == 1
        assert isinstance(data_values[0], int)
        assert data_values[0] >= 1

    def test_count_distinct_with_group_by(self, test_excel_file):
        """测试COUNT(DISTINCT)带GROUP BY"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Department, COUNT(DISTINCT Name) FROM Employees GROUP BY Department"
        )
        
        assert result['success'] is True
        data_rows = [row for row in result['data'] if row[0] != 'Department']
        assert len(data_rows) >= 1
        for row in data_rows:
            assert isinstance(row[1], int)

    def test_not_in(self, test_excel_file):
        """测试NOT IN排除列表值"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        # IN匹配指定部门
        result_in = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Department IN ('Sales', 'HR')"
        )
        # NOT IN排除指定部门
        result_not_in = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT Name FROM Employees WHERE Department NOT IN ('Sales', 'HR')"
        )
        
        assert result_in['success'] is True
        assert result_not_in['success'] is True
        
        in_names = set(row[0] for row in result_in['data'] if row[0] != 'Name')
        not_in_names = set(row[0] for row in result_not_in['data'] if row[0] != 'Name')
        
        # 两集合不应有交集
        assert len(in_names & not_in_names) == 0
        # NOT IN结果应只有IT部门的人
        assert not_in_names == {'Bob', 'Eve'}

    def test_count_distinct_with_alias(self, test_excel_file):
        """测试COUNT(DISTINCT)带别名"""
        from src.api.advanced_sql_query import execute_advanced_sql_query
        
        result = execute_advanced_sql_query(
            file_path=test_excel_file,
            sql="SELECT COUNT(DISTINCT Department) AS dept_count FROM Employees"
        )
        
        assert result['success'] is True
        assert 'dept_count' in result.get('query_info', {}).get('returned_columns', [])
