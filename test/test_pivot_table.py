#!/usr/bin/env python3
"""测试create_pivot_table函数"""

import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock
from excel_mcp_server_fastmcp.server import excel_create_pivot_table


class TestPivotTable(unittest.TestCase):
    """测试透视表创建功能"""

    def setUp(self):
        """测试前准备"""
        self.test_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.test_file.close()
        
        # 创建一个简单的测试Excel文件
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"
        
        # 添加测试数据
        headers = ["类别", "子类别", "月份", "销售额", "利润"]
        ws.append(headers)
        
        data = [
            ["电子产品", "手机", "1月", 1000, 200],
            ["电子产品", "电脑", "1月", 2000, 400],
            ["电子产品", "手机", "2月", 1200, 240],
            ["电子产品", "电脑", "2月", 2200, 440],
            ["服装", "上衣", "1月", 800, 160],
            ["服装", "裤子", "1月", 600, 120],
            ["服装", "上衣", "2月", 900, 180],
            ["服装", "裤子", "2月", 700, 140]
        ]
        
        for row in data:
            ws.append(row)
        
        wb.save(self.test_file.name)
        self.wb = wb

    def tearDown(self):
        """测试后清理"""
        if os.path.exists(self.test_file.name):
            os.unlink(self.test_file.name)

    def test_create_pivot_table_basic(self):
        """测试基本透视表创建"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            agg_func="sum"
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['rows'], ["类别"])
        self.assertEqual(result['data']['values'], ["销售额"])
        self.assertEqual(result['data']['agg_func'], "sum")

    def test_create_pivot_table_with_mean_alias(self):
        """测试mean作为average的别名"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet", 
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            agg_func="mean"  # 测试mean别名
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['agg_func'], "mean")

    def test_create_pivot_table_average(self):
        """测试average函数"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9", 
            rows=["类别"],
            values=["销售额"],
            agg_func="average"
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['agg_func'], "average")

    def test_create_pivot_table_multiple_rows_values(self):
        """测试多行多值透视表"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9",
            rows=["类别", "子类别"],
            values=["销售额", "利润"],
            agg_func="sum"
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['rows'], ["类别", "子类别"])
        self.assertEqual(result['data']['values'], ["销售额", "利润"])

    def test_create_pivot_table_with_columns(self):
        """测试包含列字段的透视表"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            columns=["月份"],
            agg_func="sum"
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['rows'], ["类别"])
        self.assertEqual(result['data']['values'], ["销售额"])
        self.assertEqual(result['data']['columns'], ["月份"])

    def test_create_pivot_table_custom_sheet_name(self):
        """测试自定义透视表工作表名称"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            agg_func="sum",
            pivot_sheet_name="自定义透视表"
        )
        
        self.assertTrue(result['success'])
        self.assertEqual(result['data']['sheet_name'], "自定义透视表")

    def test_invalid_agg_func(self):
        """测试无效的聚合函数"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="TestSheet",
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            agg_func="invalid_func"
        )
        
        self.assertFalse(result['success'])
        self.assertIn("不支持的聚合函数", result['message'])

    def test_nonexistent_sheet(self):
        """测试不存在的工作表"""
        result = excel_create_pivot_table(
            file_path=self.test_file.name,
            sheet_name="NonExistentSheet",
            data_range="A1:E9",
            rows=["类别"],
            values=["销售额"],
            agg_func="sum"
        )
        
        self.assertFalse(result['success'])
        self.assertIn("数据工作表不存在", result['message'])

    def test_all_supported_agg_funcs(self):
        """测试所有支持的聚合函数"""
        supported_funcs = ["sum", "count", "average", "mean", "max", "min", "std", "var"]
        
        for func in supported_funcs:
            result = excel_create_pivot_table(
                file_path=self.test_file.name,
                sheet_name="TestSheet",
                data_range="A1:E9",
                rows=["类别"],
                values=["销售额"],
                agg_func=func
            )
            
            self.assertTrue(result['success'], f"函数 {func} 测试失败")


if __name__ == '__main__':
    unittest.main()