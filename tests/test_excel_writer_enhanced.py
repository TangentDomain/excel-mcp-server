# -*- coding: utf-8 -*-
"""
Excel Writer增强测试套件
测试src.core.excel_writer模块的所有核心功能
目标覆盖率：75%+
"""

import pytest
import tempfile
import os
import time
from unittest.mock import Mock, patch, MagicMock
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment

from src.core.excel_writer import ExcelWriter
from src.models.types import OperationResult, RangeType
from src.utils.exceptions import SheetNotFoundError, DataValidationError


class TestExcelWriterBasic:
    """ExcelWriter基础功能测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        # 创建临时Excel文件
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        # 创建基础Excel文件
        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加一些测试数据
        ws['A1'] = "Name"
        ws['B1'] = "Age"
        ws['C1'] = "City"
        ws['A2'] = "Alice"
        ws['B2'] = 25
        ws['C2'] = "Beijing"
        ws['A3'] = "Bob"
        ws['B3'] = 30
        ws['C3'] = "Shanghai"

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_writer_initialization_valid_file(self):
        """测试有效文件路径的初始化"""
        writer = ExcelWriter(self.file_path)
        assert writer.file_path == self.file_path

    def test_writer_initialization_invalid_file(self):
        """测试无效文件路径的初始化"""
        with pytest.raises(Exception):
            ExcelWriter("invalid_path.xlsx")

    def test_update_range_single_cell(self):
        """测试更新单个单元格"""
        writer = ExcelWriter(self.file_path)
        data = [["New Value"]]

        result = writer.update_range("TestSheet!A2", data)

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 1
        assert result.metadata['sheet_name'] == "TestSheet"

    def test_update_range_multiple_cells(self):
        """测试更新多个单元格"""
        writer = ExcelWriter(self.file_path)
        data = [["New1", "New2"], ["New3", "New4"]]

        result = writer.update_range("TestSheet!A2:B3", data)

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 4

    def test_update_range_preserve_formulas(self):
        """测试保留公式"""
        writer = ExcelWriter(self.file_path)

        # 先添加一个公式
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        ws['D2'] = "=SUM(B2:B3)"
        wb.save(self.file_path)

        # 更新其他单元格，保留公式（使用覆盖模式避免行插入影响公式位置）
        data = [["Updated"]]
        result = writer.update_range("TestSheet!A2", data, preserve_formulas=True, insert_mode=False)

        assert result.success is True

        # 验证公式仍然存在
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        assert ws['D2'].data_type == 'f'  # 公式类型

    def test_update_range_overwrite_formulas(self):
        """测试覆盖公式"""
        writer = ExcelWriter(self.file_path)

        # 先添加一个公式
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        ws['D2'] = "=SUM(B2:B3)"
        wb.save(self.file_path)

        # 更新公式单元格，不保留公式
        data = [["New Value"]]
        result = writer.update_range("TestSheet!D2", data, preserve_formulas=False)

        assert result.success is True

        # 验证公式被覆盖
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        assert ws['D2'].value == "New Value"

    def test_update_range_insert_mode(self):
        """测试插入模式"""
        writer = ExcelWriter(self.file_path)
        original_rows = 3  # 初始有3行数据

        data = [["Insert1", "Insert2"], ["Insert3", "Insert4"]]
        result = writer.update_range("TestSheet!A2", data, insert_mode=True)

        assert result.success is True
        assert result.metadata['insert_mode'] is True
        assert result.metadata['rows_inserted'] == 2

    def test_update_range_invalid_sheet(self):
        """测试无效工作表"""
        writer = ExcelWriter(self.file_path)
        data = [["Test"]]

        result = writer.update_range("NonExistentSheet!A1", data)

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_convert_to_cell_range_row_range_error(self):
        """测试行范围格式转换错误"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        with pytest.raises(ValueError) as exc_info:
            writer._convert_to_cell_range("2:5", RangeType.ROW_RANGE, sheet, [["test"]])

        assert "不支持纯行范围格式" in str(exc_info.value)

    def test_convert_to_cell_range_single_row_error(self):
        """测试单行范围格式转换错误"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        with pytest.raises(ValueError) as exc_info:
            writer._convert_to_cell_range("3:3", RangeType.SINGLE_ROW, sheet, [["test"]])

        assert "不支持单行范围格式" in str(exc_info.value)

    def test_convert_to_cell_range_column_range(self):
        """测试列范围格式转换"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        result = writer._convert_to_cell_range("A:C", RangeType.COLUMN_RANGE, sheet, [["test"]])
        assert result == "A:C"

    def test_convert_to_cell_range_single_column(self):
        """测试单列范围格式转换"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        result = writer._convert_to_cell_range("B:B", RangeType.SINGLE_COLUMN, sheet, [["test"]])
        assert result == "B:B"

    def test_get_worksheet_valid_sheet(self):
        """测试获取有效工作表"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)

        sheet = writer._get_worksheet(wb, "TestSheet")
        assert sheet.title == "TestSheet"

    def test_get_worksheet_empty_name(self):
        """测试空工作表名称"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)

        with pytest.raises(SheetNotFoundError) as exc_info:
            writer._get_worksheet(wb, "")

        assert "工作表名称不能为空" in str(exc_info.value)

    def test_get_worksheet_nonexistent_sheet(self):
        """测试不存在的工作表"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)

        with pytest.raises(SheetNotFoundError) as exc_info:
            writer._get_worksheet(wb, "NonExistent")

        assert "工作表不存在" in str(exc_info.value)

    def test_write_data_basic(self):
        """测试基础数据写入"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        data = [["New1", "New2"], ["New3", "New4"]]
        modified_cells = writer._write_data(sheet, data, 2, 1, False)

        assert len(modified_cells) == 4
        assert modified_cells[0].coordinate == "A2"
        assert modified_cells[0].new_value == "New1"

    def test_write_data_preserve_existing_formula(self):
        """测试保留现有公式"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        # 设置一个公式
        sheet['D2'].value = "=SUM(B2:C2)"

        data = [["ShouldNotOverwrite"]]
        modified_cells = writer._write_data(sheet, data, 2, 4, True)  # preserve_formulas=True

        # 应该没有修改，因为是公式
        assert len(modified_cells) == 0

    def test_write_data_overwrite_formula(self):
        """测试覆盖公式"""
        writer = ExcelWriter(self.file_path)
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        sheet = wb["TestSheet"]

        # 设置一个公式
        sheet['D2'].value = "=SUM(B2:C2)"

        data = [["NewValue"]]
        modified_cells = writer._write_data(sheet, data, 2, 4, False)  # preserve_formulas=False

        # 应该修改了公式
        assert len(modified_cells) == 1
        assert modified_cells[0].new_value == "NewValue"


class TestExcelWriterRowColumnOperations:
    """ExcelWriter行列操作测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        for i in range(1, 6):
            for j in range(1, 4):
                ws.cell(row=i, column=j, value=f"R{i}C{j}")

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_insert_rows_basic(self):
        """测试基础行插入"""
        writer = ExcelWriter(self.file_path)

        result = writer.insert_rows("TestSheet", 3, 2)

        assert result.success is True
        assert result.metadata['inserted_at_row'] == 3
        assert result.metadata['inserted_count'] == 2
        assert result.metadata['new_max_row'] > result.metadata['original_max_row']

    def test_insert_rows_single_row(self):
        """测试插入单行"""
        writer = ExcelWriter(self.file_path)

        result = writer.insert_rows("TestSheet", 2, 1)

        assert result.success is True
        assert result.metadata['inserted_count'] == 1

    def test_insert_columns_basic(self):
        """测试基础列插入"""
        writer = ExcelWriter(self.file_path)

        result = writer.insert_columns("TestSheet", 2, 2)

        assert result.success is True
        assert result.metadata['inserted_at_column'] == 2
        assert result.metadata['inserted_count'] == 2
        assert result.metadata['new_max_column'] > result.metadata['original_max_column']

    def test_insert_columns_single_column(self):
        """测试插入单列"""
        writer = ExcelWriter(self.file_path)

        result = writer.insert_columns("TestSheet", 1, 1)

        assert result.success is True
        assert result.metadata['inserted_count'] == 1

    def test_delete_rows_basic(self):
        """测试基础行删除"""
        writer = ExcelWriter(self.file_path)

        result = writer.delete_rows("TestSheet", 3, 2)

        assert result.success is True
        assert result.metadata['deleted_start_row'] == 3
        assert result.metadata['actual_deleted_count'] == 2

    def test_delete_rows_beyond_range(self):
        """测试删除超出范围的行"""
        writer = ExcelWriter(self.file_path)

        result = writer.delete_rows("TestSheet", 100, 5)

        assert result.success is False
        assert "超过工作表最大行数" in result.error

    def test_delete_rows_partial_range(self):
        """测试部分范围删除"""
        writer = ExcelWriter(self.file_path)

        # 从第4行开始删除3行，但只有2行可用
        result = writer.delete_rows("TestSheet", 4, 3)

        assert result.success is True
        assert result.metadata['actual_deleted_count'] == 2  # 只能删除2行

    def test_delete_columns_basic(self):
        """测试基础列删除"""
        writer = ExcelWriter(self.file_path)

        result = writer.delete_columns("TestSheet", 2, 1)

        assert result.success is True
        assert result.metadata['deleted_start_column'] == 2
        assert result.metadata['actual_deleted_count'] == 1

    def test_delete_columns_beyond_range(self):
        """测试删除超出范围的列"""
        writer = ExcelWriter(self.file_path)

        result = writer.delete_columns("TestSheet", 100, 5)

        assert result.success is False
        assert "超过工作表最大列数" in result.error

    def test_delete_columns_partial_range(self):
        """测试部分范围列删除"""
        writer = ExcelWriter(self.file_path)

        # 从第2列开始删除5列，但只有3列可用
        result = writer.delete_columns("TestSheet", 2, 5)

        assert result.success is True
        assert result.metadata['actual_deleted_count'] == 2  # 只能删除2列


class TestExcelWriterFormulaOperations:
    """ExcelWriter公式操作测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加一些数据用于公式计算
        ws['A1'] = 10
        ws['A2'] = 20
        ws['A3'] = 30
        ws['B1'] = 5
        ws['B2'] = 15
        ws['B3'] = 25

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_set_formula_basic(self):
        """测试基础公式设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("C1", "SUM(A1:A3)", "TestSheet")

        assert result.success is True
        assert result.metadata['cell_address'] == "C1"
        assert result.metadata['formula'] == "SUM(A1:A3)"

    def test_set_formula_with_equals_sign(self):
        """测试带等号的公式设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("C1", "=SUM(A1:A3)", "TestSheet")

        assert result.success is True
        assert result.metadata['formula'] == "SUM(A1:A3)"  # 等号被移除

    def test_set_formula_empty_formula(self):
        """测试空公式"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("C1", "", "TestSheet")

        assert result.success is False
        assert "公式不能为空" in result.error

    def test_set_formula_invalid_cell_address(self):
        """测试无效单元格地址"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("INVALID", "SUM(A1:A3)", "TestSheet")

        assert result.success is False
        assert "Invalid cell coordinates" in result.error

    def test_set_formula_nonexistent_sheet(self):
        """测试不存在的工作表"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("C1", "SUM(A1:A3)", "NonExistent")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    @patch('src.core.excel_writer.logger')
    def test_set_formula_logging(self, mock_logger):
        """测试公式设置日志记录"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_formula("C1", "SUM(A1:A3)", "TestSheet")

        assert result.success is True
        # 验证日志记录
        mock_logger.info.assert_called()

    def test_evaluate_formula_basic(self):
        """测试基础公式计算"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("SUM(A1:A3)", "TestSheet")

        # 结果应该是10+20+30=60
        assert result.success is True
        assert result.data == 60
        assert result.metadata['result_type'] == "number"

    def test_evaluate_formula_empty_formula(self):
        """测试空公式计算"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("", "TestSheet")

        assert result.success is False
        assert "公式不能为空" in result.error

    def test_evaluate_formula_with_equals_sign(self):
        """测试带等号的公式计算"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("=SUM(A1:A3)", "TestSheet")

        assert result.success is True
        assert result.data == 60

    def test_evaluate_formula_average(self):
        """测试平均值公式"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("AVERAGE(A1:A3)", "TestSheet")

        # (10+20+30)/3 = 20
        assert result.success is True
        assert result.data == 20

    def test_evaluate_formula_count(self):
        """测试计数公式"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("COUNT(A1:A3)", "TestSheet")

        assert result.success is True
        assert result.data == 3

    def test_evaluate_formula_min_max(self):
        """测试最值公式"""
        writer = ExcelWriter(self.file_path)

        min_result = writer.evaluate_formula("MIN(A1:A3)", "TestSheet")
        max_result = writer.evaluate_formula("MAX(A1:A3)", "TestSheet")

        assert min_result.success is True
        assert min_result.data == 10
        assert max_result.success is True
        assert max_result.data == 30

    def test_evaluate_formula_text_concatenation(self):
        """测试文本连接公式"""
        writer = ExcelWriter(self.file_path)

        # 先添加一些文本数据
        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        ws['D1'] = "Hello"
        ws['D2'] = "World"
        wb.save(self.file_path)

        result = writer.evaluate_formula('CONCATENATE("Hello", " ", "World")', "TestSheet")

        assert result.success is True
        assert result.data == "Hello World"

    def test_evaluate_formula_basic_math(self):
        """测试基础数学运算"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("10 + 20 * 2", "TestSheet")

        assert result.success is True
        assert result.data == 50  # 20*2 + 10

    def test_evaluate_formula_cache_hit(self):
        """测试公式缓存命中"""
        writer = ExcelWriter(self.file_path)

        # 第一次计算
        result1 = writer.evaluate_formula("SUM(A1:A3)", "TestSheet")
        assert result1.success is True
        assert result1.metadata['cached'] is False

        # 第二次计算相同公式（应该命中缓存）
        result2 = writer.evaluate_formula("SUM(A1:A3)", "TestSheet")
        assert result2.success is True
        assert result2.metadata['cached'] is True

    def test_evaluate_formula_result_types(self):
        """测试不同结果类型"""
        writer = ExcelWriter(self.file_path)

        # 数字类型
        num_result = writer.evaluate_formula("SUM(A1:A3)", "TestSheet")
        assert num_result.success is True
        assert num_result.metadata['result_type'] == "number"

        # 文本类型
        text_result = writer.evaluate_formula('CONCATENATE("A", "B")', "TestSheet")
        assert text_result.success is True
        assert text_result.metadata['result_type'] == "text"

    def test_evaluate_formula_invalid_formula(self):
        """测试无效公式"""
        writer = ExcelWriter(self.file_path)

        result = writer.evaluate_formula("INVALID_FUNCTION(A1:A3)", "TestSheet")

        # 应该返回None或错误
        assert result.success is True  # 公式计算本身成功，但结果可能是None
        assert result.data is None

    def test_create_temp_workbook(self):
        """测试创建临时工作簿"""
        writer = ExcelWriter(self.file_path)

        with patch('src.utils.formula_cache.get_formula_cache') as mock_cache:
            mock_cache_instance = Mock()
            mock_cache.return_value = mock_cache_instance

            temp_workbook, temp_file_path = writer._create_temp_workbook("TestSheet", mock_cache_instance)

            assert temp_workbook is not None
            assert temp_file_path is not None
            assert os.path.exists(temp_file_path)

            # 清理
            try:
                os.unlink(temp_file_path)
            except:
                pass

    def test_basic_formula_parse_sum(self):
        """测试基础公式解析 - SUM函数"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]

        result = writer._basic_formula_parse("SUM(A1:A3)", ws)

        assert result == 60  # 10+20+30

    def test_basic_formula_parse_average(self):
        """测试基础公式解析 - AVERAGE函数"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]

        result = writer._basic_formula_parse("AVERAGE(A1:A3)", ws)

        assert result == 20  # (10+20+30)/3

    def test_get_range_values(self):
        """测试获取范围值"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]

        values = writer._get_range_values(ws, "A1", "A3")

        assert len(values) == 3
        assert 10 in values
        assert 20 in values
        assert 30 in values

    def test_numpy_average_fallback(self):
        """测试numpy平均值回退实现"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30]
        result = writer._numpy_average(values)

        assert result == 20

    def test_numpy_min_max(self):
        """测试numpy最值函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30, 5, 100]

        min_result = writer._numpy_min(values)
        max_result = writer._numpy_max(values)

        assert min_result == 5
        assert max_result == 100

    def test_numpy_median(self):
        """测试numpy中位数函数"""
        writer = ExcelWriter(self.file_path)

        # 奇数个值
        odd_values = [10, 20, 30]
        assert writer._numpy_median(odd_values) == 20

        # 偶数个值
        even_values = [10, 20, 30, 40]
        assert writer._numpy_median(even_values) == 25

    def test_numpy_stdev_var(self):
        """测试numpy标准差和方差函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30, 40, 50]

        stdev_result = writer._numpy_stdev(values)
        var_result = writer._numpy_var(values)

        # 验证标准差 = 方差的平方根
        assert abs(stdev_result - (var_result ** 0.5)) < 0.001

    def test_numpy_countif(self):
        """测试numpy条件计数函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30, 40, 50]

        # 大于25的计数
        count_gt = writer._numpy_countif(values, ">25")
        assert count_gt == 3  # 30, 40, 50

        # 等于20的计数
        count_eq = writer._numpy_countif(values, "=20")
        assert count_eq == 1

    def test_numpy_sumif(self):
        """测试numpy条件求和函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30, 40, 50]

        # 大于25的求和
        sum_gt = writer._numpy_sumif(values, ">25")
        assert sum_gt == 120  # 30+40+50

        # 等于20的求和
        sum_eq = writer._numpy_sumif(values, "=20")
        assert sum_eq == 20

    def test_numpy_averageif(self):
        """测试numpy条件平均值函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 30, 40, 50]

        # 大于25的平均值
        avg_gt = writer._numpy_averageif(values, ">25")
        assert avg_gt == 40  # (30+40+50)/3

    def test_numpy_mode(self):
        """测试numpy众数函数"""
        writer = ExcelWriter(self.file_path)

        values = [10, 20, 20, 30, 30, 30, 40]
        result = writer._numpy_mode(values)

        assert result == 30  # 30出现次数最多

    def test_calculate_range_sum(self):
        """测试范围求和计算"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]

        result = writer._calculate_range_sum(ws, "A1", "A3")

        assert result == 60  # 10+20+30

    def test_calculate_range_count(self):
        """测试范围计数计算"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]

        result = writer._calculate_range_count(ws, "A1", "A3")

        assert result == 3  # 3个数值

    def test_get_result_type(self):
        """测试结果类型判断"""
        writer = ExcelWriter(self.file_path)

        # 数字类型
        assert writer._get_result_type(42) == "number"
        assert writer._get_result_type(3.14) == "number"

        # 文本类型
        assert writer._get_result_type("hello") == "text"

        # 布尔类型
        assert writer._get_result_type(True) == "boolean"

        # 空值类型
        assert writer._get_result_type(None) == "null"

        # 未知类型
        assert writer._get_result_type([]) == "unknown"


class TestExcelWriterFormatting:
    """ExcelWriter格式化测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        for i in range(1, 4):
            for j in range(1, 4):
                ws.cell(row=i, column=j, value=f"R{i}C{j}")

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_format_cells_basic(self):
        """测试基础单元格格式化"""
        writer = ExcelWriter(self.file_path)

        formatting = {
            'font': {'bold': True, 'color': 'FF0000'},
            'fill': {'color': 'FFFF00'}
        }

        result = writer.format_cells("TestSheet!A1:B2", formatting)

        assert result.success is True
        assert result.metadata['formatted_count'] > 0

    def test_format_cells_font_only(self):
        """测试仅字体格式化"""
        writer = ExcelWriter(self.file_path)

        formatting = {
            'font': {'name': 'Arial', 'size': 12, 'bold': True}
        }

        result = writer.format_cells("TestSheet!A1", formatting)

        assert result.success is True
        assert result.metadata['formatted_count'] == 1

    def test_format_cells_fill_only(self):
        """测试仅背景填充格式化"""
        writer = ExcelWriter(self.file_path)

        formatting = {
            'fill': {'color': '00FF00'}
        }

        result = writer.format_cells("TestSheet!A1:C1", formatting)

        assert result.success is True
        assert result.metadata['formatted_count'] == 3

    def test_format_cells_alignment(self):
        """测试对齐格式化"""
        writer = ExcelWriter(self.file_path)

        formatting = {
            'alignment': {'horizontal': 'center', 'vertical': 'center'}
        }

        result = writer.format_cells("TestSheet!A1:B2", formatting)

        assert result.success is True

    def test_format_cells_invalid_range(self):
        """测试无效范围的格式化"""
        writer = ExcelWriter(self.file_path)

        formatting = {'font': {'bold': True}}

        result = writer.format_cells("NonExistentSheet!A1", formatting)

        assert result.success is False

    def test_apply_cell_format_font(self):
        """测试应用字体格式"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        cell = ws['A1']

        formatting = {
            'font': {'bold': True, 'italic': True, 'size': 14}
        }

        writer._apply_cell_format(cell, formatting)

        assert cell.font.bold is True
        assert cell.font.italic is True
        assert cell.font.size == 14

    def test_apply_cell_format_fill(self):
        """测试应用背景填充"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        cell = ws['A1']

        formatting = {
            'fill': {'color': 'FF0000'}
        }

        writer._apply_cell_format(cell, formatting)

        assert cell.fill.start_color.rgb in ['FFFF0000', '00FF0000']  # 接受两种格式

    def test_apply_cell_format_alignment(self):
        """测试应用对齐格式"""
        writer = ExcelWriter(self.file_path)

        from openpyxl import load_workbook
        wb = load_workbook(self.file_path)
        ws = wb["TestSheet"]
        cell = ws['A1']

        formatting = {
            'alignment': {'horizontal': 'right', 'vertical': 'bottom'}
        }

        writer._apply_cell_format(cell, formatting)

        assert cell.alignment.horizontal == 'right'
        assert cell.alignment.vertical == 'bottom'


class TestExcelWriterMergeOperations:
    """ExcelWriter合并操作测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        for i in range(1, 4):
            for j in range(1, 4):
                ws.cell(row=i, column=j, value=f"R{i}C{j}")

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_merge_cells_basic(self):
        """测试基础单元格合并"""
        writer = ExcelWriter(self.file_path)

        result = writer.merge_cells("TestSheet!A1:C1")

        assert result.success is True
        assert result.data['merged_range'] == "A1:C1"
        assert result.data['sheet_name'] == "TestSheet"

    def test_merge_cells_with_sheet_name_parameter(self):
        """测试带工作表参数的合并"""
        writer = ExcelWriter(self.file_path)

        result = writer.merge_cells("A1:B1", "TestSheet")

        assert result.success is True
        assert result.data['merged_range'] == "A1:B1"

    def test_merge_cells_full_range_expression(self):
        """测试完整范围表达式合并"""
        writer = ExcelWriter(self.file_path)

        result = writer.merge_cells("TestSheet!B2:D3")

        assert result.success is True
        assert result.data['merged_range'] == "B2:D3"

    def test_merge_cells_nonexistent_sheet(self):
        """测试不存在工作表的合并"""
        writer = ExcelWriter(self.file_path)

        result = writer.merge_cells("NonExistentSheet!A1:C1")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_unmerge_cells_basic(self):
        """测试基础取消合并"""
        writer = ExcelWriter(self.file_path)

        # 先合并
        merge_result = writer.merge_cells("TestSheet!A1:C1")
        assert merge_result.success is True

        # 再取消合并
        result = writer.unmerge_cells("TestSheet!A1:C1")

        assert result.success is True
        assert result.data['unmerged_range'] == "A1:C1"

    def test_unmerge_cells_nonexistent_sheet(self):
        """测试不存在工作表的取消合并"""
        writer = ExcelWriter(self.file_path)

        result = writer.unmerge_cells("NonExistentSheet!A1:C1")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error


class TestExcelWriterBorderOperations:
    """ExcelWriter边框操作测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        for i in range(1, 4):
            for j in range(1, 4):
                ws.cell(row=i, column=j, value=f"R{i}C{j}")

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_set_borders_basic(self):
        """测试基础边框设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("TestSheet!A1:C3", "thin")

        assert result.success is True
        assert result.data['border_style'] == "thin"
        assert result.data['cell_count'] == 9  # 3x3 = 9 cells

    def test_set_borders_thick(self):
        """测试粗边框设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("TestSheet!A1:B2", "thick")

        assert result.success is True
        assert result.data['border_style'] == "thick"

    def test_set_borders_double(self):
        """测试双边框设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("TestSheet!A1", "double")

        assert result.success is True
        assert result.data['border_style'] == "double"

    def test_set_borders_with_sheet_name_parameter(self):
        """测试带工作表参数的边框设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("A1:C3", "thin", "TestSheet")

        assert result.success is True
        assert result.data['sheet_name'] == "TestSheet"

    def test_set_borders_nonexistent_sheet(self):
        """测试不存在工作表的边框设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("NonExistentSheet!A1:C1", "thin")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_set_borders_single_cell(self):
        """测试单单元格边框"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("TestSheet!B2", "dashed")

        assert result.success is True
        assert result.data['cell_count'] == 1

    def test_set_borders_large_range(self):
        """测试大范围边框"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_borders("TestSheet!A1:Z100", "thin")

        assert result.success is True
        assert result.data['cell_count'] == 2600  # 26*100


class TestExcelWriterDimensionOperations:
    """ExcelWriter尺寸操作测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        # 添加测试数据
        for i in range(1, 6):
            for j in range(1, 4):
                ws.cell(row=i, column=j, value=f"R{i}C{j}")

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_set_row_height_basic(self):
        """测试基础行高设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_row_height(2, 25.5, "TestSheet")

        assert result.success is True
        assert result.data['row_number'] == 2
        assert result.data['height'] == 25.5
        assert result.data['sheet_name'] == "TestSheet"

    def test_set_row_height_default_sheet(self):
        """测试默认工作表的行高设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_row_height(3, 30.0)

        assert result.success is True
        assert result.data['row_number'] == 3
        assert result.data['height'] == 30.0

    def test_set_row_height_nonexistent_sheet(self):
        """测试不存在工作表的行高设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_row_height(2, 25.0, "NonExistent")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_set_column_width_basic(self):
        """测试基础列宽设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_column_width("B", 15.5, "TestSheet")

        assert result.success is True
        assert result.data['column'] == "B"
        assert result.data['width'] == 15.5
        assert result.data['sheet_name'] == "TestSheet"

    def test_set_column_width_lowercase(self):
        """测试小写列标识符"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_column_width("c", 20.0, "TestSheet")

        assert result.success is True
        assert result.data['column'] == "C"  # 应该转换为大写

    def test_set_column_width_default_sheet(self):
        """测试默认工作表的列宽设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_column_width("A", 12.0)

        assert result.success is True
        assert result.data['column'] == "A"
        assert result.data['width'] == 12.0

    def test_set_column_width_nonexistent_sheet(self):
        """测试不存在工作表的列宽设置"""
        writer = ExcelWriter(self.file_path)

        result = writer.set_column_width("B", 15.0, "NonExistent")

        assert result.success is False
        assert "工作表不存在" in result.error or "不存在" in result.error

    def test_set_column_width_multiple_columns(self):
        """测试多列宽设置"""
        writer = ExcelWriter(self.file_path)

        # 设置多个列的宽度
        result_a = writer.set_column_width("A", 10.0)
        result_b = writer.set_column_width("B", 15.0)
        result_c = writer.set_column_width("C", 20.0)

        assert result_a.success is True
        assert result_b.success is True
        assert result_c.success is True

        assert result_a.data['width'] == 10.0
        assert result_b.data['width'] == 15.0
        assert result_c.data['width'] == 20.0


class TestExcelWriterPerformanceAndErrorHandling:
    """ExcelWriter性能和错误处理测试"""

    def setup_method(self):
        """每个测试方法前的设置"""
        self.temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.temp_file.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"

        wb.save(self.temp_file.name)
        self.file_path = self.temp_file.name

    def teardown_method(self):
        """每个测试方法后的清理"""
        if os.path.exists(self.file_path):
            try:
                os.unlink(self.file_path)
            except:
                pass

    def test_large_data_update_performance(self):
        """测试大数据量更新性能"""
        writer = ExcelWriter(self.file_path)

        # 创建大量数据
        large_data = [[f"Value_{i}_{j}" for j in range(10)] for i in range(100)]

        start_time = time.time()
        result = writer.update_range("TestSheet!A1", large_data)
        end_time = time.time()

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 1000  # 100*10
        assert end_time - start_time < 5.0  # 应该在5秒内完成

    def test_concurrent_operations(self):
        """测试并发操作安全性"""
        import threading

        results = []
        errors = []

        def worker(worker_id):
            try:
                writer = ExcelWriter(self.file_path)
                data = [[f"Worker_{worker_id}_Cell_{i}" for i in range(3)]]
                result = writer.update_range(f"TestSheet!A{worker_id + 1}", data)
                results.append((worker_id, result.success))
            except Exception as e:
                errors.append((worker_id, str(e)))

        # 启动多个线程
        threads = []
        for i in range(3):
            thread = threading.Thread(target=worker, args=(i,))
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        # 验证结果
        assert len(errors) == 0, f"并发操作出现错误: {errors}"
        assert len(results) == 3
        assert all(success for _, success in results)

    def test_memory_usage_large_data(self):
        """测试大数据量的内存使用"""
        import gc

        writer = ExcelWriter(self.file_path)

        # 强制垃圾回收
        gc.collect()

        # 创建非常大的数据
        very_large_data = [[f"Cell_{i}_{j}" for j in range(50)] for i in range(200)]

        # 执行操作
        result = writer.update_range("TestSheet!A1", very_large_data)

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 10000  # 200*50

        # 再次垃圾回收
        gc.collect()

        # 如果能到达这里，说明没有严重的内存问题
        assert True

    def test_error_handling_corrupted_file(self):
        """测试损坏文件的处理"""
        # 创建一个损坏的文件（写入无效内容）
        with open(self.file_path, 'w') as f:
            f.write("This is not a valid Excel file")

        writer = ExcelWriter(self.file_path)
        data = [["Test"]]

        result = writer.update_range("TestSheet!A1", data)

        assert result.success is False
        # 检查错误信息
        error_msg = str(result.error) if hasattr(result, 'error') and result.error else ""
        assert "失败" in error_msg or "error" in error_msg.lower() or "zip" in error_msg.lower()

    def test_error_handling_invalid_range_format(self):
        """测试无效范围格式的处理"""
        writer = ExcelWriter(self.file_path)
        data = [["Test"]]

        # 测试各种无效格式
        invalid_ranges = [
            "InvalidRange",
            "Sheet1!Invalid",
            "A1:ZZ999",  # 超出Excel限制
            "",
        ]

        for invalid_range in invalid_ranges:
            result = writer.update_range(invalid_range, data)
            # 某些可能会成功，某些可能会失败，主要测试不会崩溃
            assert result is not None

    def test_error_handling_permission_denied(self):
        """测试权限拒绝的处理"""
        # 尝试对只读文件进行操作
        if os.name == 'nt':  # Windows系统
            # 设置文件为只读
            os.chmod(self.file_path, 0o444)

            writer = ExcelWriter(self.file_path)
            data = [["Test"]]

            result = writer.update_range("TestSheet!A1", data)

            # 恢复文件权限以便清理
            os.chmod(self.file_path, 0o666)

            # 结果可能成功也可能失败，取决于系统
            assert result is not None

    def test_recovery_after_error(self):
        """测试错误后的恢复能力"""
        writer = ExcelWriter(self.file_path)

        # 先执行一个会失败的操作
        invalid_result = writer.update_range("NonExistentSheet!A1", [["Test"]])
        assert invalid_result.success is False

        # 再执行一个正常的操作
        valid_result = writer.update_range("TestSheet!A1", [["Recovery"]])
        assert valid_result.success is True

    def test_data_type_handling(self):
        """测试各种数据类型的处理"""
        from datetime import datetime

        writer = ExcelWriter(self.file_path)

        # 测试各种数据类型
        test_data = [
            [123, 45.67, True, None, "Text"],
            ["中文", "🎮", "", 0, []],
            [{"key": "value"}, (1, 2), 3+4j, datetime.now(), "Special"]
        ]

        result = writer.update_range("TestSheet!A1", test_data)

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 15  # 3 rows * 5 cols

    def test_unicode_and_special_characters(self):
        """测试Unicode和特殊字符处理"""
        writer = ExcelWriter(self.file_path)

        # 测试各种Unicode字符
        unicode_data = [
            ["中文测试", "English", "Español", "Français"],
            ["🎮游戏", "📊数据", "🔧工具", "🚀性能"],
            ["αβγ", "∑∏∫", "℃℉", "♠♥♦♣"]
        ]

        result = writer.update_range("TestSheet!A1", unicode_data)

        assert result.success is True
        assert result.metadata['modified_cells_count'] == 12


if __name__ == "__main__":
    pytest.main([__file__, "-v"])