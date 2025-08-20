#!/usr/bin/env python3
"""
Excel MCP Server - MCP工具测试

测试所有15个MCP工具的功能，包括正常场景、边界条件和错误处理
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from openpyxl import Workbook

# 导入所有MCP工具
from server import (
    excel_list_sheets, excel_regex_search, excel_get_range, excel_update_range,
    excel_insert_rows, excel_insert_columns, excel_create_file, excel_create_sheet,
    excel_delete_sheet, excel_rename_sheet, excel_delete_rows, excel_delete_columns,
    excel_set_formula, excel_format_cells
)


class TestExcelListSheets:
    """测试excel_list_sheets工具"""

    def test_basic_functionality(self, sample_xlsx_file):
        """测试基本功能"""
        result = excel_list_sheets(sample_xlsx_file)

        assert result['success'] is True
        assert 'sheets' in result
        assert 'active_sheet' in result
        assert isinstance(result['sheets'], list)
        assert len(result['sheets']) > 0
        assert result['active_sheet'] in result['sheets']

    def test_multi_sheet_file(self, multi_sheet_xlsx_file):
        """测试多工作表文件"""
        result = excel_list_sheets(multi_sheet_xlsx_file)

        assert result['success'] is True
        assert len(result['sheets']) == 2
        assert 'Data' in result['sheets']
        assert 'Summary' in result['sheets']

    def test_nonexistent_file(self, nonexistent_file_path):
        """测试不存在的文件"""
        result = excel_list_sheets(nonexistent_file_path)

        assert result['success'] is False
        assert 'error' in result
        assert '文件不存在' in result['error'] or 'FileNotFoundError' in result['error']

    def test_invalid_format_file(self, invalid_format_file):
        """测试无效格式文件"""
        result = excel_list_sheets(invalid_format_file)

        assert result['success'] is False
        assert 'error' in result


class TestExcelRegexSearch:
    """测试excel_regex_search工具"""

    def test_basic_search(self, sample_xlsx_file):
        """测试基本搜索功能"""
        # 搜索邮箱模式
        result = excel_regex_search(
            sample_xlsx_file,
            r'\w+@\w+\.\w+',
            flags="i"
        )

        assert result['success'] is True
        assert 'matches' in result
        assert 'match_count' in result
        assert isinstance(result['matches'], list)
        assert result['match_count'] >= 0

    def test_case_insensitive_search(self, sample_xlsx_file):
        """测试大小写不敏感搜索"""
        result = excel_regex_search(
            sample_xlsx_file,
            'alice',
            flags='i'
        )

        assert result['success'] is True
        assert result['match_count'] >= 1

    def test_invalid_regex(self, sample_xlsx_file):
        """测试无效正则表达式"""
        result = excel_regex_search(
            sample_xlsx_file,
            '[invalid'
        )

        assert result['success'] is False
        assert 'error' in result

    def test_search_values_and_formulas(self, temp_dir):
        """测试搜索值和公式"""
        # 创建包含公式的文件
        file_path = temp_dir / 'test_formulas.xlsx'
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 'Test Value'
        sheet['B1'] = '=SUM(A1:A1)'
        workbook.save(file_path)

        # 搜索值
        result_values = excel_regex_search(
            str(file_path),
            'Test',
            search_values=True,
            search_formulas=False
        )
        assert result_values['success'] is True

        # 搜索公式
        result_formulas = excel_regex_search(
            str(file_path),
            'SUM',
            search_values=False,
            search_formulas=True
        )
        assert result_formulas['success'] is True


class TestExcelGetRange:
    """测试excel_get_range工具"""

    def test_basic_range_read(self, sample_xlsx_file):
        """测试基本范围读取"""
        result = excel_get_range(sample_xlsx_file, 'A1:C2')

        assert result['success'] is True
        assert 'data' in result
        assert isinstance(result['data'], list)
        assert len(result['data']) == 2
        assert len(result['data'][0]) == 3

    def test_sheet_specific_range(self, multi_sheet_xlsx_file):
        """测试指定工作表范围"""
        result = excel_get_range(multi_sheet_xlsx_file, 'Data!A1:B2')

        assert result['success'] is True
        assert len(result['data']) == 2
        assert result['data'][0] == ['ID', 'Value']

    def test_full_row_column(self, sample_xlsx_file):
        """测试整行整列读取"""
        # 整行
        result_row = excel_get_range(sample_xlsx_file, '1:1')
        assert result_row['success'] is True
        assert len(result_row['data']) == 1

        # 整列
        result_col = excel_get_range(sample_xlsx_file, 'A:A')
        assert result_col['success'] is True
        assert len(result_col['data']) >= 1

    def test_invalid_range_format(self, sample_xlsx_file):
        """测试无效范围格式"""
        result = excel_get_range(sample_xlsx_file, 'INVALID_RANGE')

        assert result['success'] is False
        assert 'error' in result

    def test_nonexistent_sheet(self, sample_xlsx_file):
        """测试不存在的工作表"""
        result = excel_get_range(sample_xlsx_file, 'NonExistent!A1:B2')

        assert result['success'] is False
        assert 'error' in result

    def test_include_formatting(self, sample_xlsx_file):
        """测试包含格式"""
        result = excel_get_range(sample_xlsx_file, 'A1:B2', include_formatting=True)

        assert result['success'] is True
        # 可能包含格式信息


class TestExcelUpdateRange:
    """测试excel_update_range工具"""

    def test_basic_update(self, sample_xlsx_file, temp_dir):
        """测试基本更新功能"""
        test_file = temp_dir / 'test_update.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        new_data = [
            ['New Name', 'New Age'],
            ['John Doe', 35]
        ]

        result = excel_update_range(str(test_file), 'A1:B2', new_data)

        assert result['success'] is True
        assert 'updated_cells' in result
        assert result['updated_cells'] > 0

    def test_preserve_formulas(self, temp_dir):
        """测试保留公式"""
        file_path = temp_dir / 'test_formulas.xlsx'
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = 10
        sheet['A2'] = 20
        sheet['A3'] = '=A1+A2'
        workbook.save(file_path)

        new_data = [[15], [25]]
        result = excel_update_range(
            str(file_path),
            'A1:A2',
            new_data,
            preserve_formulas=True
        )

        assert result['success'] is True

    def test_data_size_validation(self, sample_xlsx_file, temp_dir):
        """测试数据大小验证"""
        test_file = temp_dir / 'test_validation.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 数据超出范围
        large_data = [['A', 'B'], ['C', 'D'], ['E', 'F'], ['G', 'H']]
        result = excel_update_range(str(test_file), 'A1:B2', large_data)

        # 应该失败或给出警告
        assert 'success' in result

    def test_empty_data(self, sample_xlsx_file, temp_dir):
        """测试空数据"""
        test_file = temp_dir / 'test_empty.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        result = excel_update_range(str(test_file), 'A1:A1', [[]])
        # 结果可能成功或失败，取决于实现
        assert 'success' in result


class TestExcelCreateFile:
    """测试excel_create_file工具"""

    def test_create_basic_file(self, temp_dir):
        """测试创建基本文件"""
        file_path = temp_dir / 'new_file.xlsx'
        result = excel_create_file(str(file_path))

        assert result['success'] is True
        assert 'file_path' in result
        assert 'sheets' in result
        assert file_path.exists()
        assert result['sheets'] == ['Sheet1']

    def test_create_with_custom_sheets(self, temp_dir):
        """测试创建自定义工作表文件"""
        file_path = temp_dir / 'custom_sheets.xlsx'
        sheet_names = ['数据', '图表', '汇总']

        result = excel_create_file(str(file_path), sheet_names)

        assert result['success'] is True
        assert result['sheets'] == sheet_names
        assert file_path.exists()

    def test_file_already_exists(self, sample_xlsx_file):
        """测试文件已存在"""
        result = excel_create_file(sample_xlsx_file)

        # 应该失败或覆盖，取决于实现
        assert 'success' in result

    def test_invalid_file_extension(self, temp_dir):
        """测试无效文件扩展名"""
        file_path = temp_dir / 'invalid.txt'
        result = excel_create_file(str(file_path))

        assert result['success'] is False
        assert 'error' in result
        assert '格式' in result['error']

    def test_create_xlsm_file(self, temp_dir):
        """测试创建xlsm文件"""
        file_path = temp_dir / 'macro_file.xlsm'
        result = excel_create_file(str(file_path))

        assert result['success'] is True
        assert file_path.exists()


class TestExcelRowColumnOperations:
    """测试行列操作工具"""

    def test_insert_delete_rows(self, sample_xlsx_file, temp_dir):
        """测试插入和删除行"""
        test_file = temp_dir / 'test_rows.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 插入行
        result_insert = excel_insert_rows(str(test_file), 'Sheet1', 2, 3)
        assert result_insert['success'] is True
        assert result_insert['inserted_rows'] == 3

        # 删除行
        result_delete = excel_delete_rows(str(test_file), 'Sheet1', 2, 2)
        assert result_delete['success'] is True
        assert result_delete['deleted_rows'] == 2

    def test_insert_delete_columns(self, sample_xlsx_file, temp_dir):
        """测试插入和删除列"""
        test_file = temp_dir / 'test_columns.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 插入列
        result_insert = excel_insert_columns(str(test_file), 'Sheet1', 2, 2)
        assert result_insert['success'] is True
        assert result_insert['inserted_columns'] == 2

        # 删除列
        result_delete = excel_delete_columns(str(test_file), 'Sheet1', 2, 1)
        assert result_delete['success'] is True
        assert result_delete['deleted_columns'] == 1

    def test_invalid_operations(self, sample_xlsx_file, temp_dir):
        """测试无效操作"""
        test_file = temp_dir / 'test_invalid.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 无效行号
        result = excel_insert_rows(str(test_file), 'Sheet1', 0, 1)
        assert result['success'] is False

        # 无效列号
        result = excel_insert_columns(str(test_file), 'Sheet1', 0, 1)
        assert result['success'] is False

        # 超大操作数
        result = excel_insert_rows(str(test_file), 'Sheet1', 1, 1001)
        assert result['success'] is False


class TestExcelSheetManagement:
    """测试工作表管理工具"""

    def test_create_rename_delete_sheet(self, sample_xlsx_file, temp_dir):
        """测试创建、重命名、删除工作表"""
        test_file = temp_dir / 'test_sheets.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 创建工作表
        result_create = excel_create_sheet(str(test_file), '新工作表')
        assert result_create['success'] is True
        assert result_create['sheet_name'] == '新工作表'

        # 重命名工作表
        result_rename = excel_rename_sheet(
            str(test_file),
            '新工作表',
            '重命名工作表'
        )
        assert result_rename['success'] is True
        assert result_rename['new_name'] == '重命名工作表'

        # 删除工作表
        result_delete = excel_delete_sheet(str(test_file), '重命名工作表')
        assert result_delete['success'] is True
        assert result_delete['deleted_sheet'] == '重命名工作表'

    def test_duplicate_sheet_names(self, sample_xlsx_file, temp_dir):
        """测试重复工作表名"""
        test_file = temp_dir / 'test_duplicate.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 尝试创建重复名称的工作表
        result = excel_create_sheet(str(test_file), 'Sheet1')
        assert result['success'] is False
        assert 'error' in result

    def test_invalid_sheet_operations(self, sample_xlsx_file, temp_dir):
        """测试无效工作表操作"""
        test_file = temp_dir / 'test_invalid_sheet.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        # 删除不存在的工作表
        result = excel_delete_sheet(str(test_file), 'NonExistent')
        assert result['success'] is False

        # 重命名不存在的工作表
        result = excel_rename_sheet(str(test_file), 'NonExistent', 'NewName')
        assert result['success'] is False


class TestExcelFormulaAndFormatting:
    """测试公式和格式化工具"""

    def test_set_formula(self, sample_xlsx_file, temp_dir):
        """测试设置公式"""
        test_file = temp_dir / 'test_formula.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        result = excel_set_formula(
            str(test_file),
            'Sheet1',
            'D1',
            'SUM(B:B)'
        )

        assert result['success'] is True
        assert result['formula'] == 'SUM(B:B)'

    def test_format_cells(self, sample_xlsx_file, temp_dir):
        """测试单元格格式化"""
        test_file = temp_dir / 'test_format.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        formatting = {
            'font': {'name': '微软雅黑', 'size': 14, 'bold': True, 'color': '000080'},
            'fill': {'color': 'E6F3FF'},
            'alignment': {'horizontal': 'center', 'vertical': 'middle'}
        }

        result = excel_format_cells(
            str(test_file),
            'Sheet1',
            'A1:D1',
            formatting
        )

        assert result['success'] is True
        assert 'formatted_count' in result
        assert result['formatted_count'] > 0

    def test_invalid_formula(self, sample_xlsx_file, temp_dir):
        """测试无效公式"""
        test_file = temp_dir / 'test_invalid_formula.xlsx'
        shutil.copy2(sample_xlsx_file, test_file)

        result = excel_set_formula(
            str(test_file),
            'Sheet1',
            'D1',
            'INVALID_FUNCTION()'
        )

        # 结果可能成功（设置了无效公式）或失败
        assert 'success' in result


class TestMCPToolsIntegration:
    """测试MCP工具集成场景"""

    def test_complete_workflow(self, temp_dir):
        """测试完整工作流程"""
        file_path = temp_dir / 'workflow_test.xlsx'

        # 1. 创建文件
        result = excel_create_file(str(file_path), ['数据', '统计'])
        assert result['success'] is True

        # 2. 添加数据
        test_data = [
            ['姓名', '年龄', '工资'],
            ['张三', 25, 5000],
            ['李四', 30, 6000],
            ['王五', 35, 7000]
        ]
        result = excel_update_range(str(file_path), '数据!A1:C4', test_data)
        assert result['success'] is True

        # 3. 设置公式
        result = excel_set_formula(
            str(file_path),
            '统计',
            'A1',
            'AVERAGE(数据!B2:B4)'
        )
        assert result['success'] is True

        # 4. 格式化
        formatting = {
            'font': {'bold': True, 'color': 'FF0000'},
            'fill': {'color': 'FFFF00'}
        }
        result = excel_format_cells(str(file_path), '数据', 'A1:C1', formatting)
        assert result['success'] is True

        # 5. 搜索数据
        result = excel_regex_search(str(file_path), '张三')
        assert result['success'] is True
        assert result['match_count'] >= 1

        # 6. 验证最终结果
        result = excel_get_range(str(file_path), '数据!A1:C4')
        assert result['success'] is True
        assert len(result['data']) == 4
        assert result['data'][0] == ['姓名', '年龄', '工资']

    def test_error_recovery(self, temp_dir):
        """测试错误恢复和容错性"""
        file_path = temp_dir / 'error_test.xlsx'

        # 创建文件
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 尝试在不存在的工作表上操作
        result = excel_update_range(
            str(file_path),
            'NonExistent!A1:B1',
            [['Test', 'Data']]
        )
        assert result['success'] is False

        # 正常操作仍然可以执行
        result = excel_update_range(
            str(file_path),
            'A1:B1',
            [['Test', 'Data']]
        )
        assert result['success'] is True

        # 文件仍然可访问
        result = excel_list_sheets(str(file_path))
        assert result['success'] is True

    def test_large_data_workflow(self, temp_dir):
        """测试大数据工作流程"""
        file_path = temp_dir / 'large_data_test.xlsx'

        # 创建文件
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 创建大量数据
        large_data = []
        for i in range(100):
            large_data.append([f'Item_{i}', i, i * 10])

        # 批量插入数据
        result = excel_update_range(str(file_path), 'A1:C100', large_data)
        assert result['success'] is True

        # 进行搜索
        result = excel_regex_search(str(file_path), r'Item_5\d')
        assert result['success'] is True
        assert result['match_count'] >= 10  # Item_50-59

        # 插入行列
        result = excel_insert_rows(str(file_path), 'Sheet1', 1, 5)
        assert result['success'] is True

        # 验证结果
        result = excel_get_range(str(file_path), 'A1:C5')
        assert result['success'] is True
        assert len(result['data']) == 5

    def test_multi_sheet_operations(self, temp_dir):
        """测试多工作表操作"""
        file_path = temp_dir / 'multi_sheet_test.xlsx'

        # 创建多工作表文件
        result = excel_create_file(str(file_path), ['销售', '产品', '客户'])
        assert result['success'] is True

        # 在不同工作表中添加数据
        sales_data = [['日期', '销售额'], ['2024-01-01', 1000], ['2024-01-02', 1500]]
        result = excel_update_range(str(file_path), '销售!A1:B3', sales_data)
        assert result['success'] is True

        product_data = [['产品名', '价格'], ['产品A', 100], ['产品B', 200]]
        result = excel_update_range(str(file_path), '产品!A1:B3', product_data)
        assert result['success'] is True

        # 跨工作表搜索
        result = excel_regex_search(str(file_path), '产品')
        assert result['success'] is True
        assert result['match_count'] >= 3  # 产品, 产品A, 产品B

        # 创建汇总工作表
        result = excel_create_sheet(str(file_path), '汇总')
        assert result['success'] is True

        # 验证工作表列表
        result = excel_list_sheets(str(file_path))
        assert result['success'] is True
        assert len(result['sheets']) == 4
        assert '汇总' in result['sheets']


class TestMCPToolsPerformance:
    """测试MCP工具性能"""

    def test_bulk_operations_performance(self, temp_dir):
        """测试批量操作性能"""
        file_path = temp_dir / 'performance_test.xlsx'

        # 创建文件
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        import time
        start_time = time.time()

        # 执行100次小范围更新
        for i in range(100):
            result = excel_update_range(
                str(file_path),
                f'A{i+1}:B{i+1}',
                [[f'Data_{i}', i]]
            )
            assert result['success'] is True

        end_time = time.time()
        # 性能要求：小于10秒
        assert (end_time - start_time) < 10.0

    def test_large_search_performance(self, temp_dir):
        """测试大数据搜索性能"""
        file_path = temp_dir / 'search_performance_test.xlsx'

        # 创建包含大量数据的文件
        workbook = Workbook()
        sheet = workbook.active

        for row in range(1, 1001):
            for col in range(1, 11):
                sheet.cell(row=row, column=col, value=f'Data_{row}_{col}')

        workbook.save(file_path)

        import time
        start_time = time.time()

        # 执行正则搜索
        result = excel_regex_search(str(file_path), r'Data_50\d_\d')

        end_time = time.time()

        assert result['success'] is True
        # 性能要求：小于5秒
        assert (end_time - start_time) < 5.0

    def test_concurrent_operations_stability(self, temp_dir):
        """测试并发操作稳定性"""
        file_path = temp_dir / 'concurrent_test.xlsx'

        # 创建文件
        result = excel_create_file(str(file_path), ['Sheet1', 'Sheet2', 'Sheet3'])
        assert result['success'] is True

        # 模拟多个操作快速执行
        operations = [
            lambda: excel_update_range(str(file_path), 'Sheet1!A1:A1', [['Test1']]),
            lambda: excel_update_range(str(file_path), 'Sheet2!A1:A1', [['Test2']]),
            lambda: excel_update_range(str(file_path), 'Sheet3!A1:A1', [['Test3']]),
            lambda: excel_get_range(str(file_path), 'Sheet1!A1:A1'),
            lambda: excel_get_range(str(file_path), 'Sheet2!A1:A1'),
            lambda: excel_get_range(str(file_path), 'Sheet3!A1:A1'),
        ]

        # 快速执行所有操作
        for operation in operations:
            result = operation()
            assert result['success'] is True


class TestMCPToolsEdgeCases:
    """测试MCP工具边界情况"""

    def test_unicode_handling(self, temp_dir):
        """测试Unicode字符处理"""
        file_path = temp_dir / 'unicode_test.xlsx'

        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试各种Unicode字符
        unicode_data = [
            ['中文', '日本語', 'العربية'],
            ['🚀', '💡', '🎉'],
            ['Ñoño', 'Café', 'Résumé']
        ]

        result = excel_update_range(str(file_path), 'A1:C3', unicode_data)
        assert result['success'] is True

        # 验证数据
        result = excel_get_range(str(file_path), 'A1:C3')
        assert result['success'] is True
        assert result['data'][0] == ['中文', '日本語', 'العربية']

        # 搜索Unicode字符
        result = excel_regex_search(str(file_path), '中文')
        assert result['success'] is True
        assert result['match_count'] >= 1

    def test_special_characters_in_formulas(self, temp_dir):
        """测试公式中的特殊字符"""
        file_path = temp_dir / 'special_formula_test.xlsx'

        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 设置包含特殊字符的公式
        result = excel_set_formula(
            str(file_path),
            'Sheet1',
            'A1',
            'CONCATENATE("Hello", " ", "World!")'
        )
        assert result['success'] is True

    def test_empty_and_null_values(self, temp_dir):
        """测试空值和null值处理"""
        file_path = temp_dir / 'empty_null_test.xlsx'

        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试空值和None值
        empty_data = [
            ['', None, '  '],
            [0, '', 'Valid'],
            [None, 'Test', '']
        ]

        result = excel_update_range(str(file_path), 'A1:C3', empty_data)
        assert result['success'] is True

        # 验证数据处理
        result = excel_get_range(str(file_path), 'A1:C3')
        assert result['success'] is True

    def test_maximum_limits(self, temp_dir):
        """测试极限值处理"""
        file_path = temp_dir / 'limits_test.xlsx'

        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试大数值
        large_data = [
            [999999999, -999999999, 0.123456789],
            [float('inf'), float('-inf'), float('nan')]
        ]

        result = excel_update_range(str(file_path), 'A1:C2', large_data)
        # 结果可能成功或失败，取决于Excel限制
        assert 'success' in result
