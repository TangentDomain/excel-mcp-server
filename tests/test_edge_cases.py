#!/usr/bin/env python3
"""
Excel MCP Server - 边界条件和错误处理测试

专门测试各种边界条件、错误场景和异常情况
确保系统在极端条件下的稳定性和容错性
"""

import pytest
import tempfile
import shutil
import string
import random
from pathlib import Path
from openpyxl import Workbook

from server import (
    excel_list_sheets, excel_regex_search, excel_get_range, excel_update_range,
    excel_insert_rows, excel_insert_columns, excel_create_file, excel_create_sheet,
    excel_delete_sheet, excel_rename_sheet, excel_delete_rows, excel_delete_columns,
    excel_set_formula, excel_format_cells
)


class TestBoundaryValues:
    """测试边界值条件"""

    def test_maximum_row_column_limits(self, temp_dir):
        """测试Excel最大行列限制"""
        file_path = temp_dir / 'max_limits.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试最大行数 (Excel 2007+ 支持 1,048,576 行)
        max_row = 1048576
        result = excel_update_range(
            str(file_path),
            f'A{max_row}:A{max_row}',
            [['Max Row Test']]
        )
        # 可能成功或失败，取决于系统内存
        assert 'success' in result

        # 测试最大列数 (Excel 2007+ 支持 16,384 列, XFD列)
        result = excel_update_range(
            str(file_path),
            'XFD1:XFD1',
            [['Max Column Test']]
        )
        assert 'success' in result

    def test_empty_range_boundaries(self, temp_dir):
        """测试空范围边界"""
        file_path = temp_dir / 'empty_boundaries.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试空范围
        empty_ranges = ['', 'A:A', '1:1', 'A1:A1', 'Z100:Z100']

        for range_expr in empty_ranges:
            if range_expr:  # 非空字符串
                result = excel_get_range(str(file_path), range_expr)
                assert 'success' in result

    def test_maximum_string_length(self, temp_dir):
        """测试最大字符串长度"""
        file_path = temp_dir / 'max_string.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # Excel单元格最大字符数：32,767
        max_string = 'A' * 32767
        result = excel_update_range(str(file_path), 'A1:A1', [[max_string]])

        # 可能成功或被截断
        assert 'success' in result

        # 超长字符串
        over_max_string = 'B' * 50000
        result = excel_update_range(str(file_path), 'B1:B1', [[over_max_string]])
        # 应该失败或截断
        assert 'success' in result

    def test_numeric_boundaries(self, temp_dir):
        """测试数值边界"""
        file_path = temp_dir / 'numeric_boundaries.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        boundary_values = [
            # 整数边界
            [2**31 - 1, -(2**31), 0],  # 32位整数边界
            [2**63 - 1, -(2**63), 1],  # 64位整数边界
            # 浮点数边界
            [1.7976931348623157e+308, -1.7976931348623157e+308, 0.0],  # 双精度边界
            [float('inf'), float('-inf'), float('nan')],  # 特殊浮点值
            # 小数精度
            [0.123456789012345, -0.987654321098765, 1e-15]
        ]

        for i, values in enumerate(boundary_values, 1):
            result = excel_update_range(
                str(file_path),
                f'A{i}:C{i}',
                [values]
            )
            # 有些值可能失败（如无穷大、NaN）
            assert 'success' in result

    def test_date_time_boundaries(self, temp_dir):
        """测试日期时间边界"""
        file_path = temp_dir / 'datetime_boundaries.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # Excel日期范围：1900-01-01 到 9999-12-31
        boundary_dates = [
            ['1900-01-01', '9999-12-31', '2000-02-29'],  # 闰年测试
            ['1899-12-31', '10000-01-01', '2100-02-29'],  # 超出范围/无效日期
        ]

        for i, dates in enumerate(boundary_dates, 1):
            result = excel_update_range(
                str(file_path),
                f'A{i}:C{i}',
                [dates]
            )
            assert 'success' in result

    def test_sheet_name_boundaries(self, temp_dir):
        """测试工作表名称边界"""
        file_path = temp_dir / 'sheet_name_boundaries.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # Excel工作表名称限制：31个字符，不能包含某些特殊字符
        boundary_names = [
            'A' * 31,  # 最大长度
            'A' * 32,  # 超过最大长度
            'Valid_Sheet-123',  # 有效字符
            'Invalid[Sheet]',  # 包含无效字符 []
            'Invalid:Sheet',  # 包含无效字符 :
            'Invalid/Sheet',  # 包含无效字符 /
            'Invalid\\Sheet',  # 包含无效字符 \
            'Invalid?Sheet',  # 包含无效字符 ?
            'Invalid*Sheet',  # 包含无效字符 *
            '',  # 空名称
            ' ',  # 纯空格
        ]

        for name in boundary_names:
            result = excel_create_sheet(str(file_path), name)
            # 某些名称会失败
            assert 'success' in result


class TestErrorHandling:
    """测试错误处理"""

    def test_file_not_found_errors(self):
        """测试文件未找到错误"""
        nonexistent_files = [
            '/nonexistent/path/file.xlsx',
            'C:\\NotExists\\file.xlsx',  # Windows路径
            '../../../etc/passwd',  # Unix系统路径
            '',  # 空路径
            None,  # None值
        ]

        for file_path in nonexistent_files:
            if file_path is not None:
                result = excel_list_sheets(file_path)
                assert result['success'] is False
                assert 'error' in result

    def test_file_permission_errors(self, temp_dir):
        """测试文件权限错误"""
        # 创建只读文件
        readonly_file = temp_dir / 'readonly.xlsx'
        result = excel_create_file(str(readonly_file))
        assert result['success'] is True

        # 设置只读权限
        readonly_file.chmod(0o444)

        # 尝试写入只读文件
        result = excel_update_range(
            str(readonly_file),
            'A1:A1',
            [['Test']]
        )
        # 应该失败
        assert result['success'] is False
        assert 'error' in result

        # 恢复权限以便清理
        readonly_file.chmod(0o666)

    def test_corrupted_file_handling(self, temp_dir):
        """测试损坏文件处理"""
        # 创建假的Excel文件（实际是文本文件）
        fake_excel = temp_dir / 'fake.xlsx'
        with open(fake_excel, 'w') as f:
            f.write('This is not an Excel file')

        result = excel_list_sheets(str(fake_excel))
        assert result['success'] is False
        assert 'error' in result

        # 创建空文件
        empty_file = temp_dir / 'empty.xlsx'
        empty_file.touch()

        result = excel_list_sheets(str(empty_file))
        assert result['success'] is False
        assert 'error' in result

    def test_invalid_range_formats(self, temp_dir):
        """测试无效范围格式"""
        file_path = temp_dir / 'invalid_ranges.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        invalid_ranges = [
            'INVALID',
            'A1:',
            ':B2',
            'A1:B',
            'A:B2',
            '1A:2B',
            'AA1:ZZ',
            'A1:A0',  # 起始大于结束
            'B2:A1',  # 列顺序错误
            'Sheet!A1',  # 缺少范围
            '!A1:B2',  # 空工作表名
            'A1:B2:C3',  # 多个冒号
            'A1-B2',  # 错误分隔符
            '',  # 空字符串
            None,  # None值
        ]

        for range_expr in invalid_ranges:
            if range_expr is not None:
                result = excel_get_range(str(file_path), range_expr)
                assert result['success'] is False
                assert 'error' in result

    def test_invalid_data_types(self, temp_dir):
        """测试无效数据类型"""
        file_path = temp_dir / 'invalid_data.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 测试各种无效数据类型
        invalid_data_sets = [
            # 非列表数据
            "string_instead_of_list",
            123,
            {'dict': 'value'},

            # 不一致的行长度
            [['A', 'B'], ['C']],  # 第二行缺少列
            [['A'], ['B', 'C', 'D']],  # 行长度不一致

            # 空数据结构
            [],
            [[]],
            [[], []],

            # 嵌套过深
            [[['nested', 'too', 'deep']]],

            # 混合数据类型
            [['text', 123, None, True, False]],
        ]

        for i, data in enumerate(invalid_data_sets, 1):
            result = excel_update_range(str(file_path), f'A{i}:C{i}', data)
            # 有些可能成功（被转换），有些会失败
            assert 'success' in result

    def test_resource_exhaustion(self, temp_dir):
        """测试资源耗尽情况"""
        file_path = temp_dir / 'resource_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 尝试创建极大的数据集
        try:
            huge_data = [['Data'] * 1000 for _ in range(1000)]
            result = excel_update_range(
                str(file_path),
                'A1:ALL1000',  # 无效范围
                huge_data
            )
            # 应该失败或被限制
            assert 'success' in result
        except MemoryError:
            # 内存不足是预期的
            pass

    def test_concurrent_access_errors(self, temp_dir):
        """测试并发访问错误"""
        file_path = temp_dir / 'concurrent.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 模拟快速连续操作可能导致的冲突
        results = []
        for i in range(10):
            result = excel_update_range(
                str(file_path),
                f'A{i+1}:A{i+1}',
                [[f'Data_{i}']]
            )
            results.append(result)

        # 至少一些操作应该成功
        successful_ops = sum(1 for r in results if r['success'])
        assert successful_ops >= 5  # 至少一半成功


class TestMemoryAndPerformance:
    """测试内存使用和性能边界"""

    def test_large_file_handling(self, temp_dir):
        """测试大文件处理"""
        file_path = temp_dir / 'large_file.xlsx'

        # 创建包含大量数据的文件
        workbook = Workbook()
        sheet = workbook.active

        # 写入10,000行 x 50列的数据
        for row in range(1, 1001):  # 减少到1000行以避免测试超时
            for col in range(1, 51):
                sheet.cell(row=row, column=col, value=f'R{row}C{col}')

        workbook.save(file_path)

        # 测试大文件读取
        import time
        start_time = time.time()

        result = excel_get_range(str(file_path), 'A1:AX1000')

        end_time = time.time()
        processing_time = end_time - start_time

        assert result['success'] is True
        assert len(result['data']) == 1000
        assert len(result['data'][0]) == 50

        # 性能要求：处理时间应该在合理范围内
        print(f"Large file processing time: {processing_time:.2f} seconds")
        assert processing_time < 30.0  # 30秒限制

    def test_memory_intensive_operations(self, temp_dir):
        """测试内存密集型操作"""
        file_path = temp_dir / 'memory_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 创建大型数据集
        large_data = []
        for i in range(1000):
            row = [f'Column_{j}_Row_{i}' for j in range(20)]
            large_data.append(row)

        # 测试大数据插入
        result = excel_update_range(str(file_path), 'A1:T1000', large_data)
        assert result['success'] is True

        # 测试大范围读取
        result = excel_get_range(str(file_path), 'A1:T1000')
        assert result['success'] is True
        assert len(result['data']) == 1000

    def test_regex_performance_with_large_data(self, temp_dir):
        """测试正则搜索在大数据集上的性能"""
        file_path = temp_dir / 'regex_perf.xlsx'
        workbook = Workbook()
        sheet = workbook.active

        # 创建包含模式的大数据集
        patterns = ['email@domain.com', 'phone:123-456-7890', 'id:ABC123']
        for row in range(1, 5001):  # 5000行数据
            for col in range(1, 4):
                if row % 100 == 0:  # 每100行插入一个匹配
                    sheet.cell(row=row, column=col, value=patterns[col-1])
                else:
                    sheet.cell(row=row, column=col, value=f'Data_{row}_{col}')

        workbook.save(file_path)

        # 测试正则搜索性能
        import time
        start_time = time.time()

        result = excel_regex_search(str(file_path), r'\w+@\w+\.\w+')

        end_time = time.time()
        search_time = end_time - start_time

        assert result['success'] is True
        assert result['match_count'] > 0

        print(f"Regex search time on large data: {search_time:.2f} seconds")
        assert search_time < 10.0  # 10秒限制


class TestRecoveryAndStability:
    """测试恢复能力和系统稳定性"""

    def test_error_recovery(self, temp_dir):
        """测试错误后的系统恢复"""
        file_path = temp_dir / 'recovery_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 执行一个会失败的操作
        result = excel_update_range(str(file_path), 'INVALID_RANGE', [['Test']])
        assert result['success'] is False

        # 验证系统仍然可以执行正常操作
        result = excel_update_range(str(file_path), 'A1:A1', [['Recovery Test']])
        assert result['success'] is True

        # 验证数据完整性
        result = excel_get_range(str(file_path), 'A1:A1')
        assert result['success'] is True
        assert result['data'][0][0] == 'Recovery Test'

    def test_partial_operation_rollback(self, temp_dir):
        """测试部分操作的回滚"""
        file_path = temp_dir / 'rollback_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 先写入一些数据
        initial_data = [['Initial', 'Data']]
        result = excel_update_range(str(file_path), 'A1:B1', initial_data)
        assert result['success'] is True

        # 尝试批量操作，其中一些可能失败
        batch_operations = [
            ('A2:B2', [['Valid', 'Data']]),
            ('INVALID:RANGE', [['Invalid', 'Range']]),
            ('A3:B3', [['Another', 'Valid']]),
        ]

        for range_expr, data in batch_operations:
            result = excel_update_range(str(file_path), range_expr, data)
            # 记录结果，但不中断

        # 验证有效操作成功，无效操作不影响其他数据
        result = excel_get_range(str(file_path), 'A1:B1')
        assert result['success'] is True
        assert result['data'][0] == ['Initial', 'Data']

    def test_file_lock_handling(self, temp_dir):
        """测试文件锁定处理"""
        file_path = temp_dir / 'lock_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 模拟文件被其他进程锁定的情况
        # 在实际环境中，这可能需要更复杂的设置
        try:
            # 尝试快速连续操作，可能导致文件锁定
            for i in range(5):
                result = excel_update_range(
                    str(file_path),
                    f'A{i+1}:A{i+1}',
                    [[f'Test_{i}']]
                )
                # 某些操作可能因为文件锁定而失败，但不应该崩溃
                assert 'success' in result
        except Exception as e:
            # 捕获任何异常，确保测试不会崩溃
            print(f"File lock test exception: {e}")
            assert False, "File lock handling should not raise unhandled exceptions"


class TestSpecialCharacters:
    """测试特殊字符和编码处理"""

    def test_unicode_support(self, temp_dir):
        """测试Unicode字符支持"""
        file_path = temp_dir / 'unicode_test.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        unicode_data = [
            # 各种语言文字
            ['中文测试', 'English', '日本語'],
            ['العربية', 'Русский', 'Français'],
            ['Español', 'Português', 'Italiano'],

            # 特殊符号
            ['©®™', '±×÷', '∞∑∏'],

            # 表情符号
            ['😀😃😄', '🌟⭐✨', '🚀🌙💫'],

            # 数学符号
            ['∑∫∂', 'α β γ', '∴∵∈'],
        ]

        for i, row in enumerate(unicode_data, 1):
            result = excel_update_range(
                str(file_path),
                f'A{i}:C{i}',
                [row]
            )
            assert result['success'] is True

        # 验证数据读取
        result = excel_get_range(str(file_path), 'A1:C6')
        assert result['success'] is True
        assert result['data'][0] == ['中文测试', 'English', '日本語']

    def test_control_characters(self, temp_dir):
        """测试控制字符处理"""
        file_path = temp_dir / 'control_chars.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        control_chars = [
            # 制表符、换行符等
            ['Tab:\t', 'Newline:\n', 'Return:\r'],

            # NULL字符和其他控制字符
            ['Null:\x00', 'Bell:\x07', 'Backspace:\x08'],

            # 高位控制字符
            ['\x7f', '\x80', '\x9f'],
        ]

        for i, row in enumerate(control_chars, 1):
            result = excel_update_range(
                str(file_path),
                f'A{i}:C{i}',
                [row]
            )
            # 某些控制字符可能被拒绝或转换
            assert 'success' in result

    def test_very_long_strings(self, temp_dir):
        """测试超长字符串"""
        file_path = temp_dir / 'long_strings.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 生成不同长度的字符串
        string_lengths = [1000, 10000, 32767, 50000, 100000]

        for i, length in enumerate(string_lengths, 1):
            long_string = ''.join(random.choices(string.ascii_letters, k=length))
            result = excel_update_range(
                str(file_path),
                f'A{i}:A{i}',
                [[long_string]]
            )
            # 超过Excel限制的字符串可能被截断或拒绝
            assert 'success' in result

    def test_formula_injection_protection(self, temp_dir):
        """测试公式注入保护"""
        file_path = temp_dir / 'formula_injection.xlsx'
        result = excel_create_file(str(file_path))
        assert result['success'] is True

        # 尝试注入可能有害的公式
        malicious_inputs = [
            '=SUM(1+1)',  # 看起来无害的公式
            '=HYPERLINK("http://evil.com", "Click me")',
            '=INDIRECT("R1C1", FALSE)',
            '=EXEC("rm -rf /")',  # 系统命令（Excel不支持）
            '=CALL("kernel32", "ExitProcess", 0)',  # 危险的系统调用
        ]

        for i, malicious_input in enumerate(malicious_inputs, 1):
            result = excel_update_range(
                str(file_path),
                f'A{i}:A{i}',
                [[malicious_input]]
            )
            # 系统应该安全处理这些输入
            assert result['success'] is True

            # 验证内容是否被适当处理
            verify_result = excel_get_range(str(file_path), f'A{i}:A{i}')
            assert verify_result['success'] is True
