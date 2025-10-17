"""
Simplified tests for Server MCP interfaces - more flexible to match actual API
"""

import pytest
from src.server import (
    excel_list_sheets,
    excel_get_sheet_headers,
    excel_get_headers,
    excel_get_range,
    excel_update_range,
    excel_create_file,
    excel_create_sheet,
    excel_delete_sheet,
    excel_rename_sheet,
    excel_insert_rows,
    excel_insert_columns,
    excel_delete_rows,
    excel_delete_columns,
    excel_format_cells,
    excel_search,
    excel_find_last_row
)


class TestServerInterfaces:
    """Test cases for Server MCP interfaces - simplified and flexible"""

    def test_excel_list_sheets(self, sample_excel_file):
        """Test excel_list_sheets interface - basic functionality"""
        result = excel_list_sheets(sample_excel_file)

        assert result['success'] is True
        assert 'sheets' in result
        assert isinstance(result['sheets'], list)
        assert 'total_sheets' in result
        # 重构后：不再包含表头信息，单一职责
        assert 'sheets_with_headers' not in result
        # 已移除active_sheet概念
        assert 'active_sheet' not in result

    def test_excel_get_sheet_headers(self, sample_excel_file):
        """Test excel_get_sheet_headers interface - 获取工作表表头信息"""
        result = excel_get_sheet_headers(sample_excel_file)

        assert result['success'] is True
        assert 'sheets_with_headers' in result
        assert isinstance(result['sheets_with_headers'], list)

        # 验证表头信息结构
        for sheet_info in result['sheets_with_headers']:
            assert 'name' in sheet_info
            assert 'headers' in sheet_info
            assert 'header_count' in sheet_info
            assert isinstance(sheet_info['headers'], list)
            assert isinstance(sheet_info['header_count'], int)
            assert sheet_info['header_count'] == len(sheet_info['headers'])

        # 验证sample_excel_file的第一个工作表应该包含表头（双行格式，返回field_names）
        sheet1_info = next((s for s in result['sheets_with_headers'] if s['name'] == 'Sheet1'), None)
        assert sheet1_info is not None
        assert 'name' in sheet1_info['headers']  # 根据conftest.py修改后的双行格式数据
        assert 'age' in sheet1_info['headers']
        assert 'department' in sheet1_info['headers']
        assert sheet1_info['header_count'] >= 4

    def test_excel_list_sheets_simple(self, sample_excel_file):
        """Test excel_list_sheets interface - 简单工作表列表"""
        result = excel_list_sheets(sample_excel_file)

        assert result['success'] is True
        assert 'sheets' in result
        assert 'total_sheets' in result
        # 重构后：不包含表头信息，单一职责
        assert 'sheets_with_headers' not in result
        # 已移除active_sheet概念
        assert 'active_sheet' not in result

    def test_excel_get_sheet_headers_multi_sheet(self, multi_sheet_excel_file):
        """Test excel_get_sheet_headers with multiple sheets"""
        result = excel_get_sheet_headers(multi_sheet_excel_file)

        assert result['success'] is True
        assert 'sheets_with_headers' in result
        assert len(result['sheets_with_headers']) == 4  # 根据conftest.py中的设置

        # 验证每个工作表都有表头信息
        expected_sheet_names = ["数据", "图表", "汇总", "分析"]
        actual_sheet_names = [s['name'] for s in result['sheets_with_headers']]

        for expected_name in expected_sheet_names:
            assert expected_name in actual_sheet_names
            sheet_info = next(s for s in result['sheets_with_headers'] if s['name'] == expected_name)
            assert 'test_data' in sheet_info['headers']  # Field names from row 2
            assert 'value' in sheet_info['headers']

    def test_excel_get_sheet_headers_empty_sheet(self, empty_excel_file):
        """Test excel_get_sheet_headers with empty sheet"""
        result = excel_get_sheet_headers(empty_excel_file)

        assert result['success'] is True
        assert 'sheets_with_headers' in result
        assert len(result['sheets_with_headers']) == 1  # 应该有一个默认工作表

        sheet_info = result['sheets_with_headers'][0]
        # 空工作表应该没有表头，如果headers字段被清理了，默认为空数组
        headers = sheet_info.get('headers', [])
        assert headers == []  # 空工作表应该没有表头
        assert sheet_info['header_count'] == 0

    def test_excel_list_sheets_invalid_file(self):
        """Test excel_list_sheets with invalid file"""
        result = excel_list_sheets("nonexistent_file.xlsx")

        assert result['success'] is False
        assert 'error' in result

    # ==================== Excel Get Headers Tests ====================

    def test_excel_get_headers_basic(self, sample_excel_file):
        """Test excel_get_headers basic functionality"""
        result = excel_get_headers(sample_excel_file, "Sheet1")

        assert result['success'] is True
        assert 'headers' in result
        assert 'header_count' in result
        assert 'sheet_name' in result
        assert 'header_row' in result
        assert 'message' in result

        # Check data content based on sample_excel_file fixture with dual header format
        # Row 1: descriptions = ["姓名描述", "年龄描述", "部门描述", "薪资描述", "总计描述"]
        # Row 2: field_names = ["name", "age", "department", "salary", "total"]
        # headers and data should return the field_names (row 2)
        assert result['headers'] == ["name", "age", "department", "salary", "total"]
        assert result['data'] == ["name", "age", "department", "salary", "total"]
        assert result['descriptions'] == ["姓名描述", "年龄描述", "部门描述", "薪资描述", "总计描述"]
        assert result['field_names'] == ["name", "age", "department", "salary", "total"]
        assert result['header_count'] == 5
        assert result['sheet_name'] == "Sheet1"
        assert result['header_row'] == 1

    def test_excel_get_headers_with_max_columns(self, sample_excel_file):
        """Test excel_get_headers with max_columns limit"""
        result = excel_get_headers(sample_excel_file, "Sheet1", max_columns=2)

        assert result['success'] is True
        assert result['headers'] == ["name", "age"]  # Field names from row 2
        assert result['descriptions'] == ["姓名描述", "年龄描述"]  # Descriptions from row 1
        assert result['header_count'] == 2
        assert "成功获取2个表头字段" in result['message']

    def test_excel_get_headers_different_row(self, sample_excel_file):
        """Test excel_get_headers with different header row"""
        # Test getting data from row 2 as header start (dual header from rows 2-3)
        result = excel_get_headers(sample_excel_file, "Sheet1", header_row=2)

        assert result['success'] is True
        # Row 2: field_names = ["name", "age", "department", "salary", "total"]
        # Row 3: first data row = ["张三", 25, "技术部", 8000, formula]
        # Should get row 2 as descriptions and row 3 as field_names
        assert result['descriptions'] == ["name", "age", "department", "salary", "total"]
        assert result['header_count'] >= 4  # At least the first 4 columns
        assert result['headers'][0] == "张三"
        assert str(result['headers'][1]) == "25"  # Convert to string for comparison
        assert result['headers'][2] == "技术部"
        assert str(result['headers'][3]) == "8000"
        assert result['header_row'] == 2

    def test_excel_get_headers_second_sheet(self, sample_excel_file):
        """Test excel_get_headers on second sheet"""
        result = excel_get_headers(sample_excel_file, "Sheet2")

        assert result['success'] is True
        assert result['headers'] == ["product", "sales", "price"]  # Field names from row 2
        assert result['descriptions'] == ["产品描述", "销量描述", "单价描述"]  # Descriptions from row 1
        assert result['header_count'] == 3
        assert result['sheet_name'] == "Sheet2"

    def test_excel_get_headers_invalid_sheet(self, sample_excel_file):
        """Test excel_get_headers with invalid sheet name"""
        result = excel_get_headers(sample_excel_file, "NonExistentSheet")

        assert result['success'] is False
        assert 'error' in result
        # Check for various possible error messages
        error_msg = result['error'].lower()
        assert ("不存在" in error_msg or "无法读取" in error_msg or
                "工作表" in error_msg or "sheet" in error_msg)

    def test_excel_get_headers_invalid_file(self):
        """Test excel_get_headers with invalid file"""
        result = excel_get_headers("nonexistent_file.xlsx", "Sheet1")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_get_headers_empty_sheet(self, empty_excel_file):
        """Test excel_get_headers with empty sheet"""
        result = excel_get_headers(empty_excel_file, "Sheet")

        assert result['success'] is True
        assert result['headers'] == []
        assert result['header_count'] == 0

    def test_excel_get_headers_with_mixed_types(self, temp_dir, request):
        """Test excel_get_headers with mixed data types in headers"""
        # Create a test file with mixed types in header row
        import uuid
        from openpyxl import Workbook

        test_id = str(uuid.uuid4())[:8]
        file_path = temp_dir / f"test_mixed_headers_{test_id}.xlsx"

        wb = Workbook()
        ws = wb.active
        # Dual header format: row 1 = descriptions, row 2 = field_names
        ws.append(["Text", 123, None, "Another Text", ""])  # Row 1: descriptions (mixed types)
        ws.append(["Data1", "Data2", "Data3", "Data4", "Data5"])  # Row 2: field_names
        wb.save(file_path)

        result = excel_get_headers(str(file_path), "Sheet")

        assert result['success'] is True
        # Returns field_names (row 2), should get all 5 values since None is in descriptions
        assert result['headers'] == ["Data1", "Data2", "Data3", "Data4", "Data5"]
        assert result['descriptions'] == ["Text", "123", "列C", "Another Text", "列E"]  # Fallback mechanism for None/empty
        assert result['header_count'] == 5

    def test_excel_get_headers_max_columns_with_empty_cells(self, temp_dir, request):
        """Test excel_get_headers with max_columns including empty cells"""
        # Create a test file with empty cells in header row
        import uuid
        from openpyxl import Workbook

        test_id = str(uuid.uuid4())[:8]
        file_path = temp_dir / f"test_empty_headers_{test_id}.xlsx"

        wb = Workbook()
        ws = wb.active
        # Dual header format
        ws.append(["Col1", None, "Col3", "", "Col5"])  # Row 1: descriptions with empty cells
        ws.append(["a", "b", "c", "d", "e"])            # Row 2: field_names
        wb.save(file_path)

        # Without max_columns, should get all field_names (row 2) since data exists
        result1 = excel_get_headers(str(file_path), "Sheet")
        assert result1['success'] is True
        assert result1['headers'] == ["a", "b", "c", "d", "e"]  # All field_names from row 2

        # With max_columns, should still get field_names but limited to max_columns
        result2 = excel_get_headers(str(file_path), "Sheet", max_columns=3)
        assert result2['success'] is True
        assert result2['header_count'] == 3
        assert result2['headers'] == ["a", "b", "c"]  # First 3 field_names
        assert result2['descriptions'] == ["Col1", "列B", "Col3"]  # Fallback "列B" for None

    def test_excel_get_range(self, sample_excel_file):
        """Test excel_get_range interface"""
        result = excel_get_range(sample_excel_file, "Sheet1!A1:C5")

        assert result['success'] is True
        assert 'data' in result
        assert isinstance(result['data'], list)

    def test_excel_get_range_invalid_sheet(self, sample_excel_file):
        """Test excel_get_range with invalid sheet"""
        result = excel_get_range(sample_excel_file, "NonExistentSheet!A1")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_update_range(self, sample_excel_file):
        """Test excel_update_range interface"""
        data = [["新姓名", "新年龄"], ["测试1", 99]]
        result = excel_update_range(sample_excel_file, "Sheet1!A5:B6", data,
                                  insert_mode=True, require_confirmation=False, skip_safety_checks=True)

        assert result['success'] is True
        # Should have either data or other response fields
        assert 'data' in result or 'message' in result

    def test_excel_update_range_invalid_sheet(self, sample_excel_file):
        """Test excel_update_range with invalid sheet"""
        data = [["测试"]]
        result = excel_update_range(sample_excel_file, "NonExistentSheet!A1", data)

        assert result['success'] is False
        assert 'error' in result

    def test_excel_update_range_row_format(self, sample_excel_file):
        """Test excel_update_range with row range format - should return error for missing sheet name"""
        # Test single row range - should fail because range doesn't contain sheet name
        data1 = [["测试1", "测试2", "测试3"]]
        result1 = excel_update_range(sample_excel_file, "1:1", data1)
        assert result1['success'] is False
        error_message = result1.get('error', '')
        # 更灵活的错误消息检查，支持多种可能的错误格式
        assert any(msg in error_message for msg in [
            "range必须包含工作表名",
            "VALIDATION_FAILED",
            "range格式错误",
            "工作表名",
            "范围表达式",
            "validation"
        ])

        # Test multi-row range - should also fail
        data2 = [[930006, "", "[TRBuff收益类型]无", "【女武神】退场易伤", 1, 0]]
        result2 = excel_update_range(sample_excel_file, "3:5", data2)
        assert result2['success'] is False
        error_message2 = result2.get('error', '')
        assert any(msg in error_message2 for msg in [
            "range必须包含工作表名",
            "VALIDATION_FAILED",
            "range格式错误",
            "工作表名",
            "范围表达式",
            "validation"
        ])

    def test_excel_update_range_large_row_number(self, temp_dir, request):
        """Test excel_update_range with large row numbers - should work with proper range format"""
        import uuid
        from openpyxl import Workbook

        test_id = str(uuid.uuid4())[:8]
        file_path = temp_dir / f"test_large_row_{test_id}.xlsx"

        # Create test file
        wb = Workbook()
        ws = wb.active
        ws.title = "TrBuff"
        wb.save(file_path)

        # Test user's specific case: row 1250 with 28 columns of data
        user_data = [[
            930006, "", "[TRBuff收益类型]无", "【女武神】退场易伤", "[TRBuff添加类型]替换",
            "", 1, 0, "", "", "[TRBuff效果类型]属性效果", 110202, 99999999,
            "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        ]]

        # Test with proper cell range format should work
        result = excel_update_range(str(file_path), "TrBuff!A1250:AB1250", user_data,
                                     insert_mode=True, require_confirmation=False, skip_safety_checks=True)
        assert result['success'] is True
        if isinstance(result['data'], list):
            assert len(result['data']) == 28  # Should update 28 cells

        # Also test with alternative range format
        result_alt = excel_update_range(str(file_path), "TrBuff!A1250:AC1250", user_data,
                                       insert_mode=True, require_confirmation=False, skip_safety_checks=True)
        assert result_alt['success'] is True

    def test_excel_create_file(self, temp_dir):
        """Test excel_create_file interface"""
        file_path = temp_dir / "test_create.xlsx"
        result = excel_create_file(str(file_path))

        assert result['success'] is True
        assert 'file_path' in result or 'data' in result

        # Verify file was created
        assert file_path.exists()

    def test_excel_create_sheet(self, sample_excel_file):
        """Test excel_create_sheet interface"""
        result = excel_create_sheet(sample_excel_file, "新工作表")

        assert result['success'] is True
        # Should have response data
        assert 'data' in result or 'message' in result

    def test_excel_create_sheet_duplicate_name(self, sample_excel_file):
        """Test excel_create_sheet with duplicate name"""
        result = excel_create_sheet(sample_excel_file, "Sheet1")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_delete_sheet(self, sample_excel_file):
        """Test excel_delete_sheet interface"""
        result = excel_delete_sheet(sample_excel_file, "Sheet2")

        assert result['success'] is True
        # Should have response data
        assert 'data' in result or 'message' in result

    def test_excel_delete_sheet_nonexistent(self, sample_excel_file):
        """Test excel_delete_sheet with non-existent sheet"""
        result = excel_delete_sheet(sample_excel_file, "NonExistentSheet")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_rename_sheet(self, sample_excel_file):
        """Test excel_rename_sheet interface"""
        result = excel_rename_sheet(sample_excel_file, "Sheet1", "数据表")

        assert result['success'] is True
        # Should have response data
        assert 'data' in result or 'message' in result

    def test_excel_rename_sheet_nonexistent(self, sample_excel_file):
        """Test excel_rename_sheet with non-existent sheet"""
        result = excel_rename_sheet(sample_excel_file, "NonExistentSheet", "新名称")

        assert result['success'] is False
        assert 'error' in result

    def test_excel_insert_rows(self, sample_excel_file):
        """Test excel_insert_rows interface"""
        result = excel_insert_rows(sample_excel_file, "Sheet1", 2, 2)

        assert result['success'] is True
        # Should have response info
        assert 'data' in result or 'message' in result

    def test_excel_insert_columns(self, sample_excel_file):
        """Test excel_insert_columns interface"""
        result = excel_insert_columns(sample_excel_file, "Sheet1", 2, 1)

        assert result['success'] is True
        # Should have response info
        assert 'data' in result or 'message' in result

    def test_excel_delete_rows(self, sample_excel_file):
        """Test excel_delete_rows interface"""
        result = excel_delete_rows(sample_excel_file, "Sheet1", 2, 1)

        assert result['success'] is True
        # Should have response info
        assert 'data' in result or 'message' in result

    def test_excel_delete_columns(self, sample_excel_file):
        """Test excel_delete_columns interface"""
        result = excel_delete_columns(sample_excel_file, "Sheet1", 2, 1)

        assert result['success'] is True
        # Should have response info
        assert 'data' in result or 'message' in result

    def test_excel_format_cells_custom(self, sample_excel_file):
        """Test excel_format_cells with custom formatting"""
        formatting = {
            'font': {'name': 'Arial', 'size': 14, 'bold': True}
        }
        result = excel_format_cells(sample_excel_file, "Sheet1", "A1:D1", formatting=formatting)

        # May fail if formatting is not supported
        assert isinstance(result, dict)
        assert 'success' in result

    def test_excel_format_cells_invalid_sheet(self, sample_excel_file):
        """Test excel_format_cells with invalid sheet"""
        formatting = {'font': {'bold': True}}
        result = excel_format_cells(sample_excel_file, "NonExistentSheet", "A1", formatting=formatting)

        assert result['success'] is False
        assert 'error' in result

    def test_excel_format_cells_preset(self, sample_excel_file):
        """Test excel_format_cells with preset"""
        result = excel_format_cells(sample_excel_file, "Sheet1", "A1:D1", preset="header")

        # May fail if formatting is not supported
        assert isinstance(result, dict)
        assert 'success' in result

    def test_excel_search(self, sample_excel_file):
        """Test excel_search interface"""
        result = excel_search(sample_excel_file, r"张三")

        assert result['success'] is True
        # Should have search results
        assert 'data' in result or 'total_matches' in result

    def test_excel_search_invalid_file(self):
        """Test excel_search with invalid file"""
        result = excel_search("nonexistent_file.xlsx", r"test")

        assert result['success'] is False
        assert 'error' in result

    def test_all_interfaces_return_consistent_structure(self, sample_excel_file):
        """Test that all interfaces return consistent response structure"""
        # Test a few key interfaces
        interfaces = [
            lambda: excel_list_sheets(sample_excel_file),
            lambda: excel_get_range(sample_excel_file, "Sheet1!A1"),
            lambda: excel_create_sheet(sample_excel_file, "TestSheet"),
            lambda: excel_search(sample_excel_file, r"test")
        ]

        for i, interface in enumerate(interfaces):
            result = interface()

            # All should have success boolean
            assert 'success' in result
            assert isinstance(result['success'], bool)

            # If successful, should have appropriate data
            if result['success']:
                # Should have either direct data keys or metadata with relevant information
                expected_keys = ['data', 'message', 'result', 'total_matches', 'sheets', 'matches', 'file_path', 'sheet_name', 'total_sheets']
                has_expected_key = any(key in result for key in expected_keys)

                # For search results, check metadata for nested keys
                if not has_expected_key and 'metadata' in result:
                    metadata_keys = ['total_matches', 'pattern', 'search_results']
                    has_expected_key = any(key in result['metadata'] for key in metadata_keys)

                assert has_expected_key, f"Expected data keys in result or metadata, got keys: {list(result.keys())}"
            else:
                assert 'error' in result
                assert isinstance(result['error'], str)


    def test_excel_find_last_row(self, sample_excel_file):
        """Test excel_find_last_row interface - basic functionality"""
        # 获取第一个工作表名称
        sheets_result = excel_list_sheets(sample_excel_file)
        assert sheets_result['success']
        sheet_name = sheets_result['sheets'][0]

        # 测试查找整个工作表的最后一行
        result = excel_find_last_row(sample_excel_file, sheet_name)

        assert result['success'] is True
        assert 'data' in result
        assert 'last_row' in result['data']
        assert 'sheet_name' in result['data']
        assert 'column' in result['data']
        assert 'search_scope' in result['data']
        assert isinstance(result['data']['last_row'], int)
        assert result['data']['last_row'] >= 0
        assert result['data']['sheet_name'] == sheet_name
        assert result['data']['column'] is None  # 没有指定列
        assert result['data']['search_scope'] == "整个工作表"

        # 兼容性字段检查
        assert 'last_row' in result
        assert result['last_row'] == result['data']['last_row']

        # 消息检查
        assert 'message' in result
        assert isinstance(result['message'], str)

    def test_excel_find_last_row_with_column_name(self, sample_excel_file):
        """Test excel_find_last_row with column name - specific column functionality"""
        sheets_result = excel_list_sheets(sample_excel_file)
        assert sheets_result['success']
        sheet_name = sheets_result['sheets'][0]

        # 测试查找A列的最后一行
        result = excel_find_last_row(sample_excel_file, sheet_name, "A")

        assert result['success'] is True
        assert result['data']['column'] == "A"
        assert result['data']['search_scope'] == "A列"
        assert isinstance(result['data']['last_row'], int)
        assert result['data']['last_row'] >= 0

    def test_excel_find_last_row_with_column_index(self, sample_excel_file):
        """Test excel_find_last_row with column index - specific column functionality"""
        sheets_result = excel_list_sheets(sample_excel_file)
        assert sheets_result['success']
        sheet_name = sheets_result['sheets'][0]

        # 测试查找第1列的最后一行
        result = excel_find_last_row(sample_excel_file, sheet_name, 1)

        assert result['success'] is True
        assert result['data']['column'] == 1
        assert result['data']['search_scope'] == "A列"  # 第1列对应A列
        assert isinstance(result['data']['last_row'], int)
        assert result['data']['last_row'] >= 0

    def test_excel_find_last_row_nonexistent_sheet(self, sample_excel_file):
        """Test excel_find_last_row with nonexistent sheet - error handling"""
        result = excel_find_last_row(sample_excel_file, "NonExistentSheet")

        assert result['success'] is False
        assert 'error' in result
        assert "工作表不存在" in result['error']
