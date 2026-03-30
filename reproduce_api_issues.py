#!/usr/bin/env python3
"""
Test script to reproduce the 5 API problems mentioned by supervisor.
"""

import json
import sys
import os
sys.path.insert(0, 'src')

from excel_mcp_server_fastmcp.server import excel_get_range, excel_format_cells, excel_set_formula
from excel_mcp_server_fastmcp.api.excel_operations import ExcelOperations
from excel_mcp_server_fastmcp.core.excel_manager import ExcelManager

def test_1_range_query():
    """Test 1: read_data_from_excel range query - parameter order issue"""
    print("=== Test 1: excel_get_range parameter order ===")
    try:
        # Create a test Excel file first
        test_file = "/tmp/test_range_query.xlsx"
        manager = ExcelManager()
        manager.create_workbook(test_file)
        manager.create_sheet(test_file, "Sheet1", ["Name", "Age", "City"], ["Alice", 25, "Shenzhen"])
        manager.save_workbook(test_file)
        
        # Test with correct parameters
        result = excel_get_range(
            file_path=test_file,
            sheet_name="Sheet1",
            start_cell="A1",
            end_cell="C2"
        )
        print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        if result.get('success'):
            print("✅ Range query passed")
        else:
            print(f"❌ Range query failed: {result.get('message')}")
    except Exception as e:
        print(f"❌ Range query exception: {e}")

def test_2_format_range_missing_params():
    """Test 2: format_range missing required parameters"""
    print("\n=== Test 2: excel_format_cells missing parameters ===")
    try:
        test_file = "/tmp/test_format.xlsx"
        manager = ExcelManager()
        manager.create_workbook(test_file)
        manager.create_sheet(test_file, "Sheet1", ["Name", "Age"], ["Bob", 30])
        manager.save_workbook(test_file)
        
        # Test without required parameters like bold
        result = excel_format_cells(
            file_path=test_file,
            sheet_name="Sheet1",
            start_cell="A1"
            # Missing bold parameter
        )
        print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        if result.get('success'):
            print("✅ format_cells missing params passed")
        else:
            print(f"❌ format_cells missing params failed: {result.get('message')}")
    except Exception as e:
        print(f"❌ format_cells missing params exception: {e}")

def test_3_apply_formula_missing_params():
    """Test 3: apply_formula missing required parameters"""
    print("\n=== Test 3: excel_set_formula missing parameters ===")
    try:
        test_file = "/tmp/test_formula.xlsx"
        manager = ExcelManager()
        manager.create_workbook(test_file)
        manager.create_sheet(test_file, "Sheet1", ["A", "B", "C"], [1, 2, 3])
        manager.save_workbook(test_file)
        
        # Test without formula parameter
        result = excel_set_formula(
            file_path=test_file,
            sheet_name="Sheet1",
            cell_address="C1"
            # Missing formula parameter
        )
        print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        if result.get('success'):
            print("✅ set_formula missing params passed")
        else:
            print(f"❌ set_formula missing params failed: {result.get('message')}")
    except Exception as e:
        print(f"❌ set_formula missing params exception: {e}")

def test_4_search_logic():
    """Test 4: read_data_from_excel search logic - parameter confusion"""
    print("\n=== Test 4: excel_search parameter confusion ===")
    try:
        test_file = "/tmp/test_search.xlsx"
        manager = ExcelManager()
        manager.create_workbook(test_file)
        manager.create_sheet(test_file, "Sheet1", ["Name", "Department", "Salary"], 
                            ["Alice", "Engineering", 10000, "Bob", "Sales", 8000])
        manager.save_workbook(test_file)
        
        # Test search with potential parameter confusion
        result = excel_search(
            file_path=test_file,
            pattern="Alice",
            sheet_name="Sheet1",
            case_sensitive=False
        )
        print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        if result.get('success'):
            print("✅ Search passed")
        else:
            print(f"❌ Search failed: {result.get('message')}")
    except Exception as e:
        print(f"❌ Search exception: {e}")

def test_5_write_data_format():
    """Test 5: write_data_to_excel data format mismatch"""
    print("\n=== Test 5: write_data_to_excel format mismatch ===")
    try:
        # Check if there's a write_data function, or use excel_update_range
        test_file = "/tmp/test_write.xlsx"
        manager = ExcelManager()
        manager.create_workbook(test_file)
        manager.create_sheet(test_file, "Sheet1", ["Name", "Age"], ["Charlie", 35])
        manager.save_workbook(test_file)
        
        # Test with potentially mismatched data format
        # Try using the underlying API directly
        result = ExcelOperations.update_range(
            file_path=test_file,
            sheet_name="Sheet1",
            data=[["David", 40]],  # List of lists format
            start_cell="A2",
            # insert_mode missing might cause issues
        )
        print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        if result.get('success'):
            print("✅ Write data passed")
        else:
            print(f"❌ Write data failed: {result.get('message')}")
    except Exception as e:
        print(f"❌ Write data exception: {e}")

if __name__ == "__main__":
    test_1_range_query()
    test_2_format_range_missing_params()
    test_3_apply_formula_missing_params()
    test_4_search_logic()
    test_5_write_data_format()