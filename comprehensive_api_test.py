#!/usr/bin/env python3
"""
Test script to reproduce and fix the 5 API problems found by supervisor.
"""

import json
import sys
import os
import tempfile
from pathlib import Path

# Add src to path
sys.path.insert(0, 'src')

def create_test_excel():
    """Create a simple test Excel file"""
    from openpyxl import Workbook
    
    test_file = "/tmp/test_api_fix.xlsx"
    if os.path.exists(test_file):
        os.remove(test_file)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Add headers and data
    headers = ["Name", "Age", "City", "Salary"]
    data = [
        ["Alice", 25, "Shenzhen", 10000],
        ["Bob", 30, "Beijing", 12000],
        ["Charlie", 35, "Shanghai", 15000]
    ]
    
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, cell_value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=cell_value)
    
    wb.save(test_file)
    return test_file

def test_supervisor_issues():
    """Test the 5 issues mentioned by supervisor"""
    print("=== REPRODUCING SUPERVISOR API ISSUES ===\n")
    
    test_file = create_test_excel()
    
    try:
        from excel_mcp_server_fastmcp.server import (
            excel_get_range, 
            excel_format_cells, 
            excel_set_formula,
            excel_search,
            excel_update_range
        )
        
        print("1. Testing excel_get_range with start_cell/end_cell parameters (SHOULD FAIL)")
        try:
            # This should fail because excel_get_range expects 'range', not 'start_cell'/'end_cell'
            result = excel_get_range(
                file_path=test_file,
                sheet_name="Sheet1",
                start_cell="A1",  # ❌ Wrong parameter
                end_cell="C3"     # ❌ Wrong parameter
            )
            print(f"❌ UNEXPECTED SUCCESS: {result}")
        except TypeError as e:
            print(f"✅ EXPECTED FAILURE: {e}")
        
        print("\n2. Testing excel_format_cells without required formatting parameters")
        try:
            result = excel_format_cells(
                file_path=test_file,
                sheet_name="Sheet1",
                start_cell="A1"  # ❌ Should be 'range', and missing formatting dict
            )
            print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"Exception: {e}")
        
        print("\n3. Testing excel_set_formula without formula parameter")
        try:
            result = excel_set_formula(
                file_path=test_file,
                sheet_name="Sheet1",
                cell_address="A1"  # ❌ Missing required 'formula'
            )
            print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"✅ EXPECTED FAILURE (missing formula): {e}")
        
        print("\n4. Testing excel_search (this should work)")
        try:
            result = excel_search(
                file_path=test_file,
                pattern="Alice",
                sheet_name="Sheet1"
            )
            print(f"✅ SEARCH WORKS: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ Search failed: {e}")
        
        print("\n5. Testing excel_update_range with start_cell instead of range")
        try:
            result = excel_update_range(
                file_path=test_file,
                sheet_name="Sheet1",  # ❌ This parameter doesn't exist
                start_cell="A4",      # ❌ Should be 'range'
                data=[["David", 40, "Hangzhou", 18000]]
            )
            print(f"Result: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ Expected failure: {e}")
    
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)

def test_correct_usage():
    """Test the correct usage patterns that should work"""
    print("\n\n=== TESTING CORRECT USAGE PATTERNS ===\n")
    
    test_file = create_test_excel()
    
    try:
        from excel_mcp_server_fastmcp.server import (
            excel_get_range, 
            excel_format_cells, 
            excel_set_formula,
            excel_search,
            excel_update_range
        )
        
        print("1. Testing excel_get_range with correct range parameter")
        try:
            result = excel_get_range(
                file_path=test_file,
                range="Sheet1!A1:C3"  # ✅ Correct format
            )
            print(f"✅ excel_get_range works: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ excel_get_range failed: {e}")
        
        print("\n2. Testing excel_format_cells with correct parameters")
        try:
            result = excel_format_cells(
                file_path=test_file,
                sheet_name="Sheet1",
                range="A1:C1",
                formatting={"bold": True, "font_color": "FF0000"}  # ✅ Required formatting dict
            )
            print(f"✅ excel_format_cells works: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ excel_format_cells failed: {e}")
        
        print("\n3. Testing excel_set_formula with correct parameters")
        try:
            result = excel_set_formula(
                file_path=test_file,
                sheet_name="Sheet1",
                cell_address="D4",
                formula="=SUM(A4:C4)"  # ✅ Required formula
            )
            print(f"✅ excel_set_formula works: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ excel_set_formula failed: {e}")
        
        print("\n4. Testing excel_update_range with correct parameters")
        try:
            result = excel_update_range(
                file_path=test_file,
                range="A4:C4",  # ✅ Correct range parameter
                data=[["Eve", 28, "Guangzhou", 16000]]
            )
            print(f"✅ excel_update_range works: {json.dumps(result, indent=2, ensure_ascii=False)}")
        except Exception as e:
            print(f"❌ excel_update_range failed: {e}")
    
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)

if __name__ == "__main__":
    test_supervisor_issues()
    test_correct_usage()