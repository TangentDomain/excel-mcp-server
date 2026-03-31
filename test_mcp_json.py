#!/usr/bin/env python3
"""
Test MCP server JSON response formatting
"""

import json
import tempfile
import os
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, 'src')

def create_test_file():
    """Create a simple test Excel file"""
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        # Add some test data
        headers = ["ID", "Name", "Value"]
        test_data = [
            headers,
            [1, "Test1", 100],
            [2, "Test2", 200]
        ]
        
        for row_idx, row_data in enumerate(test_data):
            for col_idx, cell_value in enumerate(row_data):
                ws.cell(row=row_idx+1, column=col_idx+1, value=cell_value)
        
        test_file = "/tmp/mcp_test.xlsx"
        wb.save(test_file)
        return test_file
    except ImportError:
        print("❌ openpyxl not available")
        return None

def test_mcp_json_responses():
    """Test MCP functions for JSON response issues"""
    test_file = create_test_file()
    if not test_file:
        return False
    
    print("🧪 Testing MCP JSON Response Formatting")
    print("=" * 60)
    
    # Test MCP functions that should return JSON-serializable results
    test_functions = [
        ("excel_list_sheets", test_file),
        ("excel_get_range", test_file, "A1:C5"),
        ("excel_search", test_file, "Test"),
        ("excel_get_headers", test_file),
        ("excel_describe_table", test_file),
    ]
    
    all_passed = True
    
    for func_info in test_functions:
        func_name = func_info[0]
        args = func_info[1:]
        
        print(f"\n🔍 Testing {func_name}...")
        
        try:
            # Import and call the function
            if func_name == "excel_list_sheets":
                from excel_mcp_server_fastmcp.server import excel_list_sheets
                result = excel_list_sheets(*args)
            elif func_name == "excel_get_range":
                from excel_mcp_server_fastmcp.server import excel_get_range
                result = excel_get_range(*args)
            elif func_name == "excel_search":
                from excel_mcp_server_fastmcp.server import excel_search
                result = excel_search(*args)
            elif func_name == "excel_get_headers":
                from excel_mcp_server_fastmcp.server import excel_get_headers
                result = excel_get_headers(*args)
            elif func_name == "excel_describe_table":
                from excel_mcp_server_fastmcp.server import excel_describe_table
                result = excel_describe_table(*args)
            else:
                print(f"❌ Unknown function: {func_name}")
                all_passed = False
                continue
            
            # Check if result is JSON serializable
            try:
                json_str = json.dumps(result, ensure_ascii=False, indent=2)
                print(f"✅ {func_name}: JSON serialization successful")
                
                # Check for trailing characters issue
                json_str_clean = json_str.strip()
                if json_str_clean:
                    # Try to parse it back
                    parsed_back = json.loads(json_str_clean)
                    print(f"✅ {func_name}: JSON round-trip successful")
                else:
                    print(f"⚠️ {func_name}: Empty JSON response")
                    
            except json.JSONDecodeError as e:
                print(f"❌ {func_name}: JSON decode error - {e}")
                print(f"   Raw output: {repr(result)}")
                all_passed = False
                
            except TypeError as e:
                print(f"❌ {func_name}: JSON serialization error - {e}")
                print(f"   Result type: {type(result)}")
                all_passed = False
                
        except Exception as e:
            print(f"❌ {func_name}: Function call error - {e}")
            all_passed = False
    
    # Cleanup
    if os.path.exists(test_file):
        os.unlink(test_file)
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 All MCP functions passed JSON formatting tests!")
        return True
    else:
        print("⚠️ Some MCP functions failed JSON formatting tests")
        return False

if __name__ == "__main__":
    success = test_mcp_json_responses()
    sys.exit(0 if success else 1)