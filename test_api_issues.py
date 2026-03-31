#!/usr/bin/env python3
"""
Test script to reproduce the 5 API issues identified by supervisor
"""

import json
import subprocess
import sys
import tempfile
import os

def run_mcp_command(tool_name, args):
    """Run an MCP command and return the result"""
    cmd = [
        "uvx", "excel-mcp-server-fastmcp", 
        "stdio"
    ]
    
    # Create MCP request
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": tool_name,
            "arguments": args
        }
    }
    
    try:
        # Start the process
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Send request
        stdout, stderr = process.communicate(input=json.dumps(request))
        
        if process.returncode != 0:
            return {"error": f"Process failed: {stderr}"}
        
        # Parse response
        response = json.loads(stdout)
        if "error" in response:
            return {"error": response["error"]}
        
        return response.get("result", {})
        
    except Exception as e:
        return {"error": str(e)}

def test_issue_1_read_data_range_params():
    """Test 1: read_data_from_excel range query - parameter order may be reversed"""
    print("=== Testing Issue 1: read_data_from_excel range query ===")
    
    # Test with normal parameters
    result1 = run_mcp_command("read_data_from_excel", {
        "filepath": "test_data.xlsx",
        "sheet_name": "Sheet1",
        "start_cell": "A1",
        "end_cell": "C3"
    })
    
    print("Normal parameters (A1 to C3):", result1)
    
    # Test with reversed parameters 
    result2 = run_mcp_command("read_data_from_excel", {
        "filepath": "test_data.xlsx", 
        "sheet_name": "Sheet1",
        "start_cell": "C3",
        "end_cell": "A1"
    })
    
    print("Reversed parameters (C3 to A1):", result2)
    
    return result1, result2

def test_issue_2_format_range_missing_params():
    """Test 2: format_range - missing bold and other required parameters"""
    print("\n=== Testing Issue 2: format_range missing parameters ===")
    
    # Test with minimal parameters
    result = run_mcp_command("format_range", {
        "filepath": "test_data.xlsx",
        "sheet_name": "Sheet1", 
        "start_cell": "A1",
        "end_cell": "C1"
        # Missing bold, font_size, etc.
    })
    
    print("Format range with minimal params:", result)
    return result

def test_issue_3_apply_formula_missing_params():
    """Test 3: apply_formula - missing formula parameter"""
    print("\n=== Testing Issue 3: apply_formula missing formula ===")
    
    # Test without formula parameter
    result = run_mcp_command("apply_formula", {
        "filepath": "test_data.xlsx",
        "sheet_name": "Sheet1",
        "cell": "A1"
        # Missing formula parameter
    })
    
    print("Apply formula without formula param:", result)
    return result

def test_issue_4_read_data_search_confusion():
    """Test 4: read_data_from_excel search logic - may confuse sheet_name with other params"""
    print("\n=== Testing Issue 4: read_data_from_excel search confusion ===")
    
    # Test with search-like parameters
    result = run_mcp_command("read_data_from_excel", {
        "filepath": "test_data.xlsx",
        "sheet_name": "Sheet1",
        "start_cell": "A1",
        "end_cell": "C3",
        # Potentially confusing parameters that might be treated as search
    })
    
    print("Read data with potentially confusing params:", result)
    return result

def test_issue_5_write_data_format_mismatch():
    """Test 5: write_data_to_excel - data format (list of lists) mismatch"""
    print("\n=== Testing Issue 5: write_data_to_excel format mismatch ===")
    
    # Test with wrong data format
    wrong_data = "invalid data format"  # Should be list of lists
    
    result = run_mcp_command("write_data_to_excel", {
        "filepath": "test_data.xlsx",
        "sheet_name": "Sheet1",
        "data": wrong_data,
        "start_cell": "A1"
    })
    
    print("Write data with wrong format:", result)
    return result

def main():
    print("Starting API issue reproduction tests...")
    print("Test file: test_data.xlsx")
    print("=" * 50)
    
    # Test all 5 issues
    results = {}
    
    try:
        results["issue_1"] = test_issue_1_read_data_range_params()
        results["issue_2"] = test_issue_2_format_range_missing_params() 
        results["issue_3"] = test_issue_3_apply_formula_missing_params()
        results["issue_4"] = test_issue_4_read_data_search_confusion()
        results["issue_5"] = test_issue_5_write_data_format_mismatch()
        
    except Exception as e:
        print(f"Error during testing: {e}")
        results["error"] = str(e)
    
    # Summary
    print("\n" + "=" * 50)
    print("TEST SUMMARY:")
    print("=" * 50)
    
    for issue, result in results.items():
        print(f"\n{issue}:")
        if isinstance(result, tuple):
            for i, r in enumerate(result):
                print(f"  Test {i+1}: {r}")
        else:
            print(f"  Result: {result}")
    
    return results

if __name__ == "__main__":
    main()