#!/usr/bin/env python3
"""
Reproduce the 5 API issues found by supervisor
"""
import subprocess
import json
import os
import sys

def run_mcp_command(tool_name, params):
    """Run MCP command and return result"""
    try:
        # Build the command
        cmd = ["python3", "-m", "excel_mcp_server_fastmcp"]
        
        # Add tool parameters
        if params:
            for key, value in params.items():
                cmd.extend([f"--{key}", str(value)])
        
        # Run the command
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        return {
            "success": result.returncode == 0,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "returncode": result.returncode
        }
    except subprocess.TimeoutExpired:
        return {
            "success": False,
            "error": "Timeout",
            "stdout": "",
            "stderr": "Command timed out after 30 seconds"
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "stdout": "",
            "stderr": str(e)
        }

def test_api_issues():
    """Test the 5 API issues found by supervisor"""
    
    print("🔍 Testing 5 API issues found by supervisor...")
    print("=" * 60)
    
    # Issue 1: read_data_from_excel range query - parameter order might be reversed
    print("\n1. Testing read_data_from_excel range query...")
    
    # Create test file
    test_file = "/tmp/test_range.xlsx"
    write_data = [
        ["Name", "Age", "City"],
        ["Alice", "25", "New York"],
        ["Bob", "30", "Los Angeles"],
        ["Charlie", "35", "Chicago"]
    ]
    
    # Create test Excel file
    import pandas as pd
    df = pd.DataFrame(write_data[1:], columns=write_data[0])
    df.to_excel(test_file, index=False)
    
    # Test with start_cell, end_cell order
    result1 = run_mcp_command("read_data_from_excel", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "start_cell": "A1",
        "end_cell": "C3"
    })
    
    print(f"   Result: {'✅ PASS' if result1['success'] else '❌ FAIL'}")
    if not result1['success']:
        print(f"   Error: {result1['stderr']}")
    
    # Test with reversed end_cell, start_cell order (potential issue)
    result2 = run_mcp_command("read_data_from_excel", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "end_cell": "C3",
        "start_cell": "A1"
    })
    
    print(f"   Reversed order: {'✅ PASS' if result2['success'] else '❌ FAIL'}")
    if not result2['success']:
        print(f"   Error: {result2['stderr']}")
    
    # Issue 2: format_range - missing parameters handling
    print("\n2. Testing format_range missing parameters...")
    
    result3 = run_mcp_command("format_range", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "start_cell": "A1",
        "end_cell": "A1"
        # Missing bold parameter
    })
    
    print(f"   Missing bold param: {'✅ PASS' if result3['success'] else '❌ FAIL'}")
    if not result3['success']:
        print(f"   Error: {result3['stderr']}")
    
    # Issue 3: apply_formula - missing formula parameter
    print("\n3. Testing apply_formula missing formula...")
    
    result4 = run_mcp_command("apply_formula", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "cell": "A4"
        # Missing formula parameter
    })
    
    print(f"   Missing formula param: {'✅ PASS' if result4['success'] else '❌ FAIL'}")
    if not result4['success']:
        print(f"   Error: {result4['stderr']}")
    
    # Issue 4: read_data_from_excel search logic - sheet_name confusion
    print("\n4. Testing read_data_from_excel search with sheet_name...")
    
    result5 = run_mcp_command("read_data_from_excel", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "start_cell": "A1",
        "end_cell": "C3"
    })
    
    print(f"   Search logic: {'✅ PASS' if result5['success'] else '❌ FAIL'}")
    if not result5['success']:
        print(f"   Error: {result5['stderr']}")
    
    # Issue 5: write_data_to_excel - format mismatch handling
    print("\n5. Testing write_data_to_excel format mismatch...")
    
    # Try with wrong format (not list of lists)
    result6 = run_mcp_command("write_data_to_excel", {
        "file_path": test_file,
        "sheet_name": "Sheet1",
        "data": "not a list",
        "start_cell": "A1"
    })
    
    print(f"   Format mismatch: {'✅ PASS' if result6['success'] else '❌ FAIL'}")
    if not result6['success']:
        print(f"   Error: {result6['stderr']}")
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 Summary of API Issues:")
    issues = [
        ("Range query parameter order", result1['success'] and result2['success']),
        ("format_range missing parameters", result3['success']),
        ("apply_formula missing formula", result4['success']),
        ("read_data_from_excel search logic", result5['success']),
        ("write_data_to_excel format mismatch", result6['success'])
    ]
    
    passed = 0
    for issue_name, success in issues:
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"   {issue_name}: {status}")
        if success:
            passed += 1
    
    print(f"\nOverall: {passed}/5 issues passing")
    
    # Cleanup
    if os.path.exists(test_file):
        os.remove(test_file)
    
    return passed == 5

if __name__ == "__main__":
    success = test_api_issues()
    sys.exit(0 if success else 1)