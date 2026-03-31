#!/usr/bin/env python3
"""
Script to reproduce the 5 API issues found by supervisor
"""

import sys
import os
import json
import tempfile
from pathlib import Path

# Add src to path
sys.path.insert(0, 'src')

def create_test_excel():
    """Create a test Excel file for reproduction"""
    import pandas as pd
    
    # Create test data
    data = {
        'A': [1, 2, 3, 4, 5],
        'B': ['a', 'b', 'c', 'd', 'e'],
        'C': [10.5, 20.5, 30.5, 40.5, 50.5]
    }
    df = pd.DataFrame(data)
    
    # Create temp file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        df.to_excel(f.name, index=False)
        return f.name

def test_api_issues():
    """Test the 5 specific API issues"""
    test_file = create_test_excel()
    
    print("=== Testing 5 API Issues ===\n")
    
    # Test 1: read_data_from_excel range query parameter order
    print("1. Testing read_data_from_excel range query...")
    try:
        from excel_mcp_server_fastmcp.api.excel_operations import read_data_from_excel
        
        # Test with reversed parameter order
        result1 = read_data_from_excel(test_file, "Sheet1", "C1", "A3")
        print(f"   ✓ Reversed order works: {len(result1)} cells")
        
        result2 = read_data_from_excel(test_file, "Sheet1", "A1", "C3") 
        print(f"   ✓ Normal order works: {len(result2)} cells")
        
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    # Test 2: format_range missing bold parameter
    print("\n2. Testing format_range missing bold parameter...")
    try:
        from excel_mcp_server_fastmcp.api.excel_operations import format_range
        
        # Test without bold parameter
        result = format_range(test_file, "Sheet1", "A1", "B1", bold=True)
        print("   ✓ format_range with bold=True works")
        
        # Test without bold (should still work)
        result = format_range(test_file, "Sheet1", "A2", "B2")
        print("   ✓ format_range without explicit bold works")
        
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    # Test 3: apply_formula missing formula parameter
    print("\n3. Testing apply_formula missing formula parameter...")
    try:
        from excel_mcp_server_fastmcp.api.excel_operations import apply_formula
        
        # This should fail gracefully
        result = apply_formula(test_file, "Sheet1", "A1", "")
        print("   ✓ apply_formula with empty formula works")
        
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    # Test 4: read_data_from_excel search logic
    print("\n4. Testing read_data_from_excel search logic...")
    try:
        from excel_mcp_server_fastmcp.api.excel_operations import read_data_from_excel
        
        # Test different search patterns
        result = read_data_from_excel(test_file, "Sheet1", search_value="3")
        print(f"   ✓ Search by value works: {len(result)} results")
        
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    # Test 5: write_data_to_excel format mismatch
    print("\n5. Testing write_data_to_excel format mismatch...")
    try:
        from excel_mcp_server_fastmcp.api.excel_operations import write_data_to_excel
        
        # Test with properly formatted data
        data = [['x', 'y', 'z'], [1, 2, 3]]
        result = write_data_to_excel(test_file, "Sheet1", data, start_cell="D1")
        print("   ✓ write_data_to_excel with proper format works")
        
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    # Cleanup
    os.unlink(test_file)
    print("\n=== Test completed ===")

if __name__ == "__main__":
    test_api_issues()