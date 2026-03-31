#!/usr/bin/env python3
"""
Test the range_boundaries function to understand the issue with reversed ranges
"""

from openpyxl.utils import range_boundaries

def test_range_boundaries():
    """Test how range_boundaries handles different range formats"""
    
    test_ranges = [
        "A1:C3",      # Normal range
        "C3:A1",      # Reversed range
        "B2:D4",      # Another normal range
        "D4:B2",      # Another reversed range
    ]
    
    print("Testing range_boundaries function:")
    print("=" * 50)
    
    for range_expr in test_ranges:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_expr)
            print(f"'{range_expr}' -> min_col:{min_col}, min_row:{min_row}, max_col:{max_col}, max_row:{max_row}")
            
            # Check if the range is properly ordered
            if min_col > max_col or min_row > max_row:
                print(f"  ⚠️  Range is not properly ordered!")
            
        except Exception as e:
            print(f"'{range_expr}' -> ERROR: {e}")
        
        print()

if __name__ == "__main__":
    test_range_boundaries()