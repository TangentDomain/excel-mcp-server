#!/usr/bin/env python3
import sys
import os
sys.path.append(os.getcwd())

from src.server import excel_regex_search

print('=== 测试1: 不带范围表达式 ===')
result1 = excel_regex_search('test_range_search.xlsx', '@')
print('数据数量:', len(result1.get('data', [])))

print('\n=== 测试2: 带范围表达式 ===')  
result2 = excel_regex_search('test_range_search.xlsx', '@', range_expression='A1:C6')
print('数据数量:', len(result2.get('data', [])))

print('\n=== 测试3: 手动传递所有参数 ===')
result3 = excel_regex_search(
    file_path='test_range_search.xlsx',
    pattern='@',
    sheet_name=None,
    flags='',
    search_values=True,
    search_formulas=False,
    range_expression='A1:C6'
)
print('数据数量:', len(result3.get('data', [])))
