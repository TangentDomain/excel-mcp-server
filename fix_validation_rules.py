#!/usr/bin/env python3
"""
REQ-066-3: 修复验证规则创建逻辑中的关键问题
修复内容：
1. DataValidation对象创建前的参数验证
2. 修复操作符和公式参数的映射关系
3. 优化验证条件的解析逻辑
"""

import re
import sys
import os

def fix_data_validation_issues():
    """修复数据验证规则创建逻辑中的关键问题"""
    
    file_path = "src/excel_mcp_server_fastmcp/server.py"
    
    # 读取原文件
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 问题1: 修复操作符映射
    old_operator_mapping = """
        type_mapping = {
            'whole_number': 'whole',
            'text_length': 'textLength',
        }
        openpyxl_type = type_mapping.get(validation_type, validation_type)"""
    
    new_operator_mapping = """
        type_mapping = {
            'whole_number': 'whole',
            'text_length': 'textLength',
            'decimal': 'decimal',
            'date': 'date',
            'list': 'list',
            'custom': 'custom'
        }
        openpyxl_type = type_mapping.get(validation_type, validation_type)
        
        # 验证映射结果
        if not openpyxl_type:
            logger.error(f"[DATA_VALIDATION] 验证类型映射失败 - validation_type={validation_type}")
            return _fail(f"不支持的验证类型: {validation_type}",
                        meta={"error_code": "VALIDATION_FAILED"})"""
    
    # 问题2: 修复验证条件解析
    old_criteria_parsing = """                # 验证操作符
                valid_operators = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan',
                                 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
                logger.debug(f"[DATA_VALIDATION][{validation_type.upper()}] 开始验证操作符 - operator='{operator}', valid_operators={valid_operators}")
                if operator not in valid_operators:
                    logger.error(f"[DATA_VALIDATION][{validation_type.upper()}] 操作符验证失败 - operator='{operator}' not in valid_operators")
                    return _fail(f"不支持的操作符: {operator}，支持的操作符: {','.join(valid_operators)}",
                                meta={"error_code": "VALIDATION_FAILED"})
                logger.info(f"[DATA_VALIDATION][{validation_type.upper()}] 操作符验证通过 - operator='{operator}'")"""
    
    new_criteria_parsing = """                # 验证操作符 - 根据验证类型动态验证
                type_operators = {
                    'whole_number': ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
                    'decimal': ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
                    'date': ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
                    'text_length': ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
                    'list': [],
                    'custom': []
                }
                
                valid_operators = type_operators.get(validation_type, [])
                if validation_type in ['list', 'custom']:
                    # 对于list和custom类型，不需要操作符验证
                    operator = None
                    value1 = criteria
                    value2 = None
                    logger.info(f"[DATA_VALIDATION][{validation_type.upper()}] 使用简单模式 - criteria='{criteria}'")
                else:
                    # 对于数值类型，验证操作符
                    if operator not in valid_operators:
                        logger.error(f"[DATA_VALIDATION][{validation_type.upper()}] 操作符验证失败 - operator='{operator}' not in valid_operators={valid_operators}")
                        return _fail(f"不支持的操作符: {operator}，支持的操作符: {','.join(valid_operators)}",
                                    meta={"error_code": "VALIDATION_FAILED"})
                    logger.info(f"[DATA_VALIDATION][{validation_type.upper()}] 操作符验证通过 - operator='{operator}'")"""
    
    # 问题3: 修复DataValidation对象创建参数
    old_dv_creation = """        logger.info(f"[DATA_VALIDATION] 创建 DataValidation 对象 - dv_kwargs={dv_kwargs}")
        dv = DataValidation(**dv_kwargs)"""
    
    new_dv_creation = """        logger.info(f"[DATA_VALIDATION] 创建 DataValidation 对象 - dv_kwargs={dv_kwargs}")
        
        # 清理和验证参数
        cleaned_kwargs = {}
        for key, value in dv_kwargs.items():
            if value is not None:
                cleaned_kwargs[key] = value
        
        try:
            dv = DataValidation(**cleaned_kwargs)
            logger.info(f"[DATA_VALIDATION] DataValidation 对象创建成功 - type='{dv.type}', operator='{getattr(dv, 'operator', 'N/A')}'")
        except Exception as e:
            logger.error(f"[DATA_VALIDATION] DataValidation 对象创建失败 - error={str(e)}, kwargs={cleaned_kwargs}")
            return _fail(f"验证规则创建失败: {str(e)}", meta={"error_code": "VALIDATION_FAILED"})"""
    
    # 应用修复
    content = content.replace(old_operator_mapping, new_operator_mapping)
    content = content.replace(old_criteria_parsing, new_criteria_parsing)
    content = content.replace(old_dv_creation, new_dv_creation)
    
    # 写回文件
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("✅ 修复完成")
    print("修复内容:")
    print("1. 修复了验证类型映射问题")
    print("2. 优化了验证条件解析逻辑")
    print("3. 改进了DataValidation对象创建的错误处理")

if __name__ == "__main__":
    fix_data_validation_issues()
