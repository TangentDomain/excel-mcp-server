import ast
import os
import sys
from pathlib import Path

def get_function_signatures(file_path):
    """获取所有public函数的参数签名"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    tree = ast.parse(content)
    functions = {}
    
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
            params = []
            for arg in node.args.args:
                if arg.arg != 'self':  # 跳过self参数
                    params.append(arg.arg)
            functions[node.name] = params
    
    return functions

def get_docstring_params(file_path):
    """解析docstring中的参数"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    tree = ast.parse(content)
    docstring_params = {}
    
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and not node.name.startswith('_'):
            if (node.body and 
                isinstance(node.body[0], ast.Expr) and 
                isinstance(node.body[0].value, ast.Constant) and 
                isinstance(node.body[0].value.value, str)):
                
                docstring = node.body[0].value.value
                # 多种格式解析
                params = []
                
                # 格式1: Args: 段落
                if 'Args:' in docstring:
                    args_section = docstring.split('Args:')[1]
                    # 找到下一个段落的开始
                    if '\n\n' in args_section:
                        args_section = args_section.split('\n\n')[0]
                    elif 'Returns:' in args_section:
                        args_section = args_section.split('Returns:')[0]
                    lines = args_section.strip().split('\n')
                    for line in lines:
                        line = line.strip()
                        if line and (': ' in line or ':' in line):
                            param_name = line.split(':')[0].strip()
                            if param_name and param_name != 'Args':
                                params.append(param_name)
                
                # 格式2: Parameters: 段落
                elif 'Parameters:' in docstring:
                    params_section = docstring.split('Parameters:')[1]
                    if '\n\n' in params_section:
                        params_section = params_section.split('\n\n')[0]
                    lines = params_section.strip().split('\n')
                    for line in lines:
                        line = line.strip()
                        if line and ': ' in line:
                            param_name = line.split(':')[0].strip()
                            if param_name and param_name != 'Parameters':
                                params.append(param_name)
                
                # 格式3: 简单的参数列表（第一个冒号之后的所有参数行）
                else:
                    # 尝试找到参数描述部分
                    lines = docstring.split('\n')
                    for i, line in enumerate(lines):
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            # 这可能是参数行，检查后续几行
                            for j in range(i, min(i+5, len(lines))):
                                next_line = lines[j].strip()
                                if ': ' in next_line and not next_line.startswith('#'):
                                    param_name = next_line.split(':')[0].strip()
                                    if param_name and param_name not in params:
                                        params.append(param_name)
                    
                    # 如果还是没找到，尝试从函数文档中提取
                    if not params and node.body and len(node.body) > 1:
                        # 查找第一个真正的文档字符串段
                        for doc_line in node.body:
                            if isinstance(doc_line, ast.Expr) and isinstance(doc_line.value, ast.Constant):
                                doc_text = str(doc_line.value.value)
                                # 尝试从文档中提取参数名（更简单的启发式方法）
                                words = doc_text.split()
                                for word in words:
                                    if (word.endswith(':') and len(word) > 1 and 
                                        word.islower() and word not in ['args', 'parameters']):
                                        clean_word = word.rstrip(':')
                                        if clean_word not in params:
                                            params.append(clean_word)
                
                if params:
                    docstring_params[node.name] = params
    
    return docstring_params

def lint_docstring_contract(src_dir):
    """验证docstring契约"""
    errors = []
    
    for py_file in Path(src_dir).rglob('*.py'):
        if '__pycache__' in str(py_file):
            continue
            
        functions = get_function_signatures(py_file)
        docstring_params = get_docstring_params(py_file)
        
        for func_name, signature_params in functions.items():
            if func_name in docstring_params:
                doc_params = docstring_params[func_name]
                
                # 检查docstring中缺失的参数
                for param in signature_params:
                    if param not in doc_params:
                        errors.append(f"{py_file}:{func_name} | {param} | docstring中缺失参数")
                
                # 检查默认值一致性（需要进一步实现）
            else:
                errors.append(f"{py_file}:{func_name} | 所有参数 | 函数有docstring但缺少Args段")
    
    return errors

if __name__ == "__main__":
    src_dir = "src"
    errors = lint_docstring_contract(src_dir)
    
    if errors:
        for error in errors:
            print(error)
        sys.exit(1)
    else:
        print("所有函数的docstring契约验证通过")
        sys.exit(0)