#!/usr/bin/env python3
"""
Code quality analysis and improvement script
"""
import ast
import os
from typing import List, Dict, Any

def analyze_file(filepath: str) -> Dict[str, Any]:
    """Analyze a Python file for code quality issues"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            tree = ast.parse(content)
            
            # Count functions/classes
            functions = sum(1 for node in ast.walk(tree) if isinstance(node, ast.FunctionDef))
            classes = sum(1 for node in ast.walk(tree) if isinstance(node, ast.ClassDef))
            
            # Check for docstrings (fixed for Python 3.14+ compatibility)
            has_docstring = any(
                node.body and isinstance(node.body[0], ast.Expr) and 
                isinstance(node.body[0].value, (ast.Constant, ast.Str))
                for node in ast.walk(tree) 
                if isinstance(node, (ast.FunctionDef, ast.ClassDef))
            )
            
            # Look for potential improvements
            improvements = []
            
            # Check for hardcoded paths
            if '"/' in content or "'/" in content:
                improvements.append("Hardcoded paths detected")
            
            # Check for TODO/FIXME
            if 'TODO' in content or 'FIXME' in content:
                improvements.append("TODO/FIXME comments found")
            
            # Check for long functions (>50 lines)
            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef):
                    # Count lines in function
                    line_count = node.end_lineno - node.lineno + 1 if hasattr(node, 'end_lineno') else 0
                    if line_count > 50:
                        improvements.append(f"Long function: {node.name} ({line_count} lines)")
            
            return {
                'functions': functions,
                'classes': classes,
                'has_docstring': has_docstring,
                'improvements': improvements,
                'lines': len(content.split('\n'))
            }
    except Exception as e:
        return {'error': str(e)}

def main():
    """Main analysis function"""
    print("🔍 Code Quality Analysis")
    print("=" * 50)
    
    all_improvements = []
    
    for root, dirs, files in os.walk('src/excel_mcp_server_fastmcp'):
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                result = analyze_file(filepath)
                
                if 'error' in result:
                    print(f"❌ {file}: {result['error']}")
                    continue
                
                print(f"📄 {file}: {result['functions']} functions, {result['classes']} classes, {result['lines']} lines")
                if result['has_docstring']:
                    print(f"   ✅ Has docstrings")
                else:
                    print(f"   ❌ Missing docstrings")
                
                if result['improvements']:
                    print(f"   ⚠️  Improvements needed:")
                    for improvement in result['improvements']:
                        print(f"      - {improvement}")
                        all_improvements.append(f"{file}: {improvement}")
    
    if all_improvements:
        print("\n🎯 Summary of improvements needed:")
        for improvement in all_improvements:
            print(f"   - {improvement}")
    else:
        print("\n✅ No major code quality issues found!")

if __name__ == "__main__":
    main()