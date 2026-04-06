#!/usr/bin/env python3
import re
from pathlib import Path

def fix_syntax_errors(file_path):
    """修复Python文件中的语法错误"""
    content = Path(file_path).read_text(encoding='utf-8')
    
    # 修复中文字符和特殊符号导致的语法问题
    # 替换全角标点为半角
    replacements = {
        '：': ':',
        '；': ';',
        '，': ',',
        '（': '(',
        '）': ')',
        '【': '[',
        '】': ']',
        '。': '.',
        '！': '!',
        '？': '?',
        '"': '"',
        '"': '"',
        ''': "'",
        ''': "'",
        '…': '...',
        '→': '->',
        '—': '--',
        '、': ',',
        '←': '<-',
        '⇒': '=>',
        '⇐': '<=',
        '±': '+/-',
        '≠': '!=',
        '≤': '<=',
        '≥': '>=',
        '×': '*',
        '÷': '/',
        '∑': 'sum',
        '∏': 'product',
        '∈': 'in',
        '∉': 'not in',
        '∀': 'for all',
        '∃': 'there exists',
        '∅': 'empty',
        '∩': 'intersect',
        '∪': 'union',
        '⊂': 'subset',
        '⊃': 'superset',
        '⊆': 'subset or equal',
        '⊇': 'superset or equal',
        '⊕': 'xor',
        '⊗': 'tensor',
        '⊙': 'hadamard',
        '¬': 'not',
        '∧': 'and',
        '∨': 'or',
        '⊻': 'xor',
        '∴': 'therefore',
        '∵': 'because',
        '∶': ':',
        '∷': '::',
        '∸': 'minus',
        '∹': 'swap',
        '∺': 'plus minus',
        '∻': 'rev eq',
        '∼': '~',
        '∽': 'tilde',
        '∾': 'tilde minus',
        '∿': 'sine wave',
        '≀': 'wreath',
        '≁': 'not tilde',
        '≂': 'minus tilde',
        '≃': 'tilde tilde',
        '≄': 'not tilde tilde',
        '≅': 'cong',
        '≆': 'smile',
        '≇': 'not congruent',
        '≈': 'approx',
        '≉': 'not approx',
        '≊': 'asymp',
        '≋': 'asymp equal',
        '≌': 'iso',
        '≍': 'image',
        '≎': 'multimap',
        '≏': 'multimap reverse',
        '≐': 'dot equal',
        '≑': 'equal time',
        '≒': 'approx equal',
        '≓': 'eq def',
        '≔': 'coloneq',
        '≕': 'eq coloneq',
        '≖': 'bump equal',
        '≗': 'doteq',
        '≘': 'eq triangle',
        '≙': 'eq diamond',
        '≚': 'eq circle',
        '≛': 'eq parallel',
        '≜': 'eq sim',
        '≝': 'eq eq',
        '≞': 'eq cong',
        '≟': 'quest eq',
    }
    
    for old, new in replacements.items():
        content = content.replace(old, new)
    
    # 修复其他常见的语法问题
    content = re.sub(r'\)\s*\)\s*->', ') ->', content)  # 修复多余的右括号
    content = re.sub(r'df\s*:.*?->', 'df ->', content)   # 修复变量类型注解问题
    
    # 写回文件
    Path(file_path).write_text(content, encoding='utf-8')
    print(f"已修复: {file_path}")

if __name__ == "__main__":
    fix_syntax_errors("src/excel_mcp_server_fastmcp/api/advanced_sql_query.py")
