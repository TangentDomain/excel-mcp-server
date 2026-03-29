#!/usr/bin/env python3
"""
文档索引完整性检查脚本
- 检查INDEX.md是否存在且完整
- 检查NAVIGATION.md是否存在且完整  
- 检查大文档是否需要拆分
- 发现问题记录到DECISIONS.md
"""

import os
import re
from pathlib import Path
from datetime import datetime

def check_doc_index():
    """检查文档索引系统完整性"""
    docs_dir = Path("docs")
    issues = []
    
    # 检查INDEX.md
    index_file = docs_dir / "INDEX.md"
    if not index_file.exists():
        issues.append("INDEX.md 不存在")
    else:
        with open(index_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # 检查是否包含用户角色分类
            if "游戏策划" not in content and "程序开发者" not in content and "运维人员" not in content:
                issues.append("INDEX.md 缺少用户角色分类")
    
    # 检查NAVIGATION.md
    nav_file = docs_dir / "NAVIGATION.md"
    if not nav_file.exists():
        issues.append("NAVIGATION.md 不存在")
    else:
        with open(nav_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # 检查是否包含导航要素
            if "文档依赖关系" not in content and "查找流程" not in content:
                issues.append("NAVIGATION.md 缺少导航要素")
    
    # 检查大文档
    for md_file in docs_dir.glob("*.md"):
        size = md_file.stat().st_size
        if size > 10 * 1024:  # 10KB
            if md_file.name == "testing-guidelines.md" and size > 50 * 1024:  # 50KB
                issues.append(f"{md_file.name} 文件过大({size/1024:.0f}KB)，建议拆分")
    
    return issues

def main():
    print(f"🔍 开始文档索引检查... {datetime.utcnow().isoformat()}")
    
    issues = check_doc_index()
    
    if not issues:
        print("✅ 文档索引系统完整，无需修复")
        return True
    else:
        print("❌ 发现以下问题:")
        for i, issue in enumerate(issues, 1):
            print(f"  {i}. {issue}")
        
        # 记录到DECISIONS.md
        decision_content = f"""
### [文档索引系统检查] 第{len(issues)}次发现问题
- **时间**: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}
- **问题**: {', '.join(issues)}
- **建议**: 立即修复文档索引系统缺失项
- **依据**: RULES.md智能文档索引系统规则
"""
        
        with open("docs/DECISIONS.md", 'a', encoding='utf-8') as f:
            f.write(decision_content)
        
        print(f"📝 已记录到docs/DECISIONS.md")
        return False

if __name__ == "__main__":
    main()