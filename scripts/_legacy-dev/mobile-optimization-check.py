#!/usr/bin/env python3
"""
用户体验优化检查与实施脚本
识别并改善文档的移动端友好性、导航结构和交互体验
"""

import os
import re
import json
from pathlib import Path

def analyze_readme_mobile_friendly():
    """分析README.md的移动端友好性"""
    readme_path = Path("README.md")
    if not readme_path.exists():
        return {"error": "README.md not found"}
    
    with open(readme_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    analysis = {
        "mobile_friendly_issues": [],
        "navigation_issues": [],
        "interaction_issues": [],
        "content_structure": {}
    }
    
    # 检查表格移动端友好性
    table_pattern = r'\|[^|]+\|[^|]+\|[^|]+\|'
    tables = re.findall(table_pattern, content)
    
    if len(tables) > 2:
        analysis["mobile_friendly_issues"].append(
            "发现复杂表格，在移动端可能显示不全"
        )
    
    # 检查代码块长度
    code_blocks = re.findall(r'```.*?```', content, re.DOTALL)
    for i, block in enumerate(code_blocks):
        if len(block) > 800:  # 代码块过长
            analysis["mobile_friendly_issues"].append(
                f"代码块{i+1}过长({len(block)}字符)，移动端需要滚动"
            )
    
    # 检查导航结构
    headers = re.findall(r'^#+\s+(.+)$', content, re.MULTILINE)
    analysis["content_structure"]["headers"] = headers
    analysis["content_structure"]["header_count"] = len(headers)
    
    # 检查内部链接
    internal_links = re.findall(r'\[([^\]]+)\]\(([^)]+)\)', content)
    analysis["content_structure"]["internal_links"] = internal_links
    
    return analysis

def generate_mobile_optimizations():
    """生成移动端优化建议"""
    optimizations = {
        "table_optimization": {
            "title": "表格移动端优化",
            "changes": [
                "将复杂表格拆分为多个简单表格",
                "添加表格说明文字",
                "为长表格添加'点击查看详情'折叠功能"
            ]
        },
        "code_block_optimization": {
            "title": "代码块优化",
            "changes": [
                "长代码块添加滚动条",
                "代码块添加语言标识",
                "重要代码块提供复制按钮"
            ]
        },
        "navigation_enhancement": {
            "title": "导航结构优化",
            "changes": [
                "添加目录导航",
                "关键章节添加返回顶部链接",
                "改善章节间的逻辑连接"
            ]
        },
        "interactive_elements": {
            "title": "交互元素增强",
            "changes": [
                "添加'一键复制'按钮",
                "提供'展开/收起'功能",
                "添加搜索关键词高亮"
            ]
        }
    }
    return optimizations

def create_mobile_friendly_css():
    """创建移动端友好CSS样式"""
    css_content = """
/* 移动端友好样式 */
@media (max-width: 768px) {
    /* 表格响应式处理 */
    .markdown-table {
        display: block;
        overflow-x: auto;
        white-space: nowrap;
        -webkit-overflow-scrolling: touch;
    }
    
    /* 代码块滚动 */
    .code-block {
        max-height: 300px;
        overflow-y: auto;
        border-radius: 8px;
    }
    
    /* 导航优化 */
    .nav-menu {
        position: sticky;
        top: 10px;
        background: rgba(255, 255, 255, 0.9);
        padding: 10px;
        border-radius: 8px;
    }
    
    /* 字体适配 */
    body {
        font-size: 16px;
        line-height: 1.6;
    }
    
    /* 按钮优化 */
    .mobile-button {
        min-width: 44px;
        min-height: 44px;
        padding: 12px 20px;
        font-size: 16px;
    }
}

/* 通用优化 */
.readable-content {
    max-width: 1200px;
    margin: 0 auto;
    padding: 20px;
}

.highlight-box {
    background: #f0f4ff;
    border-left: 4px solid #4f46e5;
    padding: 16px;
    margin: 16px 0;
    border-radius: 0 8px 8px 0;
}

.interactive-tip {
    background: #fef3c7;
    border: 1px solid #f59e0b;
    border-radius: 8px;
    padding: 12px;
    margin: 12px 0;
}
"""
    return css_content

def create_interactive_js():
    """创建交互功能JavaScript"""
    js_content = """
// 文档交互功能
document.addEventListener('DOMContentLoaded', function() {
    // 返回顶部按钮
    const backToTop = document.createElement('button');
    backToTop.textContent = '↑';
    backToTop.className = 'back-to-top';
    backToTop.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #4f46e5;
        color: white;
        border: none;
        border-radius: 50%;
        width: 50px;
        height: 50px;
        font-size: 20px;
        cursor: pointer;
        display: none;
        z-index: 1000;
    `;
    document.body.appendChild(backToTop);
    
    // 监听滚动显示/隐藏返回顶部按钮
    window.addEventListener('scroll', function() {
        if (window.pageYOffset > 300) {
            backToTop.style.display = 'block';
        } else {
            backToTop.style.display = 'none';
        }
    });
    
    backToTop.addEventListener('click', function() {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
    
    // 代码块复制按钮
    document.querySelectorAll('pre').forEach(function(pre) {
        const copyButton = document.createElement('button');
        copyButton.textContent = '复制';
        copyButton.className = 'copy-button';
        copyButton.style.cssText = `
            position: absolute;
            top: 10px;
            right: 10px;
            background: #6b7280;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 5px 10px;
            font-size: 12px;
            cursor: pointer;
        `;
        
        pre.style.position = 'relative';
        pre.appendChild(copyButton);
        
        copyButton.addEventListener('click', function() {
            const code = pre.textContent;
            navigator.clipboard.writeText(code).then(function() {
                copyButton.textContent = '已复制';
                setTimeout(function() {
                    copyButton.textContent = '复制';
                }, 2000);
            });
        });
    });
    
    // 添加导航菜单
    const headers = document.querySelectorAll('h1, h2, h3');
    const nav = document.createElement('nav');
    nav.className = 'nav-menu';
    nav.innerHTML = '<h3>快速导航</h3><ul></ul>';
    
    const navList = nav.querySelector('ul');
    
    headers.forEach(function(header, index) {
        if (header.tagName === 'H1' || header.tagName === 'H2') {
            const li = document.createElement('li');
            li.innerHTML = `<a href="#${header.id}">${header.textContent}</a>`;
            navList.appendChild(li);
        }
    });
    
    // 找到第一个H2后插入导航
    const firstH2 = document.querySelector('h2');
    if (firstH2) {
        firstH2.parentNode.insertBefore(nav, firstH2);
    }
});

// 搜索高亮功能
function highlightSearchTerm(term) {
    const content = document.querySelector('.readable-content');
    const text = content.innerHTML;
    const regex = new RegExp(`(${term})`, 'gi');
    content.innerHTML = text.replace(regex, '<mark>$1</mark>');
}
"""
    return js_content

def main():
    """主执行函数"""
    print("开始用户体验优化检查...")
    
    # 分析当前状态
    analysis = analyze_readme_mobile_friendly()
    print("当前状态分析:")
    print(json.dumps(analysis, indent=2, ensure_ascii=False))
    
    # 生成优化建议
    optimizations = generate_mobile_optimizations()
    print("\n优化建议:")
    for category, details in optimizations.items():
        print(f"\n{details['title']}:")
        for change in details['changes']:
            print(f"  - {change}")
    
    # 创建优化文件
    print("\n创建优化文件...")
    
    # 创建移动端CSS
    css_content = create_mobile_friendly_css()
    with open('docs/mobile-friendly.css', 'w', encoding='utf-8') as f:
        f.write(css_content)
    print("✅ 创建移动端样式文件: docs/mobile-friendly.css")
    
    # 创建交互JavaScript
    js_content = create_interactive_js()
    with open('docs/mobile-friendly.js', 'w', encoding='utf-8') as f:
        f.write(js_content)
    print("✅ 创建交互功能文件: docs/mobile-friendly.js")
    
    # 创建移动端优化说明文档
    optimization_doc = f"""# 移动端优化实施指南

## 🎯 优化目标
提升文档在移动设备上的可读性和交互体验

## 📱 已实施优化

### 1. 响应式样式 (docs/mobile-friendly.css)
- 表格响应式滚动
- 代码块高度限制和滚动
- 移动端字体优化
- 按钮尺寸适配

### 2. 交互功能 (docs/mobile-friendly.js)
- 返回顶部按钮
- 代码块复制功能
- 自动导航菜单
- 搜索关键词高亮

## 🔧 使用方法

### 在HTML文档中引入
```html
<link rel="stylesheet" href="docs/mobile-friendly.css">
<script src="docs/mobile-friendly.js"></script>
```

### 在Markdown中应用
在转换为HTML的文档中，添加以下CSS类：
- `readable-content` - 内容容器
- `highlight-box` - 重要提示框
- `interactive-tip` - 交互提示

## 📊 优化效果

### 移动端体验提升
- 📱 表格显示：支持横向滚动，避免内容截断
- 📝 代码阅读：高度限制，便于上下浏览
- 🔗 快速导航：固定导航菜单，一键跳转
- 📋 操作便捷：大按钮设计，易于点击

### 性能优化
- 🚀 加载速度：CSS/JS文件轻量，最小化影响
- 💾 缓存友好：静态文件可被浏览器缓存
- ⚡ 延迟加载：JavaScript只在DOM加载后执行

## 🔄 后续优化计划

1. **移动端专门版本文档**
   - 创建适合手机阅读的简明版本
   - 优化长文档的分页显示

2. **离线支持**
   - 添加PWA功能
   - 支持离线阅读

3. **触摸优化**
   - 添加手势支持
   - 优化触摸目标大小

---

*优化版本：v1.6.50*  
*实施时间：2026-03-29*  
*更新内容：移动端友好性优化*
"""

    with open('docs/mobile-friendly-optimization.md', 'w', encoding='utf-8') as f:
        f.write(optimization_doc)
    print("✅ 创建优化说明文档: docs/mobile-friendly-optimization.md")
    
    print("\n🎉 移动端用户体验优化完成!")
    print("主要改进：")
    print("  - 响应式表格和代码块")
    print("  - 导航菜单和返回顶部")
    print("  - 一键复制功能")
    print("  - 搜索高亮和交互提示")

if __name__ == "__main__":
    main()