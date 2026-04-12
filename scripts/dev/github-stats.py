#!/usr/bin/env python3
"""
GitHub 统计工具 - 显示仓库统计信息
"""

import json
import requests
from datetime import datetime

def get_github_stats(repo="TangentDomain/excel-mcp-server"):
    """获取 GitHub 仓库统计信息"""
    try:
        # 基本信息
        info_url = f"https://api.github.com/repos/{repo}"
        info_response = requests.get(info_url, timeout=10)
        info_response.raise_for_status()
        info_data = info_response.json()
        
        # 获取贡献者信息
        contributors_url = f"https://api.github.com/repos/{repo}/contributors"
        contributors_response = requests.get(contributors_url, timeout=10)
        contributors_response.raise_for_status()
        contributors_data = contributors_response.json()
        
        # 获取最近的提交
        commits_url = f"https://api.github.com/repos/{repo}/commits?per_page=10"
        commits_response = requests.get(commits_url, timeout=10)
        commits_response.raise_for_status()
        commits_data = commits_response.json()
        
        stats = {
            "repo": repo,
            "stars": info_data.get("stargazers_count", 0),
            "forks": info_data.get("forks_count", 0),
            "watchers": info_data.get("subscribers_count", 0),
            "issues": info_data.get("open_issues_count", 0),
            "language": info_data.get("language", "Unknown"),
            "created_at": info_data.get("created_at", ""),
            "updated_at": info_data.get("updated_at", ""),
            "contributors_count": len(contributors_data),
            "recent_commits": len(commits_data),
            "total_commits": info_data.get("size", 0),  # 这是一个近似值
        }
        
        # 计算里程碑进度
        stats["milestones"] = {
            "50_stars": {"current": stats["stars"], "target": 50, "achieved": stats["stars"] >= 50},
            "100_stars": {"current": stats["stars"], "target": 100, "achieved": stats["stars"] >= 100},
            "200_stars": {"current": stats["stars"], "target": 200, "achieved": stats["stars"] >= 200},
            "500_stars": {"current": stats["stars"], "target": 500, "achieved": stats["stars"] >= 500},
        }
        
        return stats
        
    except Exception as e:
        print(f"获取 GitHub 统计信息失败: {e}")
        return None

def generate_github_readme_section(stats):
    """生成 GitHub README 相关部分"""
    if not stats:
        return ""
    
    # 生成统计徽章
    badges = f"""
[![GitHub stars](https://img.shields.io/github/stars/{stats['repo']}?style=social&label=Star&color=gold)](https://github.com/{stats['repo']}/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/{stats['repo']}?style=social)](https://github.com/{stats['repo']}/network)
[![GitHub issues](https://img.shields.io/github/issues/{stats['repo']}?style=social)](https://github.com/{stats['repo']}/issues)
[![GitHub language](https://img.shields.io/github/languages/top/{stats['repo']}?style=social)](https://github.com/{stats['repo']})
[![GitHub last commit](https://img.shields.io/github/last-commit/{stats['repo']}?style=social)](https://github.com/{stats['repo']}/commits/main)

"""
    
    # 生成统计信息
    section = f"""
## 📊 GitHub 统计

{badges}

| 指标 | 数值 | 状态 |
|------|------|------|
| ⭐ Stars | {stats['stars']} | 🎯 **目标: 100** |
| 🍴 Forks | {stats['forks']} | 📈 活跃度 |
| 👀 Watchers | {stats['watchers']} | 🔔 关注度 |
| 🐛 Issues | {stats['issues']} | 📝 待处理 |
| 💻 语言 | {stats['language']} | 🛠️ 技术 |
| 👥 贡献者 | {stats['contributors_count']} | 🤝 社区 |
| 📝 最近提交 | {stats['recent_commits']} | 🚀 活跃度 |

## 🎯 里程碑进度

"""
    
    # 添加里程碑进度
    for milestone_key, milestone_info in stats['milestones'].items():
        achieved = "✅" if milestone_info['achieved'] else "⏳"
        target_name = milestone_key.replace('_', ' ').title()
        section += f"- {achieved} **{target_name}**: {milestone_info['current']} / {milestone_info['target']}\n"
    
    section += f"""

## 📈 项目状态
- **创建时间**: {stats['created_at'][:10] if stats['created_at'] else 'Unknown'}
- **最后更新**: {stats['updated_at'][:10] if stats['updated_at'] else 'Unknown'}
- **社区活跃度**: {"🔥 高度活跃" if stats['recent_commits'] > 5 else "📊 活跃"}
- **发展潜力**: {"🚀 良好" if stats['stars'] > 10 else "🌱 成长中"}

## 🤝 参与方式

感谢关注！您可以通过以下方式参与项目：

1. 🌟 **Star**: 如果项目对您有帮助，请给我们一个 Star
2. 🐛 **Issue**: 报告 Bug 或提出功能建议
3. 💻 **Code**: 提交代码改进和修复
4. 📚 **Docs**: 改进文档和使用示例
5. 📢 **Share**: 分享项目给更多开发者

"""
    
    return section

def update_readme_with_github_stats():
    """更新 README.md 添加 GitHub 统计"""
    stats = get_github_stats()
    if not stats:
        print("无法获取 GitHub 统计信息")
        return False
    
    # 读取 README.md
    readme_path = "README.md"
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 生成新的 GitHub 统计内容
        github_section = generate_github_readme_section(stats)
        
        # 查找现有的 GitHub 统计部分并替换
        import re
        
        # 检查是否已有 GitHub 统计
        github_stats_pattern = r'## 📊 GitHub 统计.*?(?=##|\Z)'
        if re.search(github_stats_pattern, content, re.DOTALL):
            content = re.sub(github_stats_pattern, github_section.strip(), content, flags=re.DOTALL)
        else:
            # 添加到项目介绍后面
            insertion_point = content.find("## 🚀 快速开始")
            if insertion_point == -1:
                insertion_point = content.find("## Features")
            if insertion_point == -1:
                insertion_point = len(content)
            
            content = content[:insertion_point] + "\n\n" + github_section + "\n\n" + content[insertion_point:]
        
        # 写回文件
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"README.md 已更新 GitHub 统计信息")
        return True
    else:
        print("README.md 不存在")
        return False

if __name__ == "__main__":
    import os
    
    print("开始更新 GitHub 统计信息...")
    
    # 更新 README
    if update_readme_with_github_stats():
        print("✅ GitHub 统计信息更新成功")
    else:
        print("❌ GitHub 统计信息更新失败")
    
    print("📊 GitHub 统计信息更新完成！")