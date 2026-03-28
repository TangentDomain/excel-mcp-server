#!/usr/bin/env python3
"""
GitHub Star 提升计划 - Star 统计工具
获取仓库 star 信息并生成统计数据
"""

import json
import requests
from datetime import datetime
import os

def get_github_stars(repo="TangentDomain/excel-mcp-server"):
    """获取 GitHub 仓库 star 数量"""
    try:
        url = f"https://api.github.com/repos/{repo}"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        return data.get("stargazers_count", 0)
    except Exception as e:
        print(f"获取 GitHub stars 失败: {e}")
        return 0

def generate_star_badge(stars_count, repo="TangentDomain/excel-mcp-server"):
    """生成 GitHub star badge"""
    return f"""
[![GitHub stars](https://img.shields.io/github/stars/{repo}?style=social&label=Star&color=gold)](https://github.com/{repo}/stargazers)
"""

def generate_star_section(stars_count, repo="TangentDomain/excel-mcp-server", target=100):
    """生成 GitHub star 激励段落"""
    percentage = min((stars_count / target) * 100, 100)
    
    return f"""
## 🌟 GitHub Star 激励计划

当前 Stars: **{stars_count}** | 目标: **{target}** | 进度: **{percentage:.1f}%**

### 🎯 Star 激励机制

感谢每一位 Star 者的支持！您的 Star 是我们持续改进的动力：

- ⭐ **每个 Star 都是对我们工作的认可**
- 🚀 **Star 数量达到里程碑时，我们将发布新功能** 
- 🎁 **社区贡献者优先获得新功能测试资格**
- 📈 **Star 目标: {target}+ 个，让我们一起推动游戏开发工具的发展！**

### 🎁 里程碑奖励

- 🏆 **50 Stars**: 发布游戏配置模板系统
- 🎯 **100 Stars**: 发布 VSCode 插件 + 社区文档
- ⭐ **200 Stars**: 发布高级功能教程视频
- 💎 **500 Stars**: 企业版功能预览

### 🚀 如何 Star

如果您觉得这个项目对游戏开发有帮助，请给我们一个 Star：
- 🖱️ 点击上方星星按钮
- 📤 分享给需要的朋友
- 💬 在社区中分享使用体验

### 📊 Star 统计

- **创建时间**: 2024年
- **当前版本**: v1.6.37
- **累计贡献者**: 多位开发者
- **感谢您的支持！** 🙏
"""

def update_readme_with_star_info():
    """更新 README.md 添加 star 信息"""
    stars = get_github_stars()
    
    # 读取 README.md
    readme_path = "README.md"
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 生成新的 star 内容
        star_badge = generate_star_badge(stars)
        star_section = generate_star_section(stars)
        
        # 检查是否已有 star 相关内容
        if "[GitHub stars]" in content:
            # 替换现有的 star 内容
            import re
            content = re.sub(r'\[!\[GitHub stars\][^\]]+\]', star_badge.strip(), content)
            content = re.sub(r'## 🌟 GitHub Star 激励计划.*?### 📊 Star 统计', 
                           star_section, content, flags=re.DOTALL)
        else:
            # 添加到项目介绍后面
            insertion_point = content.find("## 🚀 快速开始")
            if insertion_point == -1:
                insertion_point = content.find("## Features")
            if insertion_point == -1:
                insertion_point = len(content)
            
            content = content[:insertion_point] + "\n\n" + star_badge + "\n" + star_section + "\n\n" + content[insertion_point:]
        
        # 写回文件
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"README.md 已更新，当前 stars: {stars}")
        return True
    else:
        print("README.md 不存在")
        return False

def create_star_stats_file():
    """创建 star 统计文件"""
    stars = get_github_stars()
    
    stats_data = {
        "stars": stars,
        "target": 100,
        "percentage": min((stars / 100) * 100, 100),
        "last_updated": datetime.now().isoformat(),
        "milestones": {
            "50": {"stars": 50, "achieved": stars >= 50, "reward": "游戏配置模板系统"},
            "100": {"stars": 100, "achieved": stars >= 100, "reward": "VSCode插件 + 社区文档"},
            "200": {"stars": 200, "achieved": stars >= 200, "reward": "高级功能教程视频"},
            "500": {"stars": 500, "achieved": stars >= 500, "reward": "企业版功能预览"}
        }
    }
    
    with open("star-stats.json", 'w', encoding='utf-8') as f:
        json.dump(stats_data, f, indent=2, ensure_ascii=False)
    
    print(f"Star 统计文件已创建，当前 stars: {stars}")

if __name__ == "__main__":
    print("开始执行 GitHub Star 提升计划...")
    
    # 更新 README
    if update_readme_with_star_info():
        print("✅ README.md 更新成功")
    else:
        print("❌ README.md 更新失败")
    
    # 创建统计文件
    create_star_stats_file()
    print("✅ Star 统计文件创建成功")
    
    print("🌟 GitHub Star 提升计划执行完成！")