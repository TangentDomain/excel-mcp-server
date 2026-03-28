#!/usr/bin/env python3
"""
GitHub Star 统计和感谢机制
自动获取GitHub stars信息并生成感谢内容
"""

import requests
import json
from datetime import datetime
import os

def get_github_stars():
    """获取GitHub仓库的stars数量和最新stargazers"""
    repo = "TangentDomain/excel-mcp-server"
    
    try:
        # 获取stars数量
        stars_url = f"https://api.github.com/repos/{repo}"
        response = requests.get(stars_url)
        response.raise_for_status()
        repo_data = response.json()
        stars_count = repo_data.get("stargazers_count", 0)
        
        # 获取最新stargazers（最近10个）
        stargazers_url = f"https://api.github.com/repos/{repo}/stargazers?per_page=10"
        response = requests.get(stargazers_url)
        response.raise_for_status()
        stargazers = response.json()
        
        return {
            "stars": stars_count,
            "recent_stargazers": stargazers[:10],  # 最近10个star者
            "updated_at": datetime.now().isoformat()
        }
        
    except Exception as e:
        print(f"获取GitHub信息失败: {e}")
        return {
            "stars": 0,
            "recent_stargazers": [],
            "updated_at": datetime.now().isoformat(),
            "error": str(e)
        }

def generate_thanks_message(stars_data):
    """生成感谢信息"""
    stars = stars_data["stars"]
    recent_stargazers = stars_data["recent_stargazers"]
    
    if "error" in stars_data:
        return f"⚠️ GitHub API暂时不可用 (错误: {stars_data['error']})"
    
    # 生成感谢内容
    thanks_lines = []
    
    if stars >= 100:
        thanks_lines.append(f"🎉 **突破100 stars！** 感谢社区的支持！")
    elif stars >= 50:
        thanks_lines.append(f"🌟 **当前 {stars} stars！** 目标100，继续加油！")
    elif stars >= 10:
        thanks_lines.append(f"⭐ **当前 {stars} stars！** 感谢每一位支持者！")
    else:
        thanks_lines.append(f"🌱 **当前 {stars} stars！** 感谢你的支持！")
    
    if recent_stargazers:
        thanks_lines.append("\n🙏 **感谢最新Star者：**")
        for user in recent_stargazers:
            username = user.get("login", "unknown")
            avatar = user.get("avatar_url", "")
            thanks_lines.append(f"  - @{username}")
    
    return "\n".join(thanks_lines)

def update_readme(stars_data):
    """更新README中的感谢信息"""
    thanks_message = generate_thanks_message(stars_data)
    
    # 在README.md末尾添加感谢信息
    readme_path = "README.md"
    
    try:
        with open(readme_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 查找最后一个标题的位置
        insert_pos = content.rfind("## 🎉 致谢")
        
        if insert_pos != -1:
            # 找到致谢部分的末尾
            end_pos = content.find("\n---", insert_pos)
            if end_pos == -1:
                end_pos = len(content)
            
            # 插入感谢内容
            new_content = content[:insert_pos] + f"""## 🎉 致谢

{thanks_message}

感谢所有贡献者和用户的支持！特别感谢游戏开发社区的反馈和建议。

让AI为游戏开发赋能！ 🚀

---""" + content[end_pos:]
            
            with open(readme_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            
            print("README.md 感谢信息已更新")
        else:
            print("未找到致谢部分，跳过更新")
            
    except Exception as e:
        print(f"更新README失败: {e}")

def main():
    """主函数"""
    print("🌟 GitHub Star 统计和感谢机制")
    print("=" * 40)
    
    # 获取GitHub数据
    stars_data = get_github_stars()
    
    # 显示当前状态
    stars = stars_data["stars"]
    print(f"当前Stars数量: {stars}")
    
    # 生成和显示感谢信息
    thanks_message = generate_thanks_message(stars_data)
    print("\n" + thanks_message)
    
    # 更新README
    update_readme(stars_data)
    
    print(f"\n📊 统计完成于: {stars_data['updated_at']}")

if __name__ == "__main__":
    main()