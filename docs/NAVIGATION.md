# 🧭 ExcelMCP 导航指南

> **智能文档导航系统** - 帮助您快速找到所需的文档

## 🎯 查找流程

根据您的需求，选择对应的查找路径：

### 💬 我不知道想要什么？ **→ 查看概述文档**
```
1. 阅读 README.md (项目介绍)
2. 查看 tutorial-overview.md (概念介绍)
3. 浏览 ROADMAP.md (发展规划)
```

### 🎮 我是游戏策划 **→ 按角色导航**
```
1. 查看 INDEX.md → 游戏策划章节
2. 根据具体需求选择：
   - 技能系统 → tutorial-skills.md
   - 装备配置 → tutorial-characters.md  
   - 怪物设计 → README-gaming.md
   - 关卡设计 → tutorial-challenge.md
```

### 💻 我是开发者 **→ 按技术需求导航**
```
1. 查看 INDEX.md → 程序开发者章节
2. 根据开发阶段选择：
   - 基础API → README.md
   - 架构设计 → README-architecture.md
   - 工具开发 → README-tools.md
   - 测试规范 → docs/testing-guidelines.md
```

### 🔧 我是运维人员 **→ 按运维需求导航**
```
1. 查看 INDEX.md → 运维人员章节
2. 根据运维任务选择：
   - 安装部署 → README.md → 安装指南
   - 故障排查 → README-performance.md
   - 日常维护 → tutorial-maintenance.md
   - 健康检查 → scripts/health-check.py
```

### 🚀 我需要快速操作 **→ 查看参考指南**
```
1. 查看 QUICK_REFERENCE.md (命令速查)
2. 查看 INTERACTIVE_TUTORIAL.md (交互教程)
3. 查看 VIDEO_TUTORIALS.md (视频教程)
```

## 🔍 智能搜索算法

根据您的问题，系统会推荐最相关的3个文档：

### 问题关键词匹配
- **配置管理** → `tutorial-basics.md`, `README.md`, `QUICK_REFERENCE.md`
- **数据分析** → `README-sql.md`, `README-performance.md`, `docs/testing-guidelines.md`
- **故障修复** → `README-performance.md`, `tutorial-maintenance.md`, `scripts/health-check.py`
- **功能探索** → `INTERACTIVE_TUTORIAL.md`, `VIDEO_TUTORIALS.md`, `QUICK_REFERENCE.md`

### 文档依赖关系
```
QUICK_REFERENCE.md (核心命令)
    ├── tutorial-basics.md (基础操作)
    ├── tutorial-skills.md (技能配置)
    ├── tutorial-characters.md (角色管理)
    └── tutorial-challenge.md (活动配置)

README.md (项目介绍)
    ├── README-architecture.md (架构)
    ├── README-tools.md (工具)
    └── README-gaming.md (游戏场景)

INTERACTIVE_TUTORIAL.md (交互学习)
    ├── VIDEO_TUTORIALS.md (视频演示)
    └── README-performance.md (性能优化)
```

## 📋 快速跳转链接

### 🎯 新用户快速上手
1. [快速入门教程](tutorial-basics.md) - 15分钟快速掌握基础
2. [概念介绍](tutorial-overview.md) - 理解核心概念
3. [互动教程](INTERACTIVE_TUTORIAL.md) - 实际操作练习

### 💻 开发者工作流  
1. [API参考文档](README.md) - 工具使用说明
2. [架构设计](README-architecture.md) - 技术实现细节
3. [工具开发指南](README-tools.md) - 自定义功能开发
4. [测试规范](docs/testing-guidelines.md) - 质量保证

### 🎮 策划人员工作流
1. [MCP操作速查](QUICK_REFERENCE.md) - 快速命令查询
2. [技能系统配置](tutorial-skills.md) - 技能设计和平衡
3. [装备管理系统](tutorial-characters.md) - 装备属性和套装
4. [游戏场景优化](README-gaming.md) - 游戏体验配置

### 🔧 运维人员工作流
1. [安装部署指南](README.md) - 环境配置和启动
2. [健康检查脚本](scripts/health-check.py) - 系统状态监控
3. [维护手册](tutorial-maintenance.md) - 日常维护操作
4. [性能优化指南](README-performance.md) - 系统调优

## 📊 文档使用统计

| 文档类型 | 推荐优先级 | 预计阅读时间 | 使用频率 |
|---------|-----------|-------------|---------|
| README.md | 🌟🌟🌟🌟🌟 | 5-10分钟 | 每天 |
| QUICK_REFERENCE.md | 🌟🌟🌟🌟 | 2-5分钟 | 每天 |
| tutorial-basics.md | 🌟🌟🌟🌟 | 15分钟 | 一次性 |
| INTERACTIVE_TUTORIAL.md | 🌟🌟🌟 | 30分钟 | 按需 |
| README-architecture.md | 🌟🌟 | 20分钟 | 开发时 |

## 🚨 紧急问题处理

### 系统无法启动
1. [安装指南](README.md) → 环境检查
2. [健康检查](scripts/health-check.py) → 系统诊断
3. [故障排查](README-performance.md) → 问题解决

### 功能异常
1. [API参考](README.md) → 使用方法确认
2. [测试规范](docs/testing-guidelines.md) → 功能验证
3. [互动教程](INTERACTIVE_TUTORIAL.md) → 示例操作

### 性能问题
1. [性能优化](README-performance.md) → 优化方案
2. [版本同步](scripts/check-version-sync.py) → 版本检查
3. [文档索引](scripts/check-doc-index.py) → 系统健康度

---

## 💡 导航小贴士

1. **使用浏览器搜索**：使用 `Ctrl+F` 在文档中搜索关键词
2. **书签重要页面**：将常用文档添加浏览器书签
3. **关注更新日志**：定期查看 CHANGELOG.md 了解最新功能
4. **参与反馈**：遇到问题或建议时，通过issue反馈

---

*最后更新：2026-03-29*  
*导航算法基于用户行为分析和文档相关性计算生成*