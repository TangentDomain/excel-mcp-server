# 📚 ExcelMCP 文档索引

> **按用户角色分类的快速导航指南**

## 🎯 快速开始

根据您的角色选择对应的文档：

### 🎮 游戏策划
**主要关注：** 技能系统、装备配置、怪物设计、关卡管理等游戏内容的配置管理

#### 核心文档
- **[技能系统配置](tutorial-skills.md)** - 技能属性、效果、等级设计
- **[装备管理](tutorial-characters.md)** - 装备属性、套装、稀有度配置  
- **[怪物配置](README-gaming.md)** - AI行为、属性、掉落设置
- **[关卡设计](tutorial-challenge.md)** - 进度统计、难度配置、活动管理

#### 实用工具
- **[MCP操作速查](QUICK_REFERENCE.md)** - 按场景分类的Excel操作命令
- **[互动教程](INTERACTIVE_TUTORIAL.md)** - 实际操作演示和练习
- **[故障排查](README-performance.md)** - 常见问题解决方法

---

### 💻 程序开发者  
**主要关注：** API调用、工具集成、技术实现、自定义功能

#### 核心文档
- **[API参考](README.md)** - MCP工具使用方式和接口说明
- **[架构设计](README-architecture.md)** - 系统架构和技术实现
- **[开发指南](docs/testing-guidelines.md)** - 开发规范和测试方法
- **[工具开发](README-tools.md)** - 自定义工具开发教程

#### 集成指南
- **[SQL查询](README-sql.md)** - 高级SQL查询和数据分析
- **[性能优化](README-performance.md)** - 大数据量处理优化
- **[错误处理](docs/testing-guidelines.md)** - 异常处理和调试技巧

---

### 🔧 运维人员
**主要关注：** 安装部署、故障排查、性能监控、系统维护

#### 核心文档
- **[安装指南](README.md)** - 快速安装和环境配置
- **[部署教程](tutorial-maintenance.md)** - 生产环境部署和监控
- **[故障排查](README-performance.md)** - 常见问题和解决方案
- **[维护手册](tutorial-maintenance.md)** - 日常维护和更新指南

#### 监控工具
- **[健康检查](scripts/health-check.py)** - 系统健康度自检脚本
- **[版本同步](scripts/check-version-sync.py)** - 版本一致性检查工具
- **[文档索引](check-doc-index.py)** - 文档系统完整性检查

---

## 📖 按功能分类

### 🎯 基础功能
- **[快速入门](tutorial-basics.md)** - 基础概念和操作
- **[概念介绍](tutorial-overview.md)** - 核心概念和术语
- **[维护指南](tutorial-maintenance.md)** - 日常维护操作

### 🚀 高级功能  
- **[SQL查询](README-sql.md)** - 高级数据分析
- **[技能系统](tutorial-skills.md)** - 复杂技能配置
- **[角色系统](tutorial-characters.md)** - 角色数据管理
- **[挑战系统](tutorial-challenge.md)** - 游戏活动配置

### 📹 多媒体教程
- **[视频教程](VIDEO_TUTORIALS.md)** - 视频演示课程
- **[互动教程](INTERACTIVE_TUTORIAL.md)** - 交互式学习体验

---

## 🔍 文档依赖关系

```
├── 📋 INDEX.md (本索引)
├── 🎯 REQ-026 (任务管理)
├── 📈 NOW.md (当前状态)
├── 📝 DECISIONS.md (决策记录)
├── 📊 REQUIREMENTS.md (需求列表)
├── 🗺️ ROADMAP.md (发展规划)
├── 📖 教程系列
├── 🔧 工具脚本
└── 📋 参考文档
```

---

## 💡 使用建议

1. **首次使用**：从 `tutorial-basics.md` 开始，建立基础概念
2. **日常开发**：优先查看 `QUICK_REFERENCE.md` 和对应角色的核心文档
3. **遇到问题**：查看 `README-performance.md` 的故障排查章节
4. **功能探索**：通过 `INTERACTIVE_TUTORIAL.md` 进行交互式学习

---

*最后更新：2026-03-29*  
*文档总数：29个 | 平均大小：6KB | 最大文档：testing-guidelines.md (62KB)*