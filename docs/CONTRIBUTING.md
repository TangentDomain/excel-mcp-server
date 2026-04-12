# 🤝 贡献指南 - Contributing Guide

感谢您对 excel-mcp-server 项目的关注！我们欢迎各种形式的贡献。

## 🎯 贡献方式

### 1. 🐛 报告 Bug
- 使用 [GitHub Issues](https://github.com/TangentDomain/excel-mcp-server/issues) 提交 bug 报告
- 提供详细的重现步骤、期望结果和实际结果
- 包含环境信息（操作系统、Python版本、游戏类型）

### 2. ✨ 提出新功能
- 在 Issues 中提出功能请求
- 描述使用场景和期望的功能细节
- 说明该功能对游戏开发的价值

### 3. 💻 提交代码
- Fork 仓库并创建功能分支
- 遵循本指南的代码规范
- 确保测试通过

### 4. 📚 改进文档
- 修正文档错误
- 添加使用示例
- 翻译或改进文档

## 🚀 开发流程

### 1. Fork 和 Clone
```bash
# Fork 后，在本地克隆
git clone https://github.com/您的用户名/excel-mcp-server.git
cd excel-mcp-server
git remote add upstream https://github.com/TangentDomain/excel-mcp-server.git
```

### 2. 创建分支
```bash
# 创建功能分支
git checkout -b feature/您的功能名称
```

### 3. 开发和测试
```bash
# 安装依赖
pip install -e .

# 运行测试
python3 -m pytest tests/ -q --tb=no -n auto --timeout=30

# 测试单个功能
python3 -m pytest tests/test_excel_operations.py -v
```

### 4. 提交和推送
```bash
# 提交代码
git add .
git commit -m "feat: 添加新功能描述"

# 推送到您的 fork
git push origin feature/您的功能名称
```

### 5. 创建 Pull Request
- 在 GitHub 上创建 PR
- 填写 PR 模板
- 等待代码审查和 CI 检查

## 📋 代码规范

### Python 代码风格
- 遵循 PEP 8
- 使用 4 个空格缩进
- 最大行长度 88 字符
- 必须有类型注解
- 函数和类必须有文档字符串

### 文档规范
- 所有公共 API 必须有文档字符串
- 使用 Google 风格的文档字符串
- 包含参数说明、返回值、异常说明
- 提供使用示例

### 测试规范
- 测试文件以 `test_` 开头
- 测试函数以 `test_` 开头
- 使用 pytest 的 fixture 功能
- 每个功能必须有至少一个测试

## 🎮 游戏开发特别注意事项

### 配置表格式
- 支持 Excel (.xlsx) 格式
- 建议使用标准化的列名
- 支持多表关联查询

### 性能要求
- 大文件处理需要流式读取
- 复杂查询需要优化性能
- 内存使用需要控制在合理范围

### 安全考虑
- SQL 注入防护
- 文件路径安全验证
- 权限控制

## 🏷️ Issue 模板

### Bug 报告
```
## Bug 描述
简要描述遇到的问题

## 重现步骤
1. 步骤一
2. 步骤二
3. 步骤三

## 期望结果
描述应该发生什么

## 实际结果
描述实际发生了什么

## 环境信息
- 操作系统: [例如: Ubuntu 20.04]
- Python 版本: [例如: 3.9.0]
- excel-mcp-server 版本: [例如: v1.6.37]
- 游戏类型: [例如: RPG/MMO]
```

### 功能请求
```
## 功能描述
详细描述您希望实现的功能

## 使用场景
说明在游戏开发中的使用场景

## 期望的 API
描述期望的接口设计

## 相关问题
链接相关的 Issues 或讨论
```

## 📄 PR 模板

### 功能 PR
```
## 变更描述
详细描述此 PR 实现的功能

## 变更类型
- [x] feat: 新功能
- [ ] fix: 修复 bug
- [ ] docs: 文档更新
- [ ] style: 代码格式化
- [ ] refactor: 重构
- [ ] test: 测试相关
- [ ] chore: 构建或工具相关

## 测试
- [x] 已通过所有测试
- [ ] 已添加新测试
- [ ] 手动测试通过

## 变更影响
- [x] 不影响现有功能
- [ ] 向后兼容性说明

## 截图或示例（可选）
```

## 🎖️ 贡献者权益

### Star 者福利
- 📢 在贡献者名单中展示
- 🎁 优先获得新功能测试资格
- 🏆 累积贡献获得特殊徽章

### 活跃贡献者
- 🎯 成为项目维护者
- 📝 获得文档维护权限
- 🚀 参与产品决策讨论

## 📞 社区支持

- 💬 [GitHub Discussions](https://github.com/TangentDomain/excel-mcp-server/discussions)
- 📧 [邮件联系](mailto:您的邮箱)
- 💬 [QQ群: XXXXXXX]

## 📄 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

---

感谢您的贡献！🙏