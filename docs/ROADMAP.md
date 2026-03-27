# ROADMAP — 路线图

> CEO维护，子代理参考。每个Phase有完成标志，达标即进入下一阶段。

## Phase 1：核心能力 ✅ 达标
- SQL引擎（45项功能）、安全加固、calamine性能、跨文件JOIN
- **完成标志**：100/100连续10轮 → ✅ 连续18轮

## Phase 2：体验打磨（当前）
- 返回值统一、错误结构化、合并重复工具
- 文档门面（README优化、竞品对比、使用示例）
- FROM子查询、写入性能优化
- **完成标志**：AI选工具准确率>95% + README有30秒上手教程 + 文档门面完善

## Phase 3：生态扩展（规划）
- 配置模板系统、VSCode插件、社区建设
- **完成标志**：GitHub star > 100 + 至少1个社区贡献

## 重大决策
> 详细记录见 [DECISIONS.md](DECISIONS.md)
- 不做配置导出引擎（构建工具，不是MCP查询工具）
- 不做View/写入校验/Auto Increment（SQL引擎已覆盖）
- calamine替换openpyxl读取（2300x性能提升）
