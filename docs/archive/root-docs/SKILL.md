---
name: excel-iteration
description: Excel MCP Server迭代。持续改进Excel MCP工具功能、性能和文档。触发词：Excel迭代、excel-mcp、Excel工具改进。
---

# Excel MCP 迭代 Skill

## 触发条件
- cron定时触发（每小时）
- 用户说"Excel迭代"、"运行Excel迭代"

## 项目信息
- **路径**：`/root/.openclaw/workspace/excel-mcp-server`

## 每轮流程（K0-K4）

### K0：环境准备与环境检查
- 读 `/root/.openclaw/workspace/skills/excel-iteration/SKILL.md` 了解任务
- 检查excel-mcp-server目录是否存在：`ls /root/.openclaw/workspace/excel-mcp-server/`
- 检查REQUIREMENTS.md是否存在，记录OPEN需求数量
- 检查cron状态，确认任务能正常运行
- 检查所有Python文件语法：`python -m py_compile src/*.py`
- 创建 `.step-K0.done` 标记完成

### K1：需求管理
- 读 `/root/.openclaw/workspace/excel-mcp-server/REQUIREMENTS.md`（如果存在）
- 查找第一个 `状态：OPEN` 的需求
- 按优先级执行需求（P0 > P1 > P2）
- 更新REQUIREMENTS.md状态
- 创建 `.step-K1.done` 标记完成

### K2：编码/执行
- 如果无OPEN需求，则执行Excel MCP核心任务：
  1. 读 `.cron-prompt.md` 了解上轮状态
  2. 读 `docs/NOW.md` 了解当前进展
  3. 运行测试检查（`pytest tests/ -v`）
  4. 检查API工具完整性
  5. 运行性能基准测试
  6. 生成改进建议
- 如果有编码任务，确保通过 `python -m py_compile` 检查语法
- 创建 `.step-K2.done` 标记完成

### K3：推送报告
- 生成Excel MCP迭代报告，包含生态健康总览
- 双渠道推送：
  - `message(action=send, channel=feishu, target=ou_f87e1db600e058a4ccc89ce3053ec9d5, message=报告)`
  - `message(action=send, channel=qqbot, target=qqbot:c2c:AAE80456F09EFD98D108CE91DC2B5A73, message=报告)`
- 推送后回复 NO_REPLY
- 创建 `.step-K3.done` 标记完成

### K4：链式执行
- 有OPEN需求 且 已连续执行<5轮 → 回到K0继续
- 无OPEN需求 或 满5轮 → 结束

## 链式执行逻辑
- 记录轮次号到 `.round-number`
- 每轮递增，最多5轮
- 满5轮后自动结束，等下一轮cron触发

## 核心任务
- API工具完善
- 性能优化
- 文档同步
- 测试覆盖率提升
- 错误处理改进

## 红线
1. 代码必须通过语法检查
2. 推送报告不可跳过
3. 所有时间用北京时间
4. 必须有K0-K3步骤和step标记
5. 无论哪步失败都必须推送报告
6. 保持与现有功能兼容性