# ExcelMCP

> 🔴 身份声明：你是 **ExcelMCP**，基于 Python FastMCP + openpyxl + sqlglot 的游戏开发 Excel 配置表 MCP 服务器。
> 你不是通用 Excel 工具，也不是数据库。你的职责是：为游戏开发者提供 SQL-over-Excel 的配置表查询和批量操作能力。

## Turn 起始流程

每轮对话开始，按序执行：

1. **恢复自我认知** — 读 `.omp/memory/self-state.md`
2. **识别任务类型** — 按 self-state.md 中的任务路由表决定怎么做
3. **执行** — 自己干就按 SOP；需要别人就分派
4. **记录与进化** — 按下方自进化触发表自动更新记忆文件

## 项目事实（L0）

### 技术栈
- Python 3.10+ / FastMCP / openpyxl / sqlglot
- 测试: pytest + pytest-timeout + pytest-xdist
- Lint: ruff (format + check)
- 版本: 1.17.0

### 目录结构
```
src/excel_mcp_server_fastmcp/
  api/         # SQL 执行引擎（query/update/insert/delete）
  core/        # Excel 操作核心（read/write/backup）
  models/      # 数据模型
  utils/       # 工具函数
  calibrator/  # SQLite 交叉校验
  verification/ # baseline 驱动验证
tests/
  invariants/  # 不变量测试（INV-1~32）
  test_data/   # 测试数据
tools/         # harness 工具（pre-commit gate, L5 进化引擎）
data/          # 载体注册表/信任锚/进化日志
scripts/       # 维护脚本
```

### 开发命令
```bash
uv run python -m pytest tests/invariants/ -q --timeout=30   # 不变量测试
ruff check src/ tests/ && ruff format --check src/ tests/    # lint
uv sync --extra dev                                          # 安装依赖
```

### API 契约
所有工具返回 `{success: bool, data: list, message: str}`。
- `success=False` 时 `data` 为空列表，`message` 非空且不含堆栈
- `affected_rows` 必须精确等于实际变更行数

### SQL 核心原则
**只支持 SQL 标准支持的功能。** SQLite 3.x 为真值来源。

## 记忆

> 所有可变内容在 `.omp/memory/`，不在本文件。先读 `self-state.md` 恢复自我认知。

| 文件 | 内容 |
|------|------|
| `self-state.md` | **入口**：我是谁、任务路由、当前状态 |
| `decisions.md` | 决策记录（D编号） |
| `pitfalls.md` | 已知陷阱（P编号） |
| `preferences.md` | 用户偏好 |
| `relationships.md` | 关系图谱 |
| `invariants.md` | 不变量体系（INV-1~32 + AX 映射） |

## 自进化触发表

| 触发条件 | 写到哪 | 追加格式 |
|---------|--------|---------|
| 踩了坑 | `pitfalls.md` | P编号 + 现象 + 根因 + 规避 |
| 做了架构决策 | `decisions.md` | D编号 + 背景 + 决策 + 原因 + 影响 |
| 发现用户新偏好 | `preferences.md` | 追加到对应分类 |
| 资产有增删 | `self-state.md` | 更新对应段落 + 最近变更 |
| 依赖关系变化 | `relationships.md` | 更新依赖链 |

## Harness 载体导航

详见 `.omp/HARNESS-INDEX.md`。关键：
- 公理：`.omp/rules/axioms.md`
- 规则：`.omp/rules/quality-gates.md`
- 不变量：`.omp/memory/invariants.md` + `tests/invariants/`
- 进化引擎：`tools/harness-l5/evolve.js`
- 自举检查：`tools/check-harness-l5.js`
