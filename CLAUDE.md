# Excel MCP Server

游戏开发专用 Excel 配置表 MCP 服务器。SQL-over-Excel 引擎，35 个工具，支持高级 SQL 查询、批量操作、跨文件 JOIN。

## 技术栈

- Python 3.10+ / FastMCP / openpyxl / sqlglot
- 测试: pytest + pytest-timeout + pytest-xdist
- Lint: ruff (format + check)
- 版本: 1.16.0

## 目录结构

```
src/excel_mcp_server_fastmcp/   # 主包
  api/       # SQL 执行引擎（query/update/insert/delete）
  core/      # Excel 操作核心（read/write/backup）
  models/    # 数据模型
  utils/     # 工具函数
  calibrator/ # SQLite 交叉校验工具
  verification/ # baseline 驱动验证
tests/
  invariants/  # 不变量测试（INV-1~32）
  test_data/  # 测试数据
scripts/
  pre_commit_check.sh  # 提交前检查
docs/
  SQL_PRINCIPLES.md    # SQL 标准遵循原则
```

## MCP 工具 → Python API 映射

| MCP 工具 | Python API |
|---------|-----------|
| excel_query | `execute_advanced_sql_query` |
| excel_update_query | `execute_advanced_update_query` |
| excel_insert_query | `execute_advanced_insert_query` |
| excel_delete_query | `execute_advanced_delete_query` |

## SQL 核心原则

**只支持 SQL 标准支持的功能。** SQLite 3.x 为真值来源，calibrator 交叉校验。

- 支持: 窗口函数、JOIN、聚合、CTE、子查询
- 不支持: WHERE 引用窗口函数别名、SELECT 别名在 WHERE 中
- SQL 执行顺序: FROM → JOIN → WHERE → GROUP BY → HAVING → 窗口函数 → SELECT → ORDER BY → LIMIT

## API 契约

所有工具返回 `{success: bool, data: list, message: str}`。
- `success=False` 时 `data` 为空列表，`message` 非空且不含堆栈
- `affected_rows` 必须精确等于实际变更行数

## 不变量体系

32 条不变量，三层分层。详见 `.claude/skills/INVARIANT-DRIVEN-DEV.md`。

```
L1 外部真值（4条）: INV-1 结果结构 / INV-2 SQL对齐 / INV-3 文件完整 / INV-4 行数守恒
L2 架构原则（5条）: INV-5~9 失败安全/错误分类/幂等读取/LIMIT/聚合语义
L3 具体不变量（23条）: INV-10~32 窗口函数/空表/写操作/SQL边界
```

### 判据执行

```bash
python -m pytest tests/invariants/ -q          # 全量（<60s）
python -m pytest tests/invariants/ -k "smoke"   # 烟雾（<10s）
```

当前状态: 154 passed, 3 skipped, 收敛评级 C。

## 代码规范

- 向量化优先: `pd.Series` 而非逐行
- 分发表用 `tuple` 不用 `frozenset`
- 异常回退: 类型转换失败时逐行 fallback
- 版本更新: `__init__.py` + `pyproject.toml` 同步

## 部署检查

```bash
ruff check src/ tests/ && ruff format --check src/ tests/
python -m pytest tests/invariants/ -q --tb=short --timeout=30
```

## Enforcement

CLAUDE.md 是 advisory。Hard limits 在：

- `.claude/settings.json` — permissions deny/ask/allow + hooks
- `~/.claude/hooks/preToolUse-bash-guard.sh` — regex 拦截危险命令（exit 2 block）
- `~/.claude/hooks/postToolUse-audit.sh` — 工具调用 JSONL 审计日志
- `~/.claude/hooks/stop-lint.sh` — 每轮结束自动 lint 检查
- `scripts/pre_commit_check.sh` — git commit 门禁（docstring + ruff + pytest）
