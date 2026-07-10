# 开发者指南

## 项目架构

```
src/excel_mcp_server_fastmcp/
  server.py                    MCP 工具层 (FastMCP) — 26 个工具, 2493 行
  cli.py                       CLI 入口 — 27 个子命令, 支持自更新
  │
  api/                         业务逻辑层
    advanced_sql_query.py      SQL 查询引擎 (10395 行)
    excel_operations.py        通用 Excel 操作 (2776 行)
    script_runner.py           Python 脚本沙箱 (281 行)
    header_analyzer.py         双行表头检测
    query_helpers.py           错误提示生成
  │
  core/                        底层数据访问层
    excel_reader.py            读取 (calamine → openpyxl 降级, 642 行)
    excel_writer.py            写入 (传统模式, 1948 行)
    excel_manager.py           工作表管理
    excel_search.py            搜索 (calamine → openpyxl 降级)
    excel_compare.py           比较
    excel_converter.py         格式转换
  │
  utils/                       工具模块
    validators.py              SecurityValidator + ExcelValidator
    formatter.py               结果格式化 (_wrap / _fail / _clean)
    formula_cache.py           公式计算缓存
    concurrent_utils.py        并发工具 (RLock)
    text_utils.py              文本工具
  │
  models/                      类型定义
    types.py                   OperationResult / RangeInfo 等
  │
  calibrator/                  校准工具 (开发调试用)
    core.py                    SQLite 结果对比校准
  │
  verification/                baseline 驱动验证
    runner.py                  闭环验证运行器
    scenarios.py               验证场景定义
    diff.py                    结果差异比较
```

## 关键设计决策

### 双表头支持
Excel 游戏配置表通常有双行表头：第1行中文描述，第2行英文字段名。
所有 SQL 查询接口自动识别并支持两种列名。

### 性能路径
- `calamine` (Rust 引擎): 纯数据读取/搜索，10-50x 快于 openpyxl
- `openpyxl`: 格式化读取/公式读取，作为 calamine 不可用时的降级方案
- `StreamingWriter`: 大文件修改用 calamine 读取 + write_only 写入，内存与文件大小无关

### 安全
- 所有 26 个 MCP 工具使用 `@_validate_file_path()` 装饰器验证文件路径
- 路径穿越、符号链接、隐藏文件、非法扩展名均在 SecurityValidator 中拦截
- SQL 通过 sqlglot AST 解析，非字符串拼接
- `excel_run_python` 使用沙箱环境，限制文件/进程操作，白名单内置函数

### SQL 引擎
- 基于 sqlglot AST 解析，非字符串拼接
- 与 SQLite 3.x 交叉校验（calibrator），169 条差分测试 100% 通过
- 支持 WHERE 引用窗口函数别名（自动重写为子查询）
- 支持 WHERE 引用 SELECT 别名（物化为临时列）

## 添加新 MCP 工具

```python
@mcp.tool()
@_validate_file_path()  # 路径验证
@_track_call            # 调用追踪
def excel_new_tool(file_path: str, param: str = "") -> dict:
    """工具 docstring（LLM 路由用）

    Args:
        file_path: 文件路径
        param: 参数说明
    """
    if not param:
        return _fail("参数不能为空", meta={"error_code": "MISSING_PARAM"})
    result = ExcelOperations.new_tool(file_path, param)
    return _wrap(result)
```

## 测试

```bash
# 运行全部测试
uv run python -m pytest tests/ -q --timeout=60

# 运行不变量测试 (L1-L4)
uv run python -m pytest tests/invariants/ -q --timeout=30

# 运行特定文件
uv run python -m pytest tests/test_core.py -v

# 并行运行
uv run python -m pytest tests/ -q -n auto --timeout=30
```

### 测试体系

| 层级 | 目录 | 说明 |
|------|------|------|
| L1 结果结构 | `tests/invariants/test_l1_result_structure.py` | API 返回格式不变量 |
| L2 架构 | `tests/invariants/test_l2_architecture.py` | 代码结构不变量 |
| L3 SQL 功能 | `tests/invariants/test_l3_*.py` | SQL 准确率差分测试 |
| L4 限制消除 | `tests/invariants/test_l4_limit_fixes.py` | 引擎限制修复验证 |
| 对抗测试 | `tests/adversarial/` | 随机 fuzz 读写 |
| 功能测试 | `tests/test_*.py` | 各模块功能测试 |

### 测试数据
- 测试各自创建临时文件（不使用共享文件），确保隔离
- 使用 `conftest.py` 中的 `sample_excel_file` fixture：双行表头 + 4 行数据
- `clear_sql_engine_cache` (autouse): 每次测试前清空 SQL 引擎缓存

## 重要文件

| 文件 | 行数 | 说明 |
|------|------|------|
| `server.py` | 2493 | MCP 工具定义, 26 个工具 |
| `api/advanced_sql_query.py` | 10395 | SQL 引擎 |
| `api/excel_operations.py` | 2776 | Excel 业务操作 |
| `core/excel_writer.py` | 1948 | 写入实现 |
| `core/excel_reader.py` | 642 | 读取实现 |
| `api/script_runner.py` | 281 | Python 脚本沙箱 |
| `cli.py` | 1285 | CLI 入口, 27 个子命令 |
