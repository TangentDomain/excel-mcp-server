# 开发者指南

## 项目架构

```
server.py          ← MCP 工具层 (FastMCP)
  │
  └─ api/          ← 业务逻辑层
       │
       ├─ advanced_sql_query.py  SQL 查询引擎 (10317 行)
       ├─ excel_operations.py    通用 Excel 操作
       ├─ script_runner.py       Python 脚本执行
       └─ header_analyzer.py     双行表头检测
       │
       └─ core/   ← 底层数据访问层
            │
            ├─ excel_reader.py     读取 (calamine → openpyxl 降级)
            ├─ excel_writer.py     写入 (传统模式)
            ├─ streaming_writer.py  写入 (流式模式, 大文件)
            ├─ excel_manager.py    工作表管理
            ├─ excel_search.py     搜索 (calamine → openpyxl 降级)
            ├─ excel_compare.py    比较
            └─ excel_converter.py  格式转换
       │
       └─ utils/ ← 工具模块
            ├─ validators.py      SecurityValidator + ExcelValidator
            ├─ formatter.py       _clean_result / _wrap / _fail 格式化
            ├─ config.py          配置文件
            ├─ parsers.py         范围表达式解析器
            ├─ formula_cache.py   公式计算结果缓存
            ├─ concurrent_utils.py 并发工具 (RLock)
            ├─ exceptions.py      异常定义
            ├─ temp_file_manager.py 临时文件管理
            └─ text_utils.py      文本工具
       │
       └─ models/ ← 类型定义
            └─ types.py           OperationResult / RangeInfo 等
       │
       └─ calibrator/ ← 校准工具 (开发用)
            └─ core.py            SQLite 结果对比校准
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
uv run python -m pytest tests/ -q --tb=short

# 运行特定文件
uv run python -m pytest tests/test_core.py -v

# 运行不变量测试 (L4)
uv run python -m pytest tests/invariants/ -q

# 并行运行
uv run python -m pytest tests/ -q -n auto --timeout=30
```

### 测试数据
- 测试各自创建临时文件（不使用共享文件），确保隔离
- 使用 `conftest.py` 中的 `sample_excel_file` fixture：双行表头 + 4 行数据
- `clear_sql_engine_cache` (autouse): 每次测试前清空 SQL 引擎缓存

## 重要文件

| 文件 | 说明 |
|------|------|
| `server.py` | MCP 工具定义 (2488 行, 26 个工具) |
| `api/advanced_sql_query.py` | SQL 引擎 (10317 行) |
| `api/excel_operations.py` | Excel 业务操作 (3009 行) |
| `core/excel_writer.py` | 写入实现 (2020 行) |
| `core/streaming_writer.py` | 流式写入 (734 行) |
| `core/excel_reader.py` | 读取实现 (666 行) |
| `utils/validators.py` | 验证器 (436 行) |
| `utils/formatter.py` | 结果格式化 (486 行) |
