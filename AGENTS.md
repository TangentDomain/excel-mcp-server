# Excel MCP Server - AGENTS.md

> 本文档面向AI编程助手，提供项目背景、架构设计和开发规范等关键信息。

## 项目概述

**Excel MCP Server** 是一个专为游戏开发设计的Excel配置表管理MCP服务器。基于**FastMCP**和**openpyxl**构建，通过AI自然语言指令实现Excel文件的智能化操作。

### 核心特性

- **38个专业工具**: 涵盖文件管理、数据操作、搜索分析、格式化、数据转换等
- **游戏开发专用**: 双行表头系统、配置表对比、ID重复检测等游戏行业特性
- **企业级安全**: 自动备份、操作日志、数据影响评估、撤销恢复能力
- **SQL优先查询**: 内置高级SQL查询引擎，支持复杂数据分析和聚合统计

### 项目统计

- 测试用例: 698个（通过率>99%）
- 代码覆盖率: ~78%
- 代码总行数: 13,500+行（仅测试代码）
- Python版本要求: >= 3.10

---

## 技术栈

### 核心依赖

| 依赖包 | 版本 | 用途 |
|--------|------|------|
| `fastmcp` | >=0.1.0 | MCP服务器框架 |
| `openpyxl` | >=3.1.0 | Excel文件操作核心库 |
| `mcp` | >=1.0.0 | MCP协议支持 |
| `xlcalculator` | >=0.5.0 | Excel公式计算引擎 |
| `formulas` | >=1.3.0 | 公式解析和执行 |
| `xlwings` | >=0.33.15 | Excel应用集成（Windows） |
| `psutil` | >=7.0.0 | 系统进程管理 |
| `sqlglot` | >=27.29.0 | SQL解析和执行引擎 |

### 开发依赖

- **测试框架**: pytest, pytest-asyncio, pytest-cov, pytest-mock, pytest-xdist, pytest-timeout, pytest-benchmark
- **代码质量**: black, isort, flake8, mypy
- **性能分析**: memory-profiler, psutil

### 包管理

项目使用 `uv` 作为推荐的包管理器，同时也支持 `pip`:

```bash
# 推荐：使用uv
pip install uv && uv sync

# 备选：使用pip
pip install -e ".[dev]"
```

---

## 项目架构

### 分层架构设计

```
┌─────────────────────────────────────────────────────────────┐
│                    MCP接口层 (server.py)                     │
│  - 纯委托模式，零业务逻辑                                    │
│  - @mcp.tool()装饰器定义38个工具                             │
│  - 操作日志记录和会话管理                                    │
└──────────────────────────┬──────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│              API业务逻辑层 (api/excel_operations.py)         │
│  - ExcelOperations类：高内聚的业务操作封装                   │
│  - 统一参数验证、错误处理、结果格式化                        │
│  - 主干API和分支实现分离                                     │
└──────────────────────────┬──────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│                  核心操作层 (core/)                          │
│  - excel_reader.py: Excel文件读取和范围数据获取              │
│  - excel_writer.py: Excel文件写入和修改操作                  │
│  - excel_manager.py: 文件和工作表管理                        │
│  - excel_search.py: 正则搜索和目录批量搜索                   │
│  - excel_compare.py: 配置表对比（游戏开发专用）              │
│  - excel_converter.py: 格式转换和数据导入导出                │
└──────────────────────────┬──────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│                    工具层 (utils/)                           │
│  - error_handler.py: 统一错误处理和响应格式化                │
│  - exceptions.py: 自定义异常类定义                           │
│  - validators.py: 参数验证和数据校验                         │
│  - parsers.py: 范围表达式解析                                │
│  - formatter.py: 结果格式化                                  │
│  - formula_cache.py: 公式缓存管理                            │
│  - temp_file_manager.py: 临时文件管理                        │
└─────────────────────────────────────────────────────────────┘
```

### 目录结构

```
src/
├── server.py                    # MCP服务器入口（约2000行，38个工具）
├── api/
│   ├── __init__.py
│   ├── excel_operations.py      # 主API类（约1800行）
│   └── advanced_sql_query.py    # 高级SQL查询引擎
├── core/
│   ├── __init__.py
│   ├── excel_reader.py          # Excel读取（约300行）
│   ├── excel_writer.py          # Excel写入
│   ├── excel_manager.py         # 文件和工作表管理
│   ├── excel_search.py          # 搜索功能
│   ├── excel_compare.py         # 配置表对比
│   └── excel_converter.py       # 格式转换
├── models/
│   ├── __init__.py
│   └── types.py                 # 数据类型定义（约230行）
└── utils/
    ├── __init__.py
    ├── error_handler.py         # 错误处理（约200行）
    ├── exceptions.py            # 自定义异常
    ├── validators.py            # 验证工具（约400行）
    ├── parsers.py               # 解析工具
    ├── formatter.py             # 格式化工具
    ├── formula_cache.py         # 公式缓存
    └── temp_file_manager.py     # 临时文件管理

tests/                           # 测试目录（49个测试文件）
├── conftest.py                  # pytest配置和fixture
├── test_server.py               # 服务器接口测试
├── test_api_*.py                # API层测试
├── test_core.py                 # 核心模块测试
├── test_excel_*.py              # 各功能模块测试
└── test_security_features.py    # 安全功能测试
```

---

## 代码规范

### 命名规范

- **类名**: PascalCase，如 `ExcelOperations`, `ExcelReader`
- **函数/方法**: snake_case，如 `get_range`, `update_range`
- **常量**: UPPER_SNAKE_CASE，如 `MAX_ROWS_OPERATION`
- **私有方法**: 下划线前缀，如 `_validate_range_format`
- **模块名**: 小写，如 `excel_reader.py`

### 代码组织原则

1. **纯委托模式**: `server.py` 严格委托给 `ExcelOperations`，不实现任何业务逻辑
2. **集中式处理**: API层统一处理验证、错误处理、结果格式化
3. **单一职责**: 每个核心模块只负责一类操作（读、写、管理、搜索等）
4. **1-Based索引**: 匹配Excel用户习惯，第1行=第一行

### 文档字符串规范

使用Google风格的文档字符串，包含：

```python
def example_function(param1: str, param2: int) -> Dict[str, Any]:
    """
    简要描述函数功能
    
    详细描述（可选）
    
    Args:
        param1: 参数1说明
        param2: 参数2说明
        
    Returns:
        返回值说明
        
    Raises:
        SomeException: 异常说明
        
    Example:
        result = example_function("test", 123)
    """
```

---

## 构建和测试命令

### 快速命令（Makefile）

```bash
# 运行测试
make test                      # 运行所有测试
make test-unit                 # 仅运行单元测试
make test-integration          # 仅运行集成测试
make test-performance          # 仅运行性能测试
make test-all                  # 运行所有测试并生成覆盖率报告

# 代码质量
make lint                      # 运行代码检查（flake8 + mypy）
make format                    # 格式化代码（black + isort）
make format-check              # 检查代码格式

# 覆盖率
make coverage                  # 生成覆盖率报告

# 清理
make clean                     # 清理测试产物和缓存

# CI/发布
make ci                        # 完整CI流程
make pre-release              # 发布前检查
```

### pytest命令

```bash
# 基础测试
python -m pytest tests/ -v

# 带覆盖率
python -m pytest tests/ --cov=src --cov-report=html

# 并行测试
python -m pytest tests/ -x -n auto

# 按标记运行
python -m pytest tests/ -m "not slow"           # 排除慢测试
python -m pytest tests/ -m "integration"        # 仅集成测试
python -m pytest tests/ -m "security"           # 仅安全测试
python -m pytest tests/ -m "performance"        # 仅性能测试

# 调试
python -m pytest tests/ -v -s --pdb             # 失败时进入pdb
python -m pytest tests/ --tb=long               # 详细错误信息
python -m pytest tests/ --maxfail=1             # 首次失败即停止
```

### 开发环境设置

```bash
# 安装开发依赖
make install-deps
# 或
pip install -e ".[dev,performance,quality]"

# 完整开发环境设置
make dev-setup
```

---

## 测试策略

### 测试标记分类

| 标记 | 说明 | 使用场景 |
|------|------|----------|
| `slow` | 慢速测试 | 需要较长时间运行的测试 |
| `integration` | 集成测试 | 涉及多个模块的集成测试 |
| `unit` | 单元测试 | 单一模块的单元测试 |
| `performance` | 性能测试 | 基准性能测试 |
| `security` | 安全测试 | 安全功能测试 |
| `api` | API测试 | API层接口测试 |
| `core` | 核心测试 | 核心模块测试 |
| `utils` | 工具测试 | 工具函数测试 |

### 测试覆盖率要求

- 目标覆盖率: >78%
- 核心模块: 必须覆盖主要执行路径
- 新增功能: 必须包含对应的测试用例

### 测试文件命名

- 测试文件: `test_*.py`
- 测试类: `Test*` 
- 测试函数: `test_*`

---

## 安全特性

### 数据保护机制

1. **自动备份系统**
   - `excel_create_backup`: 创建带时间戳的备份
   - `excel_restore_backup`: 一键恢复
   - `excel_list_backups`: 备份管理

2. **操作日志系统**
   - `OperationLogger` 类管理操作会话
   - `excel_get_operation_history`: 审计跟踪
   - JSON格式日志存储在 `.excel_mcp_logs/`

3. **操作预览机制**
   - `excel_preview_operation`: 预览操作影响
   - 自动风险评估（LOW/MEDIUM/HIGH）
   - 安全建议和警告

4. **数据影响评估**
   - `excel_assess_data_impact`: 全面评估操作影响
   - 分析数据类型、密度、公式等因素
   - 智能风险预测

### 默认安全行为

- `insert_mode=True`: 默认插入模式防止数据覆盖
- 严格参数验证: 所有输入都经过验证
- 规模限制: 阻止过大的操作（默认最大1000行、100列）

---

## 开发约定

### 分支命名规范

所有功能分支必须以 `feature/` 开头:

```bash
git checkout -b feature/add-new-tool
git checkout -b feature/fix-bug-123
```

### 提交信息规范

使用中文描述提交内容，格式：

```
类型: 简要描述

详细描述（可选）
```

类型包括: `功能`, `修复`, `优化`, `文档`, `测试`, `重构`

### 代码审查清单

- [ ] 是否遵循纯委托模式？
- [ ] 是否添加了适当的错误处理？
- [ ] 是否包含类型注解？
- [ ] 是否添加了文档字符串？
- [ ] 是否包含对应的测试？
- [ ] 测试覆盖率是否达标？
- [ ] 是否通过lint检查？

---

## 部署配置

### MCP客户端配置示例

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/excel-mcp-server",
        "run",
        "python",
        "-m",
        "src.server"
      ]
    }
  }
}
```

### 环境变量

- `PYTHONPATH`: 项目根目录路径
- `EXCEL_MCP_LOG_LEVEL`: 日志级别（DEBUG/INFO/WARNING/ERROR）

---

## 故障排除

### 常见问题

1. **文件被锁定**: 关闭Excel程序后重试
2. **中文乱码**: 确保UTF-8编码，检查Python环境编码
3. **大文件缓慢**: 使用精确范围，分批处理数据
4. **内存不足**: 减少单次处理数据量，及时关闭工作簿

### 调试技巧

```python
# 开启调试日志
ExcelOperations.DEBUG_LOG_ENABLED = True

# 查看操作历史
result = excel_get_operation_history("file.xlsx")
```

---

## 参考资料

- [README.md](README.md) - 项目主文档（中文）
- [README.en.md](README.en.md) - 英文文档
- [DEPLOYMENT.md](DEPLOYMENT.md) - 部署指南
- [CONTRIBUTING.md](CONTRIBUTING.md) - 贡献指南
- [SECURITY_IMPROVEMENTS_SUMMARY.md](SECURITY_IMPROVEMENTS_SUMMARY.md) - 安全改进总结
