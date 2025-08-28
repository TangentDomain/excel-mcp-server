# Excel MCP Server - AI编码助手指南

基于FastMCP和openpyxl的Excel文件MCP服务器，支持AI通过自然语言操作Excel文件。

## 核心架构

### 模块组织
- **src/server.py**: 纯MCP工具定义，委托给ExcelOperations
- **src/api/excel_operations.py**: 集中化业务逻辑处理中心
- **src/core/**: Excel操作模块(reader, writer, manager, search, compare)
- **src/utils/**: 格式化器和工具函数
- **src/models/**: 类型定义和数据模型

### 设计模式

#### 纯委托架构(重构后)
所有MCP工具使用简单委托：
```python
@mcp.tool()
def excel_list_sheets(file_path: str) -> Dict[str, Any]:
    return ExcelOperations.list_sheets(file_path)
```

#### 统一结果格式
```python
{
    'success': bool,
    'data': Any,        # 核心数据
    'message': str,
    'metadata': dict    # 附加上下文
}
```

#### 范围表达式支持
- 带工作表: `"Sheet1!A1:C10"` 或 `"TrSkill!A1:Z100"`
- 行范围: `"Sheet1!1:5"` 或 `"3:8"`
- 列范围: `"Sheet1!A:C"` 或 `"B:E"`
- 单行/列: `"Sheet1!5"` 或 `"C"`

## 开发工作流

### 运行和测试
```bash
# 开发运行
python -m src.server

# 完整测试(221个)
pytest tests/ -v

# 模块测试
pytest tests/test_api_excel_operations.py -v
pytest tests/test_core.py -v
```

### MCP客户端配置
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "python",
      "args": ["-m", "src.server"],
      "env": {"PYTHONPATH": "${workspaceRoot}"}
    }
  }
}
```

### 测试策略
- **221个测试，100%通过率**
- **分层测试**: API/核心/MCP接口测试
- **Mock隔离**: API层使用Mock确保独立性
- 关键文件: `test_api_excel_operations.py`(API), `test_core.py`(核心), `test_server.py`(MCP)

## 项目特色

### 游戏开发特化
- **Excel配置表比较**: 专为游戏配置设计
- **ID对象跟踪**: 检测新增/修改/删除对象
- **紧凑数组格式**: 优化游戏数据传输
- **TrSkill表分析**: 技能配置专项比较

### 中文/Unicode全支持
- 中文工作表名处理和编码
- Unicode标准化文本处理
- 本地化Excel特性回退机制

### Excel操作约定
- **1-based索引**匹配Excel惯例
- **默认保留公式**(`preserve_formulas=True`)
- 支持`.xlsx`和`.xlsm`格式
- 游戏表结构感知能力

## 关键依赖
- **FastMCP**: MCP服务器框架
- **openpyxl**: 核心Excel文件操作
- **xlcalculator/formulas**: 公式评估引擎
- **xlwings**: 可选Excel应用集成
- **pytest/pytest-asyncio**: 测试框架

## 常用操作

### 文件和工作表管理
文件创建、工作表CRUD、中文名支持、活动表自动管理

### 数据操作
基于范围的读写、格式保留、行列插入删除、单元格格式化预设

### 搜索分析
正则搜索、目录批量搜索、游戏配置表比较、公式评估

开发时始终使用统一的委托模式，将实现委托给核心模块，保持一致的结果格式。
