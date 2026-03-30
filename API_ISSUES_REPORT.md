# Excel MCP API 问题测试报告

## 测试时间
2026-03-30

## 测试范围
监工报告中提到的5个API问题：
1. read_data_from_excel 范围查询参数顺序问题
2. format_range 缺少必要参数时的处理
3. apply_formula 缺少 formula 参数时的处理
4. read_data_from_excel 搜索逻辑问题
5. write_data_to_excel 数据格式不匹配时的处理

---

## 测试结果总览

| # | 问题 | 状态 | 严重程度 |
|---|------|------|---------|
| 1 | read_data_from_excel 范围查询参数顺序问题 | ✅ 无问题 | - |
| 2 | format_range 缺少必要参数时的处理 | ✅ 无问题 | - |
| 3 | apply_formula 缺少 formula 参数时的处理 | ❌ 有问题 | 🔴 高 |
| 4 | read_data_from_excel 搜索逻辑问题 | ✅ 无问题 | - |
| 5 | write_data_to_excel 数据格式不匹配时的处理 | ❌ 有问题 | 🟡 中 |

---

## 详细分析

### 问题 1: read_data_from_excel 范围查询参数顺序问题

**状态**: ✅ 无问题

**测试结果**:
- 使用 `range` 参数：成功
- 使用 `start_cell` + `end_cell`：成功
- 两种调用方式都正常工作

**函数签名**:
```python
excel_get_range(
    file_path: str,
    range: str,  # 必需
    include_formatting: bool = False,
    sheet_name: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None
) -> Dict[str, Any]
```

**说明**:
- `range` 是必需参数（第2个位置）
- 如果提供了 `start_cell` 和 `end_cell`，函数会自动构建 range 表达式
- 不存在参数顺序混乱的问题

---

### 问题 2: format_range 缺少必要参数时的处理

**状态**: ✅ 无问题

**测试结果**:
- 不提供 `formatting` 或 `preset` 参数：成功
- 返回成功消息："成功格式化3个单元格"
- 应用了默认值或执行无操作

**函数签名**:
```python
excel_format_cells(
    file_path: str,
    sheet_name: str,
    range: str,
    formatting: Optional[Dict[str, Any]] = None,
    preset: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None
) -> Dict[str, Any]
```

**说明**:
- `formatting` 和 `preset` 都是可选参数
- 函数在没有指定格式时也能正常工作
- 不存在缺少必要参数的问题

---

### 问题 3: apply_formula 缺少 formula 参数时的处理

**状态**: ❌ 有问题（🔴 高）

**测试结果**:
- 缺少 `formula` 参数时：抛出 `TypeError`
- 错误信息：`excel_set_formula() missing 1 required positional argument: 'formula'`
- **问题**：没有友好的错误提示和参数验证

**函数签名**:
```python
excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str  # 必需参数
) -> Dict[str, Any]
```

**当前实现** (server.py:2071-2081):
```python
def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """在单元格写入Excel公式。"""
    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err
    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return _fail(f'🔒 安全验证失败: {_formula_err["error"]}', meta={"error_code": "FORMULA_SECURITY_FAILED"})
    return _wrap(ExcelOperations.set_formula(file_path, sheet_name, cell_address, formula))
```

**问题原因**:
1. 没有在函数内部验证 `formula` 参数是否为空或无效
2. 直接将参数传递给下层函数，如果参数缺失，Python 会抛出 TypeError
3. 错误信息不友好，不符合 MCP 工具的错误处理规范

**修复方案**:
在 `excel_set_formula` 函数开头添加参数验证：

```python
def excel_set_formula(
    file_path: str,
    sheet_name: str,
    cell_address: str,
    formula: str
) -> Dict[str, Any]:
    """在单元格写入Excel公式。"""
    # 新增：参数验证
    if not formula or not formula.strip():
        return _fail(
            '公式不能为空，请提供有效的Excel公式（如 "=A1+B1"）',
            meta={"error_code": "MISSING_FORMULA"}
        )

    _path_err = _validate_path(file_path)
    if _path_err:
        return _path_err

    _formula_err = SecurityValidator.validate_formula(formula)
    if not _formula_err['valid']:
        return _fail(f'🔒 安全验证失败: {_formula_err["error"]}', meta={"error_code": "FORMULA_SECURITY_FAILED"})

    return _wrap(ExcelOperations.set_formula(file_path, sheet_name, cell_address, formula))
```

---

### 问题 4: read_data_from_excel 搜索逻辑问题

**状态**: ✅ 无问题

**测试结果**:
- 搜索 "Alice"：成功找到匹配项
- 返回详细信息：位置、值、匹配范围等
- 搜索逻辑正常工作

**函数签名**:
```python
excel_search(
    file_path: str,
    pattern: str,
    sheet_name: Optional[str] = None,
    case_sensitive: bool = False,
    whole_word: bool = False,
    use_regex: bool = False,
    include_values: bool = True,
    include_formulas: bool = False,
    range: Optional[str] = None
) -> Dict[str, Any]
```

**说明**:
- 搜索功能完全正常
- 没有参数顺序混乱或逻辑错误的问题

---

### 问题 5: write_data_to_excel 数据格式不匹配时的处理

**状态**: ❌ 有问题（🟡 中）

**测试结果**:
- 正确格式（二维数组 `[["Test", 100]]`）：成功写入
- 错误格式（一维数组 `["Wrong", "Format"]`）：**也成功写入**（但行为不正确）
- 问题：函数没有验证数据格式，导致错误格式的数据被错误处理

**函数签名**:
```python
ExcelOperations.update_range(
    file_path: str,
    range_expression: str,
    data: List[List[Any]],  # 应该是二维数组
    preserve_formulas: bool = True,
    insert_mode: bool = False,
    streaming: bool = True
) -> Dict[str, Any]
```

**当前实现** (excel_operations.py:118-167):
```python
def update_range(
    cls,
    file_path: str,
    range_expression: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = False,
    streaming: bool = True
) -> Dict[str, Any]:
    """
    @intention 更新Excel文件中指定范围的数据，支持插入和覆盖模式

    Args:
        ...
        data: 二维数组数据 [[row1], [row2], ...]
        ...
    """
    # ... 没有验证 data 的格式 ...
```

**问题原因**:
1. Docstring 声明 `data` 应该是二维数组 `[[row1], [row2], ...]`
2. 但函数内部没有验证 `data` 的格式
3. 如果传入一维数组（如 `["Wrong", "Format"]`），StreamingWriter 会把每个字符当作一个值来处理
4. 导致数据被错误写入，但没有明确的错误提示

**错误示例**:
```python
# 错误用法
ExcelOperations.update_range(
    file_path="test.xlsx",
    range_expression="Sheet1!A6:B6",
    data=["Wrong", "Format"]  # 一维数组
)

# 实际结果：
# 第6行写入: "W", "r", "o", "n", "g" (5个字符被当作5列)
# 第7行写入: "F", "o", "r", "m", "a", "t" (6个字符被当作6列)
# 而不是预期的第6行写入: "Wrong", "Format"
```

**修复方案**:
在 `ExcelOperations.update_range` 函数开头添加数据格式验证：

```python
def update_range(
    cls,
    file_path: str,
    range_expression: str,
    data: List[List[Any]],
    preserve_formulas: bool = True,
    insert_mode: bool = False,
    streaming: bool = True
) -> Dict[str, Any]:
    """
    @intention 更新Excel文件中指定范围的数据，支持插入和覆盖模式

    Args:
        ...
        data: 二维数组数据 [[row1], [row2], ...]
        ...
    """
    # 新增：数据格式验证
    if not data:
        return cls._format_error_result("数据不能为空")

    if not isinstance(data, list):
        return cls._format_error_result(
            "数据格式错误：data 应该是二维数组 [[row1], [row2], ...]，"
            f"实际收到类型: {type(data).__name__}"
        )

    # 验证每一行是否都是列表
    for i, row in enumerate(data):
        if not isinstance(row, list):
            return cls._format_error_result(
                f"数据格式错误：第 {i+1} 行应该是列表，实际收到类型: {type(row).__name__}。"
                "data 应该是二维数组 [[row1], [row2], ...]"
            )

    # 验证通过后继续原有逻辑
    if cls.DEBUG_LOG_ENABLED:
        logger.info(f"{cls._LOG_PREFIX} 开始更新范围数据: {range_expression}, 模式: {'插入' if insert_mode else '覆盖'}")

    try:
        # ... 原有逻辑 ...
```

---

## 修复优先级

### 高优先级（🔴）
1. **apply_formula 缺少 formula 参数时的处理**
   - 位置：`server.py:2071-2081`
   - 修复：添加参数验证
   - 预计工作量：5分钟

### 中优先级（🟡）
2. **write_data_to_excel 数据格式不匹配时的处理**
   - 位置：`excel_operations.py:118-167`
   - 修复：添加数据格式验证
   - 预计工作量：10分钟

---

## 总结

1. **5个问题中，只有2个是真实问题**（#3 和 #5）
2. 其他3个问题（#1、#2、#4）在当前代码中不存在或已修复
3. **问题3**：缺少友好的错误提示，会导致 TypeError
4. **问题5**：缺少数据格式验证，会导致错误的数据写入行为
5. 两个问题都可以通过简单的参数验证修复

---

## 测试环境

- Python 版本：3.x
- Excel MCP Server 版本：v1.6.x
- 测试文件：/tmp/test_api_problems.xlsx
- 测试脚本：test_api_issues_v2.py

---

## 附录：完整测试代码

见 `test_api_issues_v2.py`
