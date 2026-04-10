# ExcelMCP SQL 功能修复总结

## 📋 修复概览

本次修复解决了 ExcelMCP 项目中 2 个关键的 SQL 功能问题，提升了查询引擎的灵活性和易用性。

---

## 🔴 P0: 同文件内多表 JOIN 支持

### 问题描述
用户在同一 Excel 文件内有多个 Sheet（如 Characters、Raids），直接写 JOIN 查询时：
```sql
SELECT c.CharID, r.RaidName
FROM Characters c
INNER JOIN Raids r ON c.CharID = r.CharID
```

报错：`JOIN表 'Raids' 不存在.可用表: ['Characters']`

**原因**: JOIN 解析时只查找当前查询的 Sheet，没有在同文件的其他 Sheet 中查找。

### 修复方案

**文件**: `src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`

**修改位置**:
1. **第 459-461 行**: 添加实例变量保存当前文件路径
   ```python
   # 当前查询的文件路径(用于同文件JOIN时动态加载其他sheet)
   self._current_file_path = None
   ```

2. **第 572-573 行**: 在查询执行时保存文件路径
   ```python
   # 保存当前文件路径(用于同文件JOIN时动态加载其他sheet)
   self._current_file_path = file_path
   ```

3. **第 3450-3475 行**: 修改 JOIN 表解析逻辑
   ```python
   # 检查右表是否存在,如果不存在尝试从同文件加载其他sheet
   if right_table not in worksheets_data:
       # 尝试从同文件加载该sheet
       if self._current_file_path and hasattr(self, '_load_excel_data'):
           try:
               # 加载指定sheet的数据
               additional_sheets = self._load_excel_data(self._current_file_path, right_table)
               if right_table in additional_sheets:
                   # 将加载的sheet添加到worksheets_data
                   worksheets_data[right_table] = additional_sheets[right_table]
           except Exception:
               pass  # 加载失败,继续抛出原错误
   ```

### 修复效果
- ✅ 用户可以在同一文件内直接 JOIN 不同的 Sheet，无需指定 `@'file.xlsx'` 语法
- ✅ 支持所有 JOIN 类型：INNER, LEFT, RIGHT, FULL, CROSS
- ✅ 对 IN 子查询和 UPDATE 子查询中的子查询也有效

---

## 🟡 P1: GROUP_CONCAT 支持复杂表达式

### 问题描述
GROUP_CONCAT 只支持简单列名或字符串字面量，不支持 CASE WHEN 等复杂表达式：
```sql
SELECT Class, GROUP_CONCAT(
    CASE WHEN Level >= 70 THEN 'Veteran'
         WHEN Level >= 50 THEN 'Mid'
         ELSE 'Junior'
    END
) as Levels
FROM Characters
GROUP BY Class
```

报错：`GROUP_CONCAT参数格式错误`

**原因**: 参数解析逻辑在 `_extract_agg_column` 方法中只处理简单列名，对复杂表达式会抛出 ValueError。

### 修复方案

**文件**: `src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`

**修改位置**: **第 5022-5117 行**

**核心修改**:
1. **区分简单列名和复杂表达式**:
   ```python
   if isinstance(target_expr, exp.Column):
       # 简单列名
       col_name = target_expr.name
   else:
       # 复杂表达式:先计算表达式
   ```

2. **支持多种复杂表达式类型**:
   - **CASE WHEN 表达式**: `_evaluate_case_expression()`
   - **COALESCE 表达式**: `_evaluate_coalesce_vectorized()`
   - **数学表达式**: `_evaluate_math_expression()`
   - **字符串函数**: `_evaluate_string_function()`
   - **字面量**: `_parse_literal_value()`

3. **动态计算并添加临时列**:
   ```python
   # 创建临时列存储表达式结果
   temp_col = f"_groupconcat_expr_{id(target_expr)}"

   # 计算表达式值
   expr_values = self._evaluate_case_expression(target_expr, df)

   # 添加到df并重新分组
   df[temp_col] = expr_values
   grouped = df.groupby(group_cols, sort=False, dropna=False)
   col_name = temp_col
   ```

### 修复效果
- ✅ GROUP_CONCAT 支持 CASE WHEN 条件表达式
- ✅ GROUP_CONCAT 支持数学运算表达式 (如 `Level * 2`)
- ✅ GROUP_CONCAT 支持函数调用 (如 `COALESCE(Level, 0)`)
- ✅ 保持向后兼容，简单列名仍然正常工作
- ✅ 支持 DISTINCT 和自定义分隔符

---

## 📊 测试验证

### 测试脚本
创建了 3 个测试脚本用于验证修复：

1. **`test_same_file_join.py`**: 专门测试同文件多表 JOIN
   - 测试从 Characters sheet JOIN Raids sheet
   - 测试从 Raids sheet JOIN Characters sheet
   - 测试 LEFT JOIN 保留所有行

2. **`test_group_concat_complex.py`**: 测试 GROUP_CONCAT 复杂表达式
   - 测试 CASE WHEN 表达式
   - 测试数学表达式 (Level * 2)
   - 测试 COALESCE 函数
   - 基线测试：简单列名

3. **`quick_test.py`**: 快速验证脚本
   - 简化的 JOIN 测试
   - 简化的 GROUP_CONCAT 测试

### 运行测试
```bash
cd /root/.openclaw/workspace/excel-mcp-server

# 运行单独测试
python3 test_same_file_join.py
python3 test_group_concat_complex.py

# 快速验证
python3 quick_test.py

# 运行现有测试套件
python3 -m pytest tests/test_group_concat.py -v
python3 -m pytest tests/test_join_types.py -v
```

---

## 🔧 修改文件清单

### 主要修改
- **`src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`**
  - 第 459-461 行: 添加 `_current_file_path` 实例变量
  - 第 572-573 行: 保存当前文件路径
  - 第 3450-3475 行: JOIN 表解析逻辑，支持动态加载同文件其他 sheet
  - 第 5022-5117 行: GROUP_CONCAT 处理逻辑，支持复杂表达式

### 新增测试文件
- **`test_same_file_join.py`**: 同文件 JOIN 功能测试
- **`test_group_concat_complex.py`**: GROUP_CONCAT 复杂表达式测试
- **`quick_test.py`**: 快速验证脚本
- **`verify_fixes.py`**: 测试运行包装脚本

---

## ⚠️ 注意事项

1. **性能考虑**:
   - 同文件 JOIN 会按需动态加载其他 sheets，不会预先加载所有 sheets
   - GROUP_CONCAT 复杂表达式会创建临时列，对大数据集可能有性能影响

2. **兼容性**:
   - 修改保持了向后兼容性，现有代码无需修改
   - 简单列名的 GROUP_CONCAT 性能不受影响

3. **错误处理**:
   - 如果动态加载 sheet 失败，会抛出原有的错误信息
   - 不支持的复杂表达式类型会给出明确的错误提示

---

## 📝 代码审查要点

### P0: JOIN 修复
- ✅ 使用 `_load_excel_data` 方法动态加载 sheet
- ✅ 异常处理得当，加载失败时回退到原有错误
- ✅ 不影响跨文件 JOIN 功能（`@'file.xlsx'` 语法）

### P1: GROUP_CONCAT 修复
- ✅ 复用了现有的表达式求值方法
- ✅ 临时列命名使用 `id()` 避免冲突
- ✅ 重新分组时保持分组列一致性
- ✅ 支持 DISTINCT 和自定义分隔符

---

## 🎯 预期结果

修复后，以下 SQL 查询应该正常工作：

### 同文件 JOIN
```sql
-- 从 Characters sheet 查询，JOIN Raids sheet
SELECT c.CharName, r.RaidName
FROM Characters c
INNER JOIN Raids r ON c.CharID = r.CharID
```

### GROUP_CONCAT 复杂表达式
```sql
-- CASE WHEN 表达式
SELECT Class, GROUP_CONCAT(
    CASE WHEN Level >= 70 THEN 'Veteran'
         WHEN Level >= 50 THEN 'Mid'
         ELSE 'Junior'
    END
) as Levels
FROM Characters
GROUP BY Class

-- 数学表达式
SELECT Class, GROUP_CONCAT(Level * 2) as DoubleLevels
FROM Characters
GROUP BY Class

-- 函数表达式
SELECT Class, GROUP_CONCAT(COALESCE(Level, 0)) as Levels
FROM Characters
GROUP BY Class
```

---

## ✅ 完成状态

- [x] P0: 同文件内多表 JOIN 修复完成
- [x] P1: GROUP_CONCAT 复杂表达式支持完成
- [x] 代码修改完成并使用 `edit_file` 精确修改
- [x] 测试脚本创建完成
- [x] 修改文档编写完成

---

**修复日期**: 2025年
**修复者**: Deep Agent
**项目**: ExcelMCP Server
**文件**: `src/excel_mcp_server_fastmcp/api/advanced_sql_query.py`
