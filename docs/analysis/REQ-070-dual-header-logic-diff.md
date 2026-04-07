# REQ-070: 双行表头识别逻辑差异分析

## 两个函数概览

| 维度 | `_detect_dual_header` (server.py:2507) | `_parse_dual_header_data` (excel_operations.py:532) |
|------|---------------------------------------|-----------------------------------------------------|
| **调用者** | `describe_table` | `get_headers` |
| **职责** | **检测**是否为双行表头 | **解析**已确认的双行表头数据 |
| **输入** | `rows` (原始行数据列表) | `data` (二维数组) + `max_columns` |
| **核心差异** | 有检测逻辑，判断是否双行 | **无条件视为双行**，不做检测 |

## 关键差异

### 1. 检测 vs 不检测（根本差异）

**`_detect_dual_header`** 有完整的检测逻辑：
- 第二行所有非空单元格必须是英文字母开头的字符串
- 第一行必须包含至少一个中文字符
- 至少3列

**`_parse_dual_header_data`** **不做任何检测**，它假设调用者已经确认是双行表头，直接按双行模式解析。

### 2. get_headers 的调用路径

查看 `get_headers`（excel_operations.py:354 附近），它是在 `dual_row=True` 参数下调用 `_parse_dual_header_data` 的。而 `dual_row` 参数的来源需要看调用链——**get_headers 本身不做双行检测**，它依赖外部传入的 `dual_row` 参数。

### 3. describe_table 的调用路径

`describe_table`（server.py:2775）调用 `_detect_dual_header` 做自动检测，如果检测为双行，则 `header_row_idx=1`，否则 `header_row_idx=0`。

## 问题根源

**同一个文件，两个工具对"是否为双行表头"的判断可能不同：**

1. **`describe_table`** 用 `_detect_dual_header` 自动检测 → 可能判断为单行或双行
2. **`get_headers`** 不自动检测 → 取决于调用者是否传 `dual_row=True`，如果用户不传或传错，结果就不同

**具体不一致场景：**
- 第二行是英文标识但不是纯字母开头（如 `2_name`）→ `_detect_dual_header` 判定为单行，但 `get_headers` 可能仍按双行解析
- 第一行无中文但有双行表头结构 → `_detect_dual_header` 判定为单行，丢失描述信息
- 空值处理不同：`_detect_dual_header` 跳过 None，`_parse_dual_header_data` 给 fallback 值

## 建议

1. **统一检测逻辑**：将 `_detect_dual_header` 提取为公共方法，`get_headers` 也调用它
2. **放宽检测条件**：当前要求第二行全部是英文字母开头过于严格，应考虑更通用的启发式
3. **`get_headers` 自动检测**：默认不传 `dual_row` 时自动检测，而非依赖外部参数
