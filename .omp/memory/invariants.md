# 不变量体系（INV-1 ~ INV-32）

> 32 条不变量，从 4 条公理（AX-001~004）推导。每条 INV 关联测试文件和执行机制。

## AX → INV 映射表

| 公理 | 推导的不变量 |
|------|-------------|
| AX-001 SQL 标准真值 | INV-2, INV-10~15, INV-25~32 |
| AX-002 写操作安全 | INV-3, INV-16~18, INV-19, INV-20, INV-22~24 |
| AX-003 失败安全可诊断 | INV-1, INV-5, INV-6 |
| AX-004 查询幂等确定 | INV-4, INV-7, INV-8, INV-9, INV-21 |

---

## L1: 外部真值不变量（INV-1 ~ INV-4）

测试文件: `tests/invariants/test_l1_result_structure.py`

### INV-1: 结果结构一致性
- **声明**: result 必须包含 success(bool), data(list), message(str)
- **所属 AX**: AX-003（失败安全可诊断）
- **测试文件**: test_l1_result_structure.py → `TestINV1ResultStructure`
- **执行机制**: pytest 直接断言 key 存在性和类型

### INV-2: SQL-SQLite 结果对齐
- **声明**: 同一 SQL 在 ExcelMCP 和 SQLite 上的结果一致
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l1_result_structure.py → `TestINV2SQLSQLiteAlignment`
- **执行机制**: calibrator 导入 Excel → SQLite，对比两边 query 结果

### INV-3: 文件完整性守恒
- **声明**: SELECT 不修改文件；写操作只改目标 sheet
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l1_result_structure.py → `TestINV3FileIntegrity`
- **执行机制**: 操作前后计算文件哈希，对比非目标 sheet 不变

### INV-4: 行数守恒
- **声明**: COUNT(*) 返回的行数 = 实际数据行数（不含表头）
- **所属 AX**: AX-004（查询幂等确定）
- **测试文件**: test_l1_result_structure.py → `TestINV4RowCount`
- **执行机制**: 对比 SELECT * 行数与 COUNT(*) 值

---

## L2: 架构原则不变量（INV-5 ~ INV-9）

测试文件: `tests/invariants/test_l2_architecture.py`

### INV-5: 失败安全
- **声明**: success=False 时 data=[], message 非空且无堆栈
- **所属 AX**: AX-003（失败安全可诊断）
- **测试文件**: test_l2_architecture.py → `TestINV5FailureSafe`
- **执行机制**: pytest 断言失败结果的 data 为空、message 有内容

### INV-6: 错误可分类
- **声明**: 所有错误消息能被 ToolCallTracker.classify_error() 归入已知类别
- **所属 AX**: AX-003（失败安全可诊断）
- **测试文件**: test_l2_architecture.py → `TestINV6ErrorClassifiable`
- **执行机制**: 遍历已知错误场景，验证 classify_error 返回已知类别

### INV-7: 幂等读取
- **声明**: 同一 SELECT 连续执行两次，结果完全一致
- **所属 AX**: AX-004（查询幂等确定）
- **测试文件**: test_l2_architecture.py → `TestINV7IdempotentRead`
- **执行机制**: 同一 SQL 执行两次，assert data 相等

### INV-8: LIMIT 约束
- **声明**: SELECT ... LIMIT N 返回行数 ≤ N
- **所属 AX**: AX-004（查询幂等确定）
- **测试文件**: test_l2_architecture.py → `TestINV8LimitConstraint`
- **执行机制**: 断言返回行数不超过 LIMIT 值

### INV-9: 聚合语义正确
- **声明**: COUNT(*) ≥ COUNT(col)；SUM 忽略 NULL；空表聚合语义
- **所属 AX**: AX-004（查询幂等确定）
- **测试文件**: test_l2_architecture.py → `TestINV9AggregateSemantics`
- **执行机制**: 构建 NULL 场景数据，验证聚合行为与 SQLite 一致

---

## L3: 具体不变量（INV-10 ~ INV-15）

测试文件: `tests/invariants/test_l3_specific.py`

### INV-10: 窗口函数唯一性
- **声明**: ROW_NUMBER() 在同一 PARTITION 内严格递增且无重复
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV10WindowFunctionUniqueness`
- **执行机制**: 构建分区数据，验证 ROW_NUMBER 严格递增无间断

### INV-11: 排名标准合规
- **声明**: RANK() 并列时跳号，DENSE_RANK() 不跳号
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV11RankingStandard`
- **执行机制**: 构建并列值场景，对比 RANK/DENSE_RANK 行为

### INV-12: 空表安全
- **声明**: 空表上任意 SELECT 返回空数据行但不报错
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV12EmptyTableSafe`
- **执行机制**: 对空表执行多种 SELECT，验证 success=True 且 data 仅含表头

### INV-13: 特殊字符安全
- **声明**: 列名/值含中文、emoji、单引号、反斜杠时不崩溃
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV13SpecialCharSafe`
- **执行机制**: 创建含特殊字符的 sheet，执行 CRUD 验证不崩溃

### INV-14: 除零安全
- **声明**: 1/0 返回 NULL 而非 inf 或崩溃
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV14DivisionByZeroSafe`
- **执行机制**: 执行 SELECT 1/0，验证结果为 NULL

### INV-15: LIKE 安全
- **声明**: LIKE 模式含正则元字符时不崩溃；超长模式被拒绝
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_specific.py → `TestINV15LikeSafe`
- **执行机制**: 用含 `%_` 等元字符的 LIKE 模式查询，验证不崩溃

---

## L3: 写操作不变量（INV-16 ~ INV-24）

测试文件: `tests/invariants/test_l3_write_operations.py`

### INV-16: UPDATE 后读回验证
- **声明**: UPDATE 后 SELECT 读回验证 SET 表达式生效
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV16UpdateReadback`
- **执行机制**: UPDATE → SELECT 读回，断言值匹配

### INV-17: INSERT 行数守恒
- **声明**: INSERT N 行后 COUNT(*) 增加 N
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV17InsertRowCount`
- **执行机制**: 记录前 COUNT → INSERT N → 记录后 COUNT，差值 = N

### INV-18: DELETE 行数守恒
- **声明**: DELETE 后 COUNT(*) 减少 affected_rows
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV18DeleteRowCount`
- **执行机制**: 记录前 COUNT → DELETE → 记录后 COUNT，差值 = affected_rows

### INV-19: 写操作 SQLite 对齐
- **声明**: UPDATE/INSERT/DELETE 后 ExcelMCP 和 SQLite 数据对齐
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_advanced.py → `TestINV19WriteSQLiteAlignment`
- **执行机制**: calibrator 同步后对比 ExcelMCP 与 SQLite 查询结果

### INV-20: 公式列守恒
- **声明**: UPDATE 后非目标列公式仍存在
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV20FormulaPreservation`
- **执行机制**: openpyxl 读回 sheet，检查非目标列 cell.formula 保留

### INV-21: 跨文件 JOIN 真值
- **声明**: 跨文件 JOIN 结果与 SQLite 真值对齐
- **所属 AX**: AX-004（查询幂等确定）
- **测试文件**: test_l3_advanced.py → `TestINV21CrossFileJoin`
- **执行机制**: 两个文件导入 calibrator → SQLite，对比 JOIN 结果

### INV-22: affected_rows 精确
- **声明**: affected_rows == 实际变更行数
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV22AffectedRowsAccuracy`
- **执行机制**: UPDATE N 行，断言 affected_rows == N；无匹配时 == 0

### INV-23: 无匹配写操作安全
- **声明**: UPDATE/DELETE WHERE 无匹配 → 文件不变
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV23NoMatchWriteSafety`
- **执行机制**: 执行无匹配的 UPDATE/DELETE，计算文件哈希不变

### INV-24: NULL 写入语义
- **声明**: 数值列写 NULL 后读回为 NULL/空
- **所属 AX**: AX-002（写操作安全）
- **测试文件**: test_l3_write_operations.py → `TestINV24NullWriteSemantics`
- **执行机制**: UPDATE SET col=NULL → SELECT 读回，验证 NULL 语义

---

## L3: SQL 功能边界不变量（INV-25 ~ INV-32）

测试文件: `tests/invariants/test_l3_sql_features.py`

### INV-25: DISTINCT 语义正确性
- **声明**: DISTINCT 返回去重后的唯一行
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV25Distinct`
- **执行机制**: 构建重复数据，验证 DISTINCT 结果唯一且行数正确

### INV-26: HAVING 子句正确性
- **声明**: HAVING 子句在 GROUP BY 后正确过滤
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV26Having`
- **执行机制**: GROUP BY + HAVING 条件，验证过滤结果

### INV-27: IS NULL / IS NOT NULL 正确性
- **声明**: IS NULL / IS NOT NULL 正确识别空值
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV27NullComparison`
- **执行机制**: 构建 NULL/非 NULL 数据，验证 IS NULL 语义

### INV-28: 子查询正确性
- **声明**: 子查询（IN (SELECT...)）正确执行
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV28Subquery`
- **执行机制**: IN (SELECT ...) 子查询，验证结果与 SQLite 对齐

### INV-29: OFFSET 边界正确性
- **声明**: OFFSET 边界条件
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV29OffsetBoundary`
- **执行机制**: OFFSET 超过行数时返回空结果；OFFSET 0 等同无 OFFSET

### INV-30: NOT IN / NOT LIKE 语义正确性
- **声明**: NOT IN / NOT LIKE 语义正确
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV30NotOperators`
- **执行机制**: NOT IN / NOT LIKE 查询，验证与 SQLite 行为一致

### INV-31: 双行表头写操作正确性
- **声明**: 双行表头表的写操作正确性
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV31DualHeaderWrite`
- **执行机制**: 双行表头 sheet 执行 UPDATE，验证值正确更新

### INV-32: _ROW_NUMBER_ 写操作正确性
- **声明**: _ROW_NUMBER_ 在 UPDATE WHERE 中正确工作
- **所属 AX**: AX-001（SQL 标准真值）
- **测试文件**: test_l3_sql_features.py → `TestINV32RowNumberWrite`
- **执行机制**: UPDATE ... WHERE _ROW_NUMBER_ = N，验证仅目标行变更
