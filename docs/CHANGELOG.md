# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## v1.17.0 - 2026-07-10

### Added
- **Skill 接入方式（首推）**：通过全局 skill 安装 CLI，支持 self-update 自动更新和自举
- **项目级 Excel Skill** (`.omp/skills/excel/SKILL.md` v3.2.0)：覆盖全局 skill，文档与代码同步

### Fixed
- **SELECT t.* FROM subquery**：`_apply_select_expressions` 增加 `Column(this=Star)` 分支，支持限定星号
- **WHERE 引用窗口函数别名**：自动重写为子查询包装（`_rewrite_where_window_alias`），透明支持
- **WHERE 引用 SELECT 别名（非窗口）**：物化别名为临时列（`_materialize_select_aliases_for_where`），与 SQLite 对齐
- **run-python 沙箱扩展**：补充 `__build_class__`/`super`/`ZeroDivisionError` 等 18 个安全内置函数；解除 `sys` 级联阻断（`statistics`/`fractions` 可导入）
- **除零返回 NULL**：整数除法 `inf` → `NaN` → `Int64` nullable → 序列化为 `None`（SQL NULL）；`_serialize_value` 补充 `pd.NA` 处理
- **dtype 降级测试同步**：R58 已移除 float32 降级（精度不足），测试同步为 float64

### Changed
- **模块 docstring 更新**：`advanced_sql_query.py` 标注已支持 FROM 子查询、qualified star、WHERE 别名
- **SKILL.md 已知限制更新**：移除已修复的 3 条限制，SQL 功能表新增 qualified star / WHERE 别名 / 沙箱扩展

### Tests
- 全量 1447 passed, 3 skipped, 1 xfailed, 0 failed
- 新增 15 条 L4 限制消除测试 (`test_l4_limit_fixes.py`)
- SQL 差分测试 169 条全通过，19 类别 100%

---

## v1.16.x - 2026-06

### Added
- 窗口函数全家桶：`AVG/SUM/MIN/MAX/COUNT OVER`, `NTH_VALUE`, `GROUP_CONCAT`, `NTILE`, `PERCENT_RANK`, `CUME_DIST`, `LAG`, `LEAD`, `FIRST_VALUE`, `LAST_VALUE`, `ROWS BETWEEN`
- `UPDATE` 支持窗口函数
- `_ROW_NUMBER_` 虚拟列（SELECT 和 UPDATE 均可用）
- 整体性能提升 10.3x

### Fixed
- `IN/NOT IN` 在 UPDATE 中的静默失败
- NULL 三值逻辑对齐 SQL 标准
- 整数除法截断向零（与 SQLite 一致）
- self-join 表别名列引用
- EXISTS 关联子查询内外表区分
- ORDER BY 位置编号 + COALESCE(NULL,...)
- 括号包裹聚合表达式 Paren 解包
- NULLIF WHERE + 字符串函数嵌入数学表达式
- uint 类型减法下溢防护
- COUNT(表达式) 支持

---

## v1.7.x - 2026-04

### Fixed
- `excel_list_charts` AttributeError 修复
- `excel_clear_validation` 清除范围数据验证无效修复
- 错误码修正：3 处 OPERATION_FAILED → SHEET_NOT_FOUND

---

## v1.6.x - 2026-03

### Added
- 44 个基础 Excel 操作工具
- FastMCP 框架集成
- openpyxl 写入支持
- 双行表头智能检测
- SQL 查询引擎（基于 sqlglot AST 解析）
- 跨文件 JOIN 支持
- 流式写入（大文件优化）
- 公式计算缓存
- 安全验证器（路径穿越/符号链接防护）
- SQLite 交叉校准器
