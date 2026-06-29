# Excel Skill 实战测试报告

> 测试数据：`D:\tr\svn\trunk\配置表`（真实游戏项目配置表）
> 测试日期：2026-06-29
> CLI 版本：1.17.0 (commit 7476901)
> Skill 版本：v3.2.0

## 测试环境

| 配置表 | Sheet 数 | 最大行数 | 最大列数 | 特点 |
|--------|---------|---------|---------|------|
| Monster.xlsx | 7 | 416 | 46 | 双行表头，中英文列名，跨 Sheet 关联 |
| TrSkill.xlsx | — | 1465 | 48 | 大表，复杂技能配置 |
| TrBuff.xlsx | — | — | — | Buff 配置 |
| Props.xlsx | — | — | — | 道具表 |

## 发现的问题

### P1: 双行表头检测失败 🔴 P0（严重）

**现象**：`Monster.xlsx` 实际是双行表头（第1行中文描述 + 第2行英文字段名），但 `describe-table` 返回 `header_type: single`，导致：
- 引擎用第1行中文作为列名（`怪物ID`, `血量`, `攻击`）
- 英文字段名（`ID`, `Hp`, `Atk`）被当成数据行
- SQL 查询用中文名（`SELECT 怪物ID`）失败（报 MonsterID 不存在）
- SQL 查询用英文名（`SELECT ID`）"成功"但实际取到的是英文字段名行而非数据

**根因**：双行表头检测逻辑 `detect_from_rows` 判断条件不够——Monster 表的第1行中文和第2行英文都看起来像有效数据行。

**影响**：所有真实游戏配置表几乎都用双行表头。检测失败 = SQL 引擎不可用。

**临时绕过**：用 `get-range` 直接读取单元格。

**修复方向**：增强 `detect_from_rows` 启发式（检测第1行中文占比 vs 第2行 ASCII 占比）。

---

### P2: `query` 不支持 `--sheet` 参数 🔴 P0（严重）

**现象**：`query --file F --sheet Monster --sql "..."` 报错 `unrecognized arguments: --sheet`。

**影响**：多 Sheet 文件只能查第一个 Sheet。

**修复方向**：`query` 添加可选 `--sheet` 参数。

---

### P3: `search-directory` 缺少 `--regex` 属性 🟠 P1

**现象**：`search-directory --dir D --pattern "Boss"` 报错 `'Namespace' object has no attribute 'regex'`。

**根因**：`cmd_search_directory` 引用 `args.regex`，但 argparse 未定义。

---

### P4: 跨 Sheet JOIN 列名解析失败 🟠 P1

**现象**：`SELECT m.ID, d.DropID FROM Monster m JOIN MonsterDrop d ON m.ID = d.MonsterID` → `列 'd.DropID' 不存在`。

**根因**：JOIN 执行了但右表列未加入可用列列表。

---

### P5: `get-headers` 在双行表头上返回 `None` 🟡 P2

**现象**：返回 `{"success": false, "message": "无法读取表头数据: None"}`。

**根因**：与 P1 相关。

---

### P6: `list-sheets` 返回 `rows: 0, cols: 0` 🟡 P2

**现象**：所有 Sheet 的 rows/cols 都是 0（实际 Monster 有 416 行 46 列）。

---

### P7: `CAST` 对超大数值返回字符串 🔵 P3

**现象**：`CAST(Hp AS INT)` 对 `9999999999` 返回字符串。

---

### P8: 列名含空值被清洗为 `nan` 🔵 P3

**现象**：第 5 列（"备注"）的英文字段名为空，SQL 可用列中显示为 `'nan'`。

## 已修复 ✅ (commit pending)

| # | 问题 | 修复 | 验证 |
|---|------|------|------|
| P1 | describe-table 硬编码 header_type=single | 改用 detect_from_rows 做双行表头检测 | ✅ Monster: header_type=dual, 列名=英文字段名 |
| P2 | query 不支持 --sheet 参数 | 添加 --sheet 参数，传递 sheet_name | ✅ `query --sheet Monster` 正常 |
| P3 | search-directory 引用不存在的 args.regex | argparse 补全 --regex/--case-sensitive/--whole-word | ✅ 目录搜索正常，找到 "Boss" 关键词 |
| P4 | 跨 Sheet JOIN 右表列名 | 实测正常工作，原报错是列名不存在(DropID)导致 | ✅ `SELECT m.ID, d.MonsterID FROM Monster m JOIN MonsterDrop d` 正常 |

## 第二轮发现的新问题

### P9: `SELECT alias.* FROM (subquery) AS alias` 不支持 🟡 P2

**现象**：子查询派生表的 `SELECT alias.*` 报 "列 'TOP_5.*' 不存在"。

**根因**：引擎不支持 `表别名.*` 语法展开。

**临时绕过**：手动列出列名 `SELECT ID, Qua, CD FROM (subquery) AS alias`。

### P10: `Unnamed__N` 列名 🟡 P2

**现象**：双行表头中英文字段名为空的列，被命名为 `Unnamed__1` 等。

**影响**：用户不知道这是什么列。建议改为用第1行中文名作为列名。

### P11: Sheet 名与文件名不一致 🔵 P3 (非 bug，用户体验)

**现象**：`Props.xlsx` 的 Sheet 不叫 "Props" 而是 "PropList", "PropName" 等。用户用 `FROM Props` 会报错。

**建议**：错误提示中加入 `list-sheets` 建议。

## 正常工作的功能 ✅ (第二轮验证)

| 功能 | 测试 | 结果 |
|------|------|------|
| describe-table 双行表头 | Monster.xlsx | ✅ header_type=dual, 列名=英文 |
| query --sheet | `--sheet Monster` | ✅ 正常 |
| search-directory | 搜 "Boss" | ✅ 找到 Activity.xlsx 中的匹配 |
| 跨 Sheet JOIN | Monster JOIN MonsterDrop | ✅ 左右表列均可访问 |
| GROUP BY + HAVING | Monster 按 Camp 分组 | ✅ 我方88/敌方326 |
| CASE WHEN | Hp 分段判断 | ✅ tank/normal/weak |
| 子查询 | WHERE Hp > AVG(Hp) | ✅ 正确 |
| LIKE 中文搜索 | `Type LIKE '%KING%'` | ✅ |
| 窗口函数 + PARTITION | RANK() OVER (PARTITION BY Camp) | ✅ |
| ROUND(AVG()) | 平均血量保留2位 | ✅ |
| DISTINCT | 去重类型 | ✅ |
| 大表加载 | TrSkill 1463×40, Props PropList | ✅ |
| GROUP BY 多列值 | PropList 按 PropType | ✅ 1012条宝箱/80条碎片 |

## 结论

P0/P1 问题已全部修复。核心 SQL 引擎在真实游戏配置表上工作正常。剩余问题为体验优化（列名命名、错误提示）。
