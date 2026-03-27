# 第133轮 - REQ-029 修复2个P0阻断性bug ✅

---

## 状态
版本：v1.6.16 | 工具：44 | 测试：1154

## 本轮完成
- **REQ-029 修复2个P0阻断性bug**：
  - Bug 1：JOIN后SQL表别名`r.名称`不生效，返回列名丢失别名前缀
  - Bug 2：streaming写入后openpyxl read_only模式max_row=None，describe_table崩溃

### Bug 1修复详情
- **根因**：`_extract_select_alias`在解析`SELECT r.名称`时，只返回列名`名称`，丢失了表别名前缀`r.`
- **修复**：当Column表达式有table_part时，保留`table.col`格式作为默认别名
- **影响**：JOIN查询 `SELECT r.名称, d.物品名称` 现在正确返回带前缀的列名

### Bug 2修复详情
- **根因**：streaming写入后openpyxl read_only模式`ws.max_row`返回None，`total_rows`可能未定义
- **修复**：增加try/except包裹`total_rows`引用，失败时降级到iter_rows统计
- **影响**：空文件或异常文件不再崩溃，返回保守值0

## 自我进化评估
- 📊 测试通过率：1154/1154 (100%)
- 📊 代码改动：2文件，+28行/-2行
- 📊 发布：v1.6.16 ✅
- 📊 新bug：0
- 📊 修复质量：精准定位根因，最小改动修复

## 下轮待办
- [ ] REQ-006 工程治理（持续迭代）
- [ ] REQ-010 文档与门面优化
- [ ] MCP真实验证（至少8项游戏场景）
- [ ] README中英文同步检查
