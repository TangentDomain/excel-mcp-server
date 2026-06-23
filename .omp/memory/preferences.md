# 用户偏好

>用户反复表达的行为偏好，无需每次重申。

## 代码风格

- 精简优先：不写多余注释，代码自文档化
- 中文注释：注释和 commit message 用中文，代码标识符用英文
- 类型标注：Python 代码必须有 type hints

## 测试

- INV 优先：不变量测试（tests/invariants/）比常规测试优先级高
- 测试驱动：修 bug 时先写测试再改代码
- 不建 mock：测试行为而非 plumbing

## 质量门禁

- ruff 必过：format + check 全部通过才能提交
- 不变量全绿：INV-1~32 全部通过
- pre-commit gate：提交前自动运行 ruff + pytest

## SQL 标准

- SQL 优先：能用 SQL 解决的问题优先用 SQL 工具而非 Python 脚本
- SQLite 准则：SQL 行为以 SQLite 3.x 为准
- Excel 限制：SQL 与 Excel 冲突时 SQL 优先，Excel 限制记为已知问题

## 输出

- 中文输出：对话输出用中文
- 结构化结果：返回 JSON 结构化结果，不返回自然语言描述
- 错误脱敏：错误信息不泄露堆栈和内部路径
