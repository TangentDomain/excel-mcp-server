# Conventional Commits 规范

本文档定义了 excel-mcp-server 项目的提交信息规范。

## 格式规范

提交信息必须遵循以下格式：

```
[REQ-XXX] type: 简短描述

[可选的详细描述]

[可选的脚注]
```

### 示例

```
[REQ-028] fix: insert_mode 默认值改为 false

修复了 excel_update_range 函数的 insert_mode 参数默认值，
从 true 改为 false，避免覆盖已有数据时意外插入新行。

Closes REQ-028
```

## Type 类型定义

Type 必须是以下之一：

- **feat**: 新功能
- **fix**: 修复 bug  
- **refactor**: 重构（既不是新功能也不是修复 bug）
- **docs**: 文档更新
- **test**: 测试相关
- **chore**: 构建或辅助工具变动
- **perf**: 性能优化

## Scope 作用域

可选的作用域，用于标识影响范围：

```
[REQ-029] feat(api): 添加 docstring 契约验证脚本
```

## 提交规则

1. **必须使用** `[REQ-XXX]` 前缀（针对需求编号）
2. **必须指定** type，且必须是预定义的类型之一
3. **简短描述**必须以动词开头（小写）
4. **详细描述**是可选的，但建议超过50行代码时添加
5. **多个段落**用空行分隔
6. **禁止**在提交信息中包含 `Fixes #123` 或 `Closes #123`，改为使用 `Closes REQ-XXX`

## 违规处理

违反规范的提交信息会被 CI 检查拒绝，必须使用 `git commit --amend` 修正。