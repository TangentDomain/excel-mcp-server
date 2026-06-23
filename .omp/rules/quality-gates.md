# Quality Gates

> 机器可判定的质量规则。每条规则定义判定命令和执行时机。
> 本文件由 Harness L5 进化引擎维护，修改请遵循 L5 提案流程。

---

## ruff-format

- description: Ruff 格式检查，确保代码风格一致
- alwaysApply: true
- globs: ["src/**/*.py", "tests/**/*.py"]
- version: "1.0.0"

判定:

```bash
ruff format --check src/ tests/
```

执行: pre-commit gate (block)

---

## ruff-check

- description: Ruff lint 检查，捕获代码质量问题
- alwaysApply: true
- globs: ["src/**/*.py", "tests/**/*.py"]
- version: "1.0.0"

判定:

```bash
ruff check src/ tests/
```

执行: pre-commit gate (block)

---

## invariant-tests

- description: 不变量测试套件（INV-1~32），154 个测试用例
- alwaysApply: true
- globs: ["tests/invariants/**/*.py"]
- version: "1.0.0"

判定:

```bash
uv run python -m pytest tests/invariants/ -q --timeout=30
```

执行: pre-commit gate (block)

---

## docstring-contract

- description: Docstring 契约验证，确保工具描述与签名一致
- alwaysApply: true
- globs: ["src/**/*.py"]
- version: "1.0.0"

判定:

```bash
python3 scripts/lint_docstring_contract.py --quiet
```

执行: pre-commit gate (block)

---

## no-conflict-markers

- description: 检测 staged diff 中未解决的合并冲突标记
- alwaysApply: true
- globs: ["**/*.py", "**/*.md", "**/*.yml", "**/*.yaml", "**/*.json", "**/*.toml"]
- version: "1.0.0"

判定: `git diff --cached` 中无 `<<<<<<<` 行

执行: pre-commit gate (block)

---

## no-secrets

- description: 检测 staged diff 中可能的密钥/凭证泄露
- alwaysApply: true
- globs: ["**/*.py", "**/*.md", "**/*.yml", "**/*.yaml", "**/*.json", "**/*.toml"]
- version: "1.0.0"

判定: `git diff --cached` 中无 `AKIA`、`LTAI`、`ft522` 等密钥模式

执行: pre-commit gate (block)
