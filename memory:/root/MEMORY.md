# Harness Maturity Model: excel-mcp-server

## Current Maturity: L4
- **Definition**: Full permissions deny/ask/allow, PreToolUse guard (JSON stdin + regex), PostToolUse audit (JSONL), Stop lint hook, Enforcement section in CLAUDE.md, 4 tri-directional Rules, IDD Skill, git hooks.
- **Achieved**: 2026-06-06

## Evolution History
- L2.5 → L3: CLAUDE.md 584→85 lines, 4 Rules created, IDD Skill created, settings.json PreToolUse guard
- L3 → L4: Full permissions model, 3 hook scripts, Enforcement section, pre-commit syntax fix

## L4 Architecture
### Permissions (.claude/settings.json)
- **deny**: destructive commands, credential reads, git force ops
- **ask**: git push, pip install, npm publish
- **allow**: daily dev tools, ruff, pytest, git add/commit, MCP excel tools

### Hooks (~/.claude/hooks/)
- `preToolUse-bash-guard.sh`: JSON stdin → jq parse → regex block → exit 2
- `postToolUse-audit.sh`: JSONL append to ~/.claude/audit/YYYY-MM-DD.jsonl
- `stop-lint.sh`: ruff check on git-modified .py files after each turn

### Rules (.claude/rules/)
1. `sql-standards` — SQLite 3.x alignment
2. `file-safety` — backup/overwrite guards
3. `api-contract` — response schema
4. `testing-standards` — invariant test requirements

### Skills (.claude/skills/)
1. `INVARIANT-DRIVEN-DEV.md` — 32 INV, L1→L2→L3 derivation tree, adversary table

## Invariant System
- **L1 (Axioms)**: INV-1~4 result structure, SQL alignment, file integrity, row conservation
- **L2 (Architecture)**: INV-5~9 fail-safe, error classification, idempotent reads, LIMIT, aggregation
- **L3 (Concrete)**: INV-10~32 window functions, empty tables, write ops, SQL boundaries
- **Tests**: 154 passed, 3 skipped, convergence rating C

## Known Debt
- `server.py` 2784 lines — should split by functional domain
- `advanced_sql_query.py` 10545 lines — SQL engine, complex but cohesive
- Pandas FutureWarning: 4 fixes applied in previous session

## GAN Adversarial Framework
1. **Document Quality**: Gen vs Disc. Target: error[] empty.
2. **Search Quality**: Gen vs Disc. Target: total >90.
3. **Code Correctness**: Gen vs Disc. Target: zero errors.
4. **Knowledge Honesty**: Gen vs Disc. Target: no fabrication.
