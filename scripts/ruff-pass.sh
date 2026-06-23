#!/bin/bash
# scripts/ruff-pass.sh — Loop I1 判据: ruff format + ruff check
# exit 0 = pass, exit 2 = fail
set -e

result=$(ruff format --check src/ tests/ 2>&1 && ruff check src/ tests/ 2>&1)
if [ $? -eq 0 ]; then
  echo '{"decision":"pass"}'
  exit 0
else
  echo "{\"decision\":\"fail\",\"reason\":\"ruff 失败: $(echo "$result" | tail -3)\"}"
  exit 2
fi
