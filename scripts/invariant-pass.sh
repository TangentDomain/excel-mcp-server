#!/bin/bash
# scripts/invariant-pass.sh — Loop I1 判据: pytest invariants
# exit 0 = pass, exit 2 = fail
set -e

output=$(python -m pytest tests/invariants/ -q --tb=short --timeout=30 2>&1)
if [ $? -eq 0 ]; then
  echo '{"decision":"pass"}'
  exit 0
else
  echo "{\"decision\":\"fail\",\"reason\":\"不变量测试失败: $(echo "$output" | tail -3)\"}"
  exit 2
fi
