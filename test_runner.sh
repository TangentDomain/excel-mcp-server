#!/bin/bash
cd /root/.openclaw/workspace/excel-mcp-server

echo "=========================================="
echo "TEST 1: GROUP_CONCAT Tests"
echo "=========================================="
python3 -m pytest tests/test_group_concat.py -v 2>&1 | tail -50
TEST1_EXIT=${PIPESTATUS[0]}

echo ""
echo "=========================================="
echo "TEST 2: Right JOIN Test"
echo "=========================================="
python3 -m pytest tests/test_join_types.py::TestRightJoin::test_basic_right_join -v 2>&1 | tail -30
TEST2_EXIT=${PIPESTATUS[0]}

echo ""
echo "=========================================="
echo "SUMMARY"
echo "=========================================="
if [ $TEST1_EXIT -eq 0 ]; then
    echo "GROUP_CONCAT tests: PASSED ✓"
else
    echo "GROUP_CONCAT tests: FAILED ✗"
fi

if [ $TEST2_EXIT -eq 0 ]; then
    echo "Right JOIN test: PASSED ✓"
else
    echo "Right JOIN test: FAILED ✗"
fi

exit $((TEST1_EXIT + TEST2_EXIT))
