#!/bin/bash
# ExcelMCP 修复验证测试执行脚本

echo "========================================================================"
echo "ExcelMCP SQL 功能修复验证测试"
echo "========================================================================"
echo ""

cd /root/.openclaw/workspace/excel-mcp-server

echo "1. 运行快速验证测试..."
echo "------------------------------------------------------------------------"
python3 quick_test.py
QUICK_TEST_EXIT=$?

echo ""
echo "2. 运行同文件 JOIN 详细测试..."
echo "------------------------------------------------------------------------"
python3 test_same_file_join.py
JOIN_TEST_EXIT=$?

echo ""
echo "3. 运行 GROUP_CONCAT 复杂表达式详细测试..."
echo "------------------------------------------------------------------------"
python3 test_group_concat_complex.py
GROUPCONCAT_TEST_EXIT=$?

echo ""
echo "========================================================================"
echo "测试结果总结"
echo "========================================================================"
echo "快速验证测试: $([ $QUICK_TEST_EXIT -eq 0 ] && echo '✅ 通过' || echo '❌ 失败')"
echo "同文件 JOIN 测试: $([ $JOIN_TEST_EXIT -eq 0 ] && echo '✅ 通过' || echo '❌ 失败')"
echo "GROUP_CONCAT 复杂表达式测试: $([ $GROUPCONCAT_TEST_EXIT -eq 0 ] && echo '✅ 通过' || echo '❌ 失败')"
echo ""

if [ $QUICK_TEST_EXIT -eq 0 ] && [ $JOIN_TEST_EXIT -eq 0 ] && [ $GROUPCONCAT_TEST_EXIT -eq 0 ]; then
    echo "✅ 所有测试通过！修复成功！"
    exit 0
else
    echo "❌ 部分测试失败，请检查详细输出"
    exit 1
fi
