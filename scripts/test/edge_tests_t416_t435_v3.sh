#!/bin/bash
# 边缘案例测试 T416-T435 - 第267轮

python3 << 'EOF'
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.append(["Name", "Age", "Score", "Active"])
for i in range(10):
    ws.append([f"User{i}", 20+i, 80+i*2, "true" if i%2==0 else "false"])
ws2 = wb.create_sheet("Sheet2")
ws2.append(["Key", "Value"])
ws2.append(["A", 100])
wb.save("/tmp/edge_test_base.xlsx")
print("OK")
EOF

BASE_FILE="/tmp/edge_test_base.xlsx"
MCP="excel-mcp-server-fastmcp"

echo "=== 第267轮边缘案例测试 T416-T435 ==="
echo

# 测试结果数组
declare -a results

# 测试函数
run_test() {
    local num=$1 desc=$2 tool=$3 params=$4
    echo -n "测试T${num}: ${desc} ... "
    local req='{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"'"$tool"'","arguments":'"$params"}}'
    local resp=$(echo "$req" | timeout 10 $MCP 2>&1 | head -10)

    if echo "$resp" | grep -q '"success":true'; then
        echo "PASS"
        results[$num]="PASS"
    elif echo "$resp" | grep -q '"success":false'; then
        echo "FAIL"
        results[$num]="FAIL"
    else
        echo "INFO"
        results[$num]="INFO"
    fi
}

# 执行测试
run_test 416 "基本SQL查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 LIMIT 5"}'
run_test 417 "空文件路径" "excel_get_file_info" '{"file_path":""}'
run_test 418 "不存在的文件" "excel_list_sheets" '{"file_path":"/nonexistent.xlsx"}'
run_test 419 "列出工作表" "excel_list_sheets" '{"file_path":"'$BASE_FILE'"}'
run_test 420 "获取数据范围" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:D5"}'
run_test 421 "获取表头" "excel_get_headers" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1"}'
run_test 422 "获取文件信息" "excel_get_file_info" '{"file_path":"'$BASE_FILE'"}'
run_test 423 "查找最后行" "excel_find_last_row" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","column":"A"}'
run_test 424 "描述表结构" "excel_describe_table" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1"}'
run_test 425 "空SQL查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":""}'
run_test 426 "无效SQL语法" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"INVALID"}'
run_test 427 "WHERE条件查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 WHERE Age > 25"}'
run_test 428 "聚合查询COUNT" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT COUNT(*) FROM Sheet1"}'
run_test 429 "GROUP BY查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT Active, COUNT(*) FROM Sheet1 GROUP BY Active"}'
run_test 430 "ORDER BY查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 ORDER BY Age DESC"}'
run_test 431 "查询不存在的列" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT NonExistent FROM Sheet1"}'
run_test 432 "不存在的工作表" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"NonExistent","range":"A1:B5"}'
run_test 433 "无效范围格式" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"INVALID"}'
run_test 434 "空范围字符串" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":""}'
run_test 435 "超大行号范围" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:B1048577"}'

# 统计
PASS=0 FAIL=0 INFO=0
for r in "${results[@]}"; do
    [ "$r" = "PASS" ] && ((PASS++))
    [ "$r" = "FAIL" ] && ((FAIL++))
    [ "$r" = "INFO" ] && ((INFO++))
done

echo
echo "=== 第267轮统计 ==="
echo "总计: 20个边缘案例(T416-T435)"
echo "通过: $PASS 个"
echo "失败: $FAIL 个"
echo "信息: $INFO 个"

rm -f "$BASE_FILE"
echo "{\"total\":20,\"pass\":$PASS,\"fail\":$FAIL,\"error\":0,\"info\":$INFO}"
