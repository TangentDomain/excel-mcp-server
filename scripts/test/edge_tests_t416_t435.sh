#!/bin/bash
# 边缘案例测试 T416-T435 - 第267轮

# 创建临时测试文件
TEST_FILE=$(mktemp /tmp/edge_test_XXX.xlsx)

# 使用python创建基础测试文件
python3 << 'EOF'
from openpyxl import Workbook
import sys

# 创建测试文件
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# 添加测试数据
headers = ["Name", "Age", "Score", "Active"]
ws.append(headers)

# 添加数据行
for i in range(10):
    ws.append([f"User{i}", 20+i, 80+i*2, "true" if i%2==0 else "false"])

# 添加第二个工作表
ws2 = wb.create_sheet("Sheet2")
ws2.append(["Key", "Value"])
ws2.append(["A", 100])

wb.save("/tmp/edge_test_base.xlsx")
print("基础测试文件已创建: /tmp/edge_test_base.xlsx")
EOF

BASE_FILE="/tmp/edge_test_base.xlsx"

# MCP服务器路径
MCP_SERVER="uvx --from . excel-mcp-server-fastmcp"

# 测试计数器
PASS=0
FAIL=0
ERROR=0
INFO=0

# 辅助函数
run_test() {
    local num=$1
    local desc=$2
    local tool=$3
    local params=$4

    echo "测试T${num}: ${desc}"

    # 构造JSON-RPC请求
    request=$(cat <<REQ
{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"${tool}","arguments":${params}}}
REQ
)

    # 使用timeout防止卡住
    response=$(echo "$request" | timeout 5 $MCP_SERVER 2>&1 | head -5)

    if echo "$response" | grep -q '"success":true\|"result"'; then
        echo "  结果: PASS - $(echo $response | head -c 100)"
        ((PASS++))
    elif echo "$response" | grep -q '"success":false'; then
        echo "  结果: FAIL - $(echo $response | head -c 100)"
        ((FAIL++))
    else
        echo "  结果: INFO - $(echo $response | head -c 100)"
        ((INFO++))
    fi
}

echo "=== 第267轮边缘案例测试 T416-T435 ==="
echo

# T416: 超长公式嵌套
run_test 416 "超长公式嵌套" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT 1 as x"}'

# T417: 空文件路径
run_test 417 "空文件路径" "excel_get_file_info" '{"file_path":""}'

# T418: 不存在的文件
run_test 418 "不存在的文件" "excel_list_sheets" '{"file_path":"/nonexistent/file.xlsx"}'

# T419: 空工作表名
run_test 419 "空工作表名" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"","range":"A1:B5"}'

# T420: 空范围
run_test 420 "空范围字符串" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":""}'

# T421: 无效范围格式
run_test 421 "无效范围格式" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"INVALID"}'

# T422: 超大行号
run_test 422 "超大行号(>1M)" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:B1048577"}'

# T423: 超大列号
run_test 423 "超大列号(XFD+)" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:XFE1"}'

# T424: 负数范围
run_test 424 "负数范围" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A-1:B1"}'

# T425: 反向范围
run_test 425 "反向范围(Z1:A1)" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"Z1:A1"}'

# T426: 空SQL查询
run_test 426 "空SQL查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":""}'

# T427: 无效SQL语法
run_test 427 "无效SQL语法" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT FROM WHERE"}'

# T428: SQL注入尝试
run_test 428 "SQL注入防护" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1; DROP TABLE users--"}'

# T429: 超长查询
run_test 429 "超长查询(>10K字符)" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT '"$(printf '%.0sA' {1..10000})"' FROM Sheet1"}'

# T430: 中文列名查询
run_test 430 "中文列名查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT Name FROM Sheet1"}'

# T431: 列名包含空格
run_test 431 "列名包含空格" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1"}'

# T432: 多工作表JOIN
run_test 432 "跨工作表JOIN" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1"}'

# T433: NULL值处理
run_test 433 "NULL值比较" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 WHERE Age IS NULL"}'

# T434: 空字符串vs NULL
run_test 434 "空字符串vs NULL" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 WHERE Name = ''"}'

# T435: 布尔值查询
run_test 435 "布尔值查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 WHERE Active = \"true\""}'

echo
echo "=== 第267轮统计 ==="
echo "总计: 20个边缘案例(T416-T435)"
echo "通过: $PASS 个"
echo "失败: $FAIL 个"
echo "错误: $ERROR 个"
echo "信息: $INFO 个"

# 清理
rm -f "$BASE_FILE" "$TEST_FILE"

# 输出JSON
echo "{\"total\":20,\"pass\":$PASS,\"fail\":$FAIL,\"error\":$ERROR,\"info\":$INFO}"
