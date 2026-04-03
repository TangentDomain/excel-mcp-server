#!/bin/bash
# 边缘案例测试 T416-T435 - 第267轮（使用已安装的MCP服务器）

# 创建测试文件
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
EOF

BASE_FILE="/tmp/edge_test_base.xlsx"
MCP_SERVER="excel-mcp-server-fastmcp"

PASS=0
FAIL=0
ERROR=0
INFO=0

# 测试函数 - 使用stdin进行通信
test_mcp() {
    local num=$1
    local desc=$2
    local tool=$3
    local params=$4

    echo "测试T${num}: ${desc}"

    # 构造JSON-RPC请求
    local request='{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"'"$tool"'","arguments":'"$params"}}'

    # 使用已安装的MCP服务器
    local response=$(echo "$request" | timeout 10 $MCP_SERVER 2>&1 | head -20)

    # 检查响应
    if echo "$response" | grep -q '"success":true'; then
        echo "  结果: PASS"
        ((PASS++))
    elif echo "$response" | grep -q '"success":false\|"error"'; then
        echo "  结果: FAIL - $(echo $response | grep -o '"message":"[^"]*"' | head -1)"
        ((FAIL++))
    else
        echo "  结果: INFO - $(echo $response | head -c 80)"
        ((INFO++))
    fi
}

echo "=== 第267轮边缘案例测试 T416-T435 ==="
echo

# T416: 基本查询
test_mcp 416 "基本SQL查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 LIMIT 5"}'

# T417: 空文件路径
test_mcp 417 "空文件路径" "excel_get_file_info" '{"file_path":""}'

# T418: 不存在的文件
test_mcp 418 "不存在的文件" "excel_list_sheets" '{"file_path":"/nonexistent/file.xlsx"}'

# T419: 列出工作表
test_mcp 419 "列出工作表" "excel_list_sheets" '{"file_path":"'$BASE_FILE'"}'

# T420: 获取范围
test_mcp 420 "获取数据范围" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:D5"}'

# T421: 获取表头
test_mcp 421 "获取表头" "excel_get_headers" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1"}'

# T422: 获取文件信息
test_mcp 422 "获取文件信息" "excel_get_file_info" '{"file_path":"'$BASE_FILE'"}'

# T423: 查找最后行
test_mcp 423 "查找最后行" "excel_find_last_row" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","column":"A"}'

# T424: 描述表
test_mcp 424 "描述表结构" "excel_describe_table" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1"}'

# T425: 空SQL查询
test_mcp 425 "空SQL查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":""}'

# T426: 无效SQL
test_mcp 426 "无效SQL语法" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"INVALID SQL"}'

# T427: WHERE条件
test_mcp 427 "WHERE条件查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 WHERE Age > 25"}'

# T428: 聚合查询
test_mcp 428 "聚合查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT COUNT(*) as cnt FROM Sheet1"}'

# T429: GROUP BY
test_mcp 429 "GROUP BY查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT Active, COUNT(*) FROM Sheet1 GROUP BY Active"}'

# T430: ORDER BY
test_mcp 430 "ORDER BY查询" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT * FROM Sheet1 ORDER BY Age DESC"}'

# T431: 不存在的列
test_mcp 431 "查询不存在的列" "excel_query_sql" '{"file_path":"'$BASE_FILE'","query":"SELECT NonExistent FROM Sheet1"}'

# T432: 不存在的工作表
test_mcp 432 "不存在的工作表" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"NonExistent","range":"A1:B5"}'

# T433: 无效范围
test_mcp 433 "无效范围格式" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"INVALID"}'

# T434: 空范围
test_mcp 434 "空范围字符串" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":""}'

# T435: 超大范围
test_mcp 435 "超大行号范围" "excel_get_range" '{"file_path":"'$BASE_FILE'","sheet_name":"Sheet1","range":"A1:B1048577"}'

echo
echo "=== 第267轮统计 ==="
echo "总计: 20个边缘案例(T416-T435)"
echo "通过: $PASS 个"
echo "失败: $FAIL 个"
echo "错误: $ERROR 个"
echo "信息: $INFO 个"

rm -f "$BASE_FILE"

# 输出JSON
echo "{\"total\":20,\"pass\":$PASS,\"fail\":$FAIL,\"error\":$ERROR,\"info\":$INFO}"
