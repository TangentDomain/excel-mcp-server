{
  "REQUIREMENTS": {
    "REQ-035": {
      "title": "配置化：硬编码常量提取为配置项",
      "type": "refactor",
      "priority": "P2",
      "status": "OPEN",
      "source": "自审",
      "attempts": 1,
      "last_failure": "第275轮误标DONE：缓存映射修复不是根因。实际bug在数据加载阶段，original_rows=379但MapEvent sheet只有59行，所有sheet数据被混在一起。详见FEEDBACK.md OPEN-#1",
      "description": "多个硬编码值应提取为可配置常量：max_files=100, query_cache_ttl=300, target_mb=512.0, MAX_RESULT_ROWS=500等。",
      "notes": "分布在server.py和advanced_sql_query.py中"
    },

    "REQ-036": {
      "title": "边缘案例自动化测试：每轮自动搜索并验证奇怪场景",
      "type": "feature",
      "priority": "P1",
      "status": "IN-PROGRESS",
      "source": "CEO",
      "attempts": 16,
      "last_failure": "",
      "description": "每轮执行时，自动搜索一些稀奇古怪的Excel使用场景（如超长公式、特殊字符sheet名、合并单元格+筛选、数据透视表嵌套、条件格式+VBA、超大文件性能等），用uvx安装的MCP工具实际调用测试，记录是否崩溃/返回错误/正常处理。",
      "acceptance_criteria": [
        "每轮至少测试1个新边缘案例",
        "测试结果记录到docs/EDGE-CASE-TESTS.md",
        "格式：日期、案例描述、操作步骤、预期结果、实际结果、是否通过",
        "崩溃或错误自动创建REQ",
        "优先从Stack Overflow/GitHub Issues搜索真实用户遇到的奇怪问题"
      ],
      "notes": "第243轮10案例6通过4失败(REQ-038/039/040)；第245轮10案例9通过1失败(REQ-041)；第247轮5案例全通过(REQ-042修复)；第248轮16案例15通过1信息；第249轮16案例13通过3信息(含server.py修复)；第250轮15案例12通过2信息1失败(||拼接不支持)；第252轮33案例33全通过(核心API稳定性验证)；第253轮25案例11通过11信息3失败(streaming写入不可见)；第254轮20案例20全通过(含check_duplicate_ids列名查找bug修复+发布v1.7.6)；第255轮30案例25通过3信息2失败(T132/T133 ROUND/ABS不支持+T141嵌套聚合计算列丢失)；第256轮20案例19通过1失败(T168 evaluate_formula独立数学表达式不支持)；第257轮20案例(T211-T230)20全通过(REQ-044/045/046验证+batch_insert_rows_at CellInfo bug修复+SQL子查询+空表边界)；第259轮20案例(T256-T275)17通过3信息0失败；第260轮20案例(T276-T295)12通过1失败5信息2依赖(convert_format CSV→xlsx不支持)；第261轮20案例(T296-T315)19通过1信息0失败(T315 Sheet自身对比NoneType bug已修复+发布v1.7.10)；第262轮20案例(T316-T335)20全通过(format_cells number_format bug修复+发布v1.7.11)；第263轮21案例(T336-T355)0失败19信息2通过(测试脚本参数名不匹配，无新BUG)；第264轮20案例(T356-T375)17通过3信息0失败(DataValidationError 3参数调用BUG+ExcelWriter缺失导入BUG，已修复+发布v1.7.12)；第265轮20案例(T376-T395)14通过6信息0失败(excel_describe_table缺失@mcp.tool装饰器BUG，已修复+发布v1.7.13)；第266轮20案例(T396-T415)15通过4失败1错误0实际BUG(excel_list_charts AttributeError修复+clear_validation范围匹配BUG修复+错误码SHEET_NOT_FOUND修正+发布v1.7.14)；第267轮20案例(T416-T435)MCP冒烟通过，边缘测试脚本执行较慢但无新BUG发现；第268轮测试脚本函数名不匹配(excel_create_workbook→excel_create_file)，MCP冒烟通过"
    }
  }
}
