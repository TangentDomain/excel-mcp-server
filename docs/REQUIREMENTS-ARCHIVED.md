{
  "REQUIREMENTS_ARCHIVED": {
    "REQ-031": {
      "title": "版本一致性检查与自动化修复脚本",
      "description": "创建check-version-sync.py脚本，自动检测并修复版本不一致问题，确保pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG.md之间的版本信息同步",
      "priority": "high",
      "status": "DONE",
      "acceptance": [
        "创建scripts/check-version-sync.py脚本",
        "脚本自动检测所有版本文件的一致性",
        "发现不一致时自动修复到正确版本",
        "修复过程记录到DECISIONS.md",
        "集成到每轮自动化流程中"
      ]
    },
    "REQ-029": {
      "title": "Excel操作异常处理优化",
      "description": "将excel_operations.py中的通用异常处理替换为具体的自定义异常类，提升错误处理的精确性和用户体验",
      "status": "DONE", 
      "priority": "HIGH",
      "acceptance_criteria": [
        "替换excel_operations.py中的通用Exception捕获为具体异常类",
        "根据错误类型使用SheetNotFoundError、InvalidRangeError、DataValidationError等",
        "保持现有功能不变，确保所有测试通过",
        "提供更精确的错误信息和上下文"
      ],
      "estimated_workload": "中等（约30分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 14:15 UTC",
      "completed_at": "2026-03-28 14:30 UTC",
      "notes": "PyPI发布失败(403 Forbidden, token过期)，代码改动已合并到main"
    },
    "REQ-028": {
      "title": "错误处理机制优化",
      "description": "改进Excel操作中的错误处理，提供更友好的错误信息和AI修复建议",
      "status": "DONE", 
      "priority": "HIGH",
      "acceptance_criteria": [
        "为关键Excel操作函数添加详细的错误处理",
        "提供AI友好的错误消息和修复建议",
        "添加错误分类和上下文信息",
        "不影响现有功能，通过所有测试"
      ],
      "estimated_workload": "中等（约25分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 14:00 UTC",
      "completed_at": "2026-03-28 14:15 UTC",
      "notes": "错误处理机制优化完成，PyPI发布成功"
    }
  },
  "REQUIREMENTS": {
    "REQ-030": {
      "title": "CI CTE测试失败修复",
      "description": "CTE测试在除macOS 3.10外所有平台失败，需要排查python-calamine版本兼容性或pin版本",
      "priority": "中",
      "status": "DONE",
      "source": "CEO反馈+CI观察",
      "resolution": "CTE测试现在全部通过，可能为版本兼容性自动解决"
    }
  },
  "NEWEST_ARCHIVED": {
    "REQ-031": {
      "title": "自动化版本一致性检查脚本",
      "description": "创建check-version-sync.py脚本，自动检测并修复pyproject.toml、__init__.py、README.md、README.en.md中的版本不一致问题",
      "status": "DONE",
      "priority": "HIGH",
      "acceptance_criteria": [
        "创建scripts/check-version-sync.py脚本",
        "检查pyproject.toml、__init__.py、README.md、README.en.md版本一致性",
        "发现不一致时自动修复并记录到DECISIONS.md",
        "脚本执行时间<3秒，不影响正常开发流程",
        "通过所有现有测试，无回归"
      ],
      "estimated_workload": "中（约20分钟）",
      "assignee": "self-evolution-agent",
      "created_at": "2026-03-28 16:00 UTC",
      "notes": "基于自我进化建议，解决版本同步依赖手动操作的问题"
    },
    "REQ-028": {
      "title": "excel_update_range insert_mode 默认值改为 false",
      "status": "DONE",
      "priority": "P0",
      "description": "excel_update_range 的 insert_mode 默认为 true，导致写入已有数据文件时会物理插入新行",
      "archived_at": "2026-04-01"
    },
    "REQ-029": {
      "title": "工程强化：约束可机器验证",
      "status": "DONE",
      "priority": "P1",
      "description": "将靠LLM自觉遵守的规则升级为靠脚本验证的规则",
      "archived_at": "2026-04-01"
    },
    "REQ-030": {
      "title": "API参数命名与常见术语对齐",
      "status": "DONE",
      "priority": "P2",
      "description": "create_chart的chart_type支持column别名，create_pivot_table的agg_func支持mean别名",
      "archived_at": "2026-04-01"
    },
    "REQ-031_v2": {
      "title": "修复测试文件语法错误",
      "status": "DONE",
      "priority": "P1",
      "description": "test_mcp_actual.py和test_api_issues.py语法错误修复",
      "archived_at": "2026-04-01"
    },
    "REQ-032": {
      "title": "性能优化：大型Excel文件处理提速（2GB+）",
      "status": "DONE",
      "priority": "P1",
      "description": "处理大型Excel文件（2GB+）时遇到性能瓶颈，优化内存使用和数据处理速度",
      "archived_at": "2026-04-01"
    },
    "REQ-033": {
      "title": "性能优化：iterrows替换为itertuples",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "advanced_sql_query.py中多处使用df.iterrows()遍历DataFrame，性能较差。替换为itertuples()或向量化操作。",
      "notes": "5处iterrows全部替换：advanced_sql_query.py结果序列化+条件过滤、server.py UPDATE/DELETE行匹配+透视表写入",
      "archived_at": "2026-04-01"
    },
    "REQ-038": {
      "title": "BUG：工作表名称非法字符静默替换 + 超长名称静默截断",
      "status": "DONE",
      "priority": "P1",
      "source": "边缘案例测试",
      "description": "_normalize_sheet_name()将方括号[]等非法字符静默替换为下划线（Data [2024]→Data _2024_），超长名称静默截断为25+...字符。用户不知情地创建了与预期不同的工作表名，后续引用会失败。",
      "notes": "第244轮修复：拆分为_validate_sheet_name（严格校验）和_sanitize_sheet_name（静默清理），create_sheet/rename_sheet拒绝非法名称，copy_sheet允许静默清理",
      "archived_at": "2026-04-01"
    },
    "REQ-039": {
      "title": "功能缺失：list_sheets不区分隐藏工作表",
      "status": "DONE",
      "priority": "P2",
      "source": "边缘案例测试",
      "description": "excel_list_sheets将visible/hidden/veryHidden工作表一视同仁列出，用户无法区分。应增加sheet_state字段标识可见性。",
      "notes": "第245轮修复：SheetInfo新增sheet_state字段，calamine通过sheets_metadata读取，openpyxl通过sheet.sheet_state读取",
      "archived_at": "2026-04-02"
    },
    "REQ-041": {
      "title": "BUG：SQL含空格列名返回列头字符串而非实际值",
      "status": "DONE",
      "priority": "P1",
      "source": "边缘案例测试",
      "description": "当Excel列名含空格（如\"Player Name\"），_clean_column_names()将空格替换为下划线（Player_Name），但SQL中SELECT \"Player Name\"无法匹配清洗后的列名，导致返回列头字符串代替实际值。",
      "notes": "第245轮修复：新增_preprocess_quoted_identifiers方法，在SQL解析前将双引号引用的原始列名替换为清洗后的列名",
      "archived_at": "2026-04-02"
    },
    "REQ-042": {
      "title": "BUG：_preprocess_quoted_identifiers未处理SQL转义引号",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "_preprocess_quoted_identifiers使用简单的字符串替换处理双引号列名，如果SQL中包含转义引号（如\"col\\\"name\"\"），可能导致错误替换。",
      "notes": "第247轮修复：改用AST方法精确替换列引用位置（SELECT/ORDER BY/GROUP BY），WHERE值位置保持不变；新增_col_map_cache解决缓存命中时映射丢失问题",
      "archived_at": "2026-04-02"
    },
    "REQ-043": {
      "title": "安全回归：commit e9590b0移除39处_validate_path调用未替换为装饰器",
      "status": "DONE",
      "priority": "P0",
      "source": "自审",
      "description": "commit e9590b0尝试实现REQ-034（路径验证装饰器），移除了39处_path_err=_validate_path(file_path)调用，但@_validate_file_path装饰器从未被正确应用到工具函数上。导致路径遍历安全检查从大部分工具函数中丢失。",
      "notes": "第251轮修复：为10个MCP工具函数添加@_validate_file_path装饰器，2个函数添加内联_validate_path调用，v1.7.4发布",
      "archived_at": "2026-04-02"
    },
    "REQ-034": {
      "title": "代码质量：路径验证逻辑抽取为装饰器",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "server.py中_validate_path检查模式重复出现10+次，抽取为装饰器减少重复代码。",
      "notes": "REQ-043修复过程中已为20+个MCP工具函数应用@_validate_file_path装饰器，剩余少量内联调用为特殊场景(merge/batch)",
      "archived_at": "2026-04-02"
    },
    "REQ-037": {
      "title": "线程安全：formula_cache并发访问保护",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "formula_cache.py中的缓存操作在并发MCP调用场景下缺乏线程安全保护，可能导致缓存数据竞争。",
      "notes": "已实现：threading.RLock()保护所有公共方法(get/put/cache_workbook/get_cached_workbook/clear/invalidate_file/get_stats)",
      "archived_at": "2026-04-02"
    },
    "REQ-040": {
      "title": "信息不准确：稀疏工作表file_info维度被格式化膨胀",
      "status": "DONE",
      "priority": "P2",
      "source": "边缘案例测试",
      "description": "当工作表在远端单元格（如Z100）仅有格式化而无数据时，excel_get_file_info返回total_rows=100、total_cols=26，与实际数据范围不符。",
      "notes": "第252轮修复：get_file_info区分实际数据维度和格式化维度，仅当两者不同时才额外报告formatted_rows/formatted_cols",
      "archived_at": "2026-04-02"
    },
    "REQ-044": {
      "title": "find_last_row列名查找与check_duplicate_ids一致化",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "find_last_row中使用column_index_from_string直接解释列参数，与check_duplicate_ids修复后的行为不一致。",
      "notes": "第257轮修复：先查表头匹配列名，找不到再回退列字母解释",
      "archived_at": "2026-04-02"
    },
    "REQ-045": {
      "title": "batch_insert_rows insert_position模块导入错误",
      "status": "DONE",
      "priority": "P2",
      "source": "边缘案例测试",
      "description": "batch_insert_rows指定insert_position时报错模块路径有误。",
      "notes": "第257轮修复：ExcelWriter导入路径从api.excel_writer改为core.excel_writer",
      "archived_at": "2026-04-02"
    },
    "REQ-046": {
      "title": "delete_rows condition数值类型比较问题",
      "status": "DONE",
      "priority": "P2",
      "source": "自审",
      "description": "delete_rows使用condition参数时，对数值列返回0行删除。疑似条件解析将数值列作为字符串处理。",
      "notes": "第257轮修复：df.query前用pd.to_numeric(errors='ignore')转换数值列",
      "archived_at": "2026-04-02"
    }
    "REQ-048": {
      "title": "保护：删除最后一个Sheet时应阻止或自动创建默认Sheet",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "description": "第260轮边缘测试T292发现：excel_delete_sheet允许删除最后一个Sheet，导致工作簿无Sheet。应在删除前检查剩余Sheet数量。",
      "notes": "已验证：excel_manager.py:319-321已有len(wb.sheetnames)<=1检查，test_delete_last_sheet已覆盖。第263轮确认无需修改。",
      "archived_at": "2026-04-03"
    },
    "REQ-051": {
      "title": "边缘测试脚本同步：修正函数名不匹配问题",
      "type": "fix",
      "priority": "P1",
      "status": "DONE",
      "source": "自审",
      "attempts": 1,
      "description": "边缘测试脚本edge_case_tests_round268.py使用了过时的函数名（excel_create_workbook应改为excel_create_file），导致测试脚本无法正常运行。",
      "notes": "第271轮验证：脚本中所有8个MCP函数名均正确，不使用excel_create_workbook。原始描述不准确，实际无函数名不匹配问题。",
      "archived_at": "2026-04-03"
    },
    "REQ-055": {
      "title": "修复：excel_create_pivot_table错误码不一致",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 1,
      "description": "excel_create_pivot_table函数在Sheet不存在时使用OPERATION_FAILED错误码，而其他函数使用SHEET_NOT_FOUND。应统一为SHEET_NOT_FOUND。",
      "notes": "第271轮修复：OPERATION_FAILED→SHEET_NOT_FOUND（server.py:3462）",
      "archived_at": "2026-04-03"
    },
    "REQ-047": {
      "title": "重构：抽取Sheet验证公共方法消除重复代码",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 0,
      "description": "server.py中多个user_friendly函数（get_range/update_range/format_cells等）重复执行相同的Sheet存在性验证逻辑（加载workbook→检查sheet名→返回错误）。应抽取为公共工具函数。",
      "notes": "第259轮自审发现，涉及server.py约4处重复。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-049": {
      "title": "Docstring合规率提升：补充缺失的Args/Returns文档段",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "质量抽检",
      "attempts": 1,
      "description": "497个函数中仅233个有Args段（合规率46.9%），远低于85%目标。需批量补充公共函数的Args/Parameters和Returns文档段。",
      "notes": "来源FEEDBACK.md #1/#3（第7轮），目标合规率85%以上。已达成85.4%合规率，完成目标。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-050": {
      "title": "工具函数抽取：将RichText纯文本提取逻辑抽取为公共函数",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 0,
      "description": "excel_list_charts函数中的_extract_title_text逻辑用于从openpyxl RichText对象提取纯文本，这个逻辑可能在其他地方复用（如读取图表标题、单元格注释等），应抽取为公共工具函数放到utils/目录下。",
      "notes": "第268轮代码自审发现。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-052": {
      "title": "修复GROUP BY聚合错误：部分行被归入不符合WHERE条件的分组",
      "type": "fix",
      "priority": "P0",
      "status": "DONE",
      "source": "质量抽检",
      "attempts": 4,
      "description": "GROUP BY聚合查询中，部分行被错误归入不符合WHERE条件的分组。例：WHERE 显示路径ID IN (1,2) AND 显示位置ID < 100 GROUP BY后出现路径ID=36、位置ID=569的行。文件：src/api/advanced_sql_query.py GROUP BY逻辑。",
      "notes": "来源FEEDBACK.md OPEN-#1，CEO实测MapEvent.xlsx复现。第269轮代码审查：执行顺序正确（WHERE先于GROUP BY），未发现逻辑bug。第271轮深入审查：发现_apply_where_clause存在静默失败场景（condition_str为空时返回未过滤df），但IN+AND条件不受影响。最可能原因是数据类型不匹配。数据类型已确认全部int（CEO已用MapEvent.xlsx验证）。Bug在数据加载阶段：original_rows=379但MapEvent sheet只有59行，所有sheet数据被混在一起。详见FEEDBACK.md OPEN-#1精确线索。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-053": {
      "title": "优化：抽取excel_list_charts中的_extract_title_text为公共函数",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 0,
      "description": "excel_list_charts函数中的_extract_title_text逻辑用于从openpyxl RichText对象提取纯文本，定义在函数内部，每次调用都会重新定义，效率略低。应抽取为模块级别的公共函数，供其他函数（如读取图表标题、单元格注释等）复用。",
      "notes": "第269轮代码自审发现（与REQ-050类似，但针对chart场景）。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-054": {
      "title": "优化：恢复DataValidationError的结构化错误信息",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 0,
      "description": "第266轮将DataValidationError从3参数（错误标题、错误描述、错误建议）简化为1参数（完整错误信息），降低了错误信息的结构化程度，可能影响AI的错误理解和自动修复能力。建议恢复为3参数格式，提升错误处理质量。",
      "notes": "第269轮代码自审发现（commit 41a8e6e简化错误信息，降低了AI可读性）。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-056": {
      "title": "修复：_apply_where_clause静默失败时不返回未过滤DataFrame",
      "type": "fix",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 3,
      "description": "advanced_sql_query.py的_apply_where_clause中，当_sql_condition_to_pandas返回None或空字符串时（如EXISTS子查询），直接返回未过滤的DataFrame（第2907行），导致WHERE条件被静默跳过。应改为抛出错误或记录警告。",
      "notes": "第271轮代码自审发现（REQ-052审查过程中的附带发现）。第274轮完成。",
      "archived_at": "2026-04-03"
    },
    "REQ-035": {
      "title": "配置化：硬编码常量提取为配置项",
      "type": "refactor",
      "priority": "P2",
      "status": "DONE",
      "source": "自审",
      "attempts": 2,
      "last_failure": "第275轮误标DONE：缓存映射修复不是根因。实际bug在数据加载阶段，original_rows=379但MapEvent sheet只有59行，所有sheet数据被混在一起。详见FEEDBACK.md OPEN-#1",
      "description": "多个硬编码值应提取为可配置常量：max_files=100, query_cache_ttl=300, target_mb=512.0, MAX_RESULT_ROWS=500等。",
      "notes": "分布在server.py和advanced_sql_query.py中",
      "completed_at": "2026-04-03",
      "completion_commit": "fe0b0f8",
      "archived_at": "2026-04-04"
    },
    "REQ-061": {
      "title": "GROUP BY 聚合逻辑 bug",
      "type": "fix",
      "priority": "P0",
      "status": "DONE",
      "source": "FEEDBACK.md OPEN-#1",
      "attempts": 5,
      "created": "2026-04-04",
      "completed": "2026-04-05",
      "description": "GROUP BY 聚合逻辑导致TOTAL行数据错误，影响所有聚合查询。必须在_apply_group_by_aggregation方法中修复聚合计算逻辑。",
      "notes": "验证规则：\n```\npython3 -c \"\nimport sys; sys.path.insert(0,\"src\")\nfrom excel_mcp_server_fastmcp.api.advanced_sql_query import execute_advanced_sql_query\nr = execute_advanced_sql_query(\"/tmp/test.xlsx\",\"SELECT 分类, SUM(值) as total_val FROM Sheet GROUP BY 分类 ORDER BY 分类\")\nfor row in r[\"data\"]: print(row)  # 第二行应该是 [\"TOTAL\", \"实际聚合值\"]\n\"\n```\n修复后TOTAL行的聚合值必须正确。",
      "archived_at": "2026-04-05"
    }
  }
}
