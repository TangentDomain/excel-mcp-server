# 📖 Complete Tool List

ExcelMCP provides 53 specialized tools for Excel configuration table management. All tools are optimized for game development workflows with Chinese language support.

## Excel File Operations

### Basic File Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_open_workbook` | Open Excel workbook | `path`, `mode` | Open skills/equipment configuration |
| `excel_close_workbook` | Close workbook | `path`, `force_cleanup` | Clean up after operations |
| `excel_list_sheets` | List all sheets | `path` | Check available tables |
| `excel_create_workbook` | Create new workbook | `path` | New game project setup |
| `excel_copy_workbook` | Copy workbook | `source`, `destination` | Backup configuration files |
| `excel_sheet_exists` | Check if sheet exists | `path`, `sheet_name` | Verify table existence |

### Sheet Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_create_sheet` | Create new sheet | `path`, `sheet_name` | Add new game data tables |
| `excel_delete_sheet` | Delete sheet | `path`, `sheet_name` | Remove unused tables |
| `excel_rename_sheet` | Rename sheet | `path`, `old_name`, `new_name` | Rename tables for clarity |
| `excel_copy_sheet` | Copy sheet | `path`, `source`, `destination` | Duplicate game configurations |
| `excel_move_sheet` | Move sheet | `path`, `sheet_name`, `position` | Reorganize table order |
| `excel_get_headers` | Get column headers | `path`, `sheet`, `row` | Get table structure |

## Data Reading Operations

### Range Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_get_range` | Get cell range data | `path`, `sheet`, `range` | Read skill/equipment data |
| `excel_get_last_row` | Find last row with data | `path`, `sheet`, `column` | Find data table boundaries |
| `excel_get_last_column` | Find last column with data | `path`, `sheet`, `row` | Determine table width |
| `excel_read_data` | Read structured data | `path`, `sheet`, `start_cell`, `end_cell` | Read entire tables |
| `excel_find_empty_rows` | Find empty rows | `path`, `sheet`, `start_row`, `criteria` | Find gaps in data |

### Search and Query
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_find_value` | Find specific value | `path`, `sheet`, `value`, `range` | Search for specific skills/items |
| `excel_search_columns` | Search in columns | `path`, `sheet`, `column`, `value` | Find skills by type |
| `excel_search_rows` | Search in rows | `path`, `sheet`, `row`, `value` | Find equipment by ID |
| `excel_query_range` | Query with conditions | `path`, `sheet`, `range`, `conditions` | Filter game data |

## Data Writing Operations

### Cell Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_write_cell` | Write single cell | `path`, `sheet`, `cell`, `value` | Update skill damage values |
| `excel_write_cells` | Write multiple cells | `path`, `sheet`, `data_dict` | Batch update game stats |
| `excel_set_formula` | Set cell formula | `path`, `sheet`, `cell`, `formula` | Calculate derived stats |
| `excel_clear_cell` | Clear cell | `path`, `sheet`, `cell` | Reset game values |
| `excel_clear_range` | Clear range | `path`, `sheet`, `range` | Clear table sections |

### Row Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_write_row` | Write entire row | `path`, `sheet`, `row`, `data` | Add new game data |
| `excel_insert_row` | Insert row | `path`, `sheet`, `row`, `data`, `shift` | Add new skills/items |
| `excel_update_row` | Update row | `path`, `sheet`, `row`, `data` | Modify existing data |
| `excel_delete_row` | Delete row | `path`, `sheet`, `row`, `shift` | Remove outdated skills |
| `excel_duplicate_row` | Duplicate row | `path`, `sheet`, `source`, `target` | Copy game configurations |
| `excel_batch_insert_rows` | Batch insert rows | `path`, `sheet`, `start_row`, `data` | Bulk add game data |
| `excel_batch_update_rows` | Batch update rows | `path`, `sheet`, `conditions`, `updates` | Bulk update game stats |

### Column Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_write_column` | Write entire column | `path`, `sheet`, `column`, `data` | Update skill categories |
| `excel_insert_column` | Insert column | `path`, `sheet`, `column`, `data`, `shift` | Add new stat columns |
| `excel_update_column` | Update column | `path`, `sheet`, `column`, `data` | Modify game attributes |
| `excel_delete_column` | Delete column | `path`, `sheet`, `column`, `shift` | Remove unused stats |
| `excel_rename_column` | Rename column | `path`, `sheet`, `old_name`, `new_name` | Rename game attributes |

## Advanced Operations

### SQL Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_query_sql` | Execute SQL query | `path`, `query`, `use_cache` | Complex game data analysis |
| `excel_join_tables` | Join two tables | `path1`, `sheet1`, `path2`, `sheet2`, `join_type` | Combine skills and equipment |
| `excel_group_by` | Group and aggregate | `path`, `sheet`, `group_columns`, `agg_columns` | Analyze game statistics |
| `excel_where_query` | Query with WHERE | `path`, `sheet`, `conditions` | Filter game data by criteria |
| `excel_order_by` | Sort results | `path`, `sheet`, `columns`, `direction` | Sort skills by damage |

### Comparison Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_compare_sheets` | Compare two sheets | `path1`, `sheet1`, `path2`, `sheet2` | Compare game versions |
| `excel_compare_workbooks` | Compare workbooks | `path1`, `path2` | Compare entire game projects |
| `excel_find_differences` | Find differences | `path1`, `sheet1`, `path2`, `sheet2` | Detect game configuration changes |

### Import/Export Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_import_csv` | Import from CSV | `path`, `csv_path`, `sheet`, `start_cell` | Import game data from external sources |
| `excel_export_csv` | Export to CSV | `path`, `sheet`, `csv_path` | Export game data for analysis |
| `excel_import_json` | Import from JSON | `path`, `json_data`, `sheet`, `start_cell` | Import game configurations |
| `excel_export_json` | Export to JSON | `path`, `sheet`, `json_path` | Export game data for APIs |

### Analysis Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_analyze_data` | Analyze data | `path`, `sheet`, `analysis_type` | Game statistics analysis |
| `excel_describe_table` | Get table statistics | `path`, `sheet` | Understand game data structure |
| `excel_calculate_stats` | Calculate statistics | `path`, `sheet`, `columns` | Game performance metrics |

### Specialized Game Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_calculate_damage` | Calculate damage | `path`, `sheet`, `formula` | Game combat calculations |
| `excel_balance_check` | Check game balance | `path`, `sheet`, `balance_type` | Game balance analysis |
| `excel_create_template` | Create game template | `template_type`, `output_path` | New game project setup |
| `excel_progression_analysis` | Analyze game progression | `path`, `sheet`, `progression_type` | Player progression analysis |

## Configuration Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_set_style` | Set cell formatting | `path`, `sheet`, `range`, `style` | Format game tables |
| `excel_freeze_panes` | Freeze panes | `path`, `sheet`, `cell` | Fix headers for game data |
| `excel_protect_sheet` | Protect sheet | `path`, `sheet`, `password` | Protect game configurations |
| `excel_validate_data` | Set data validation | `path`, `sheet`, `range`, `validation` | Game data validation |

## Performance Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_write_only_workbook` | Create streaming workbook | `path` | Large game data files |
| `excel_write_only_override` | Streaming data update | `path`, `sheet`, `data` | Large game data updates |
| `excel_clear_cache` | Clear query cache | `cache_type` | Clear game data cache |
| `excel_get_performance_stats` | Get performance metrics | | Monitor game data performance |

## Utility Operations
| Tool | Description | Parameters | Game Use Case |
|------|-------------|------------|---------------|
| `excel_get_version` | Get ExcelMCP version | | Check tool version |
| `excel_get_help` | Get tool help | `tool_name` | Learn specific tools |
| `excel_validate_file` | Validate Excel file | `path` | Check game data files |
| `excel_convert_format` | Convert file format | `path`, `target_format` | Convert game data formats |

## Tool Categories by Use Case

### 🎮 Game Development Tools (15 tools)
- Skills management, equipment management, monster data, balance analysis
- Progression analysis, template creation, combat calculations

### 📊 Data Analysis Tools (12 tools)
- Statistical analysis, data comparison, SQL queries, data validation
- Performance monitoring, cache management

### 🔧 Configuration Tools (10 tools)
- File operations, sheet management, data import/export
- Style management, protection, validation

### ⚡ Performance Tools (8 tools)
- Streaming operations, batch processing, memory management
- Performance optimization, caching

### 🛠️ Utility Tools (8 tools)
- Version checking, help system, file validation, format conversion
- System tools and utilities

## Tool Usage Examples

### Game Development Example
```
# Open skills configuration
excel_open_workbook("skills.xlsx")

# Find all fire skills
fire_skills = excel_query_sql("skills.xlsx", "SELECT * FROM skills WHERE 技能类型 = '火系'")

# Update fire skill damage
excel_batch_update_rows("skills.xlsx", "技能类型 = '火系'", {"伤害": "伤害 * 1.2"})

# Save changes
excel_close_workbook("skills.xlsx")
```

### Data Analysis Example
```
# Analyze equipment statistics
equipment_stats = excel_analyze_data("equipment.xlsx", "equipment", "descriptive")

# Compare two versions
differences = excel_compare_sheets("equipment_v1.xlsx", "equipment", 
                                 "equipment_v2.xlsx", "equipment")

# Export results for review
excel_export_json("equipment_stats.xlsx", "statistics", "equipment_analysis.json")
```

### Performance Example
```
# Process large game data file
excel_write_only_workbook("large_game_data.xlsx")

# Stream data in batches
for batch in game_data_batches:
    excel_write_only_override("large_game_data.xlsx", "skills", batch)

# Close with cleanup
excel_close_workbook("large_game_data.xlsx", force_cleanup=True)
```