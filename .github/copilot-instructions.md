# Excel MCP Server - AI Assistant Guide

This project is an Excel Model Context Protocol (MCP) server that enables AI assistants to interact with Excel files through natural language commands. Built with FastMCP framework and openpyxl.

## Architecture Overview

### Core Components
- **src/server.py**: FastMCP server entry point with MCP tool definitions
- **src/core/**: Modular Excel operations
  - `excel_reader.py`: Reading operations
  - `excel_writer.py`: Writing/modification operations
  - `excel_manager.py`: File and worksheet management
  - `excel_search.py`: Regex search functionality
  - `excel_compare.py`: Game development specialized comparison
- **src/utils/**: Utilities with unified error handling and formatters
- **src/models/**: Type definitions and data models

### Key Design Patterns

#### Unified Error Handling
All MCP tools use `@unified_error_handler` decorator pattern:
```python
@mcp.tool()
@unified_error_handler("operation_name", extract_context_fn, return_dict=True)
def excel_operation(...):
    # Implementation delegates to core modules
```

#### Result Formatting
Consistent result format using `format_operation_result()`:
```python
return {
    'success': bool,
    'data': Any,  # Core result data
    'message': str,
    'metadata': dict  # Additional context
}
```

#### Range Expression Patterns
Supports two range formats:
- With sheet: `"Sheet1!A1:C10"` or `"TrSkill!A1:Z100"`
- Without sheet: `"A1:C10"` (requires separate sheet_name parameter)

## Development Workflows

### Running the Server
```bash
# Development
python -m src.server

# Testing
pytest tests/ -v
```

### MCP Client Configuration
Add to your MCP client config:
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "python",
      "args": ["-m", "src.server"],
      "env": {"PYTHONPATH": "${workspaceRoot}"}
    }
  }
}
```

### Testing Strategy
- Comprehensive test coverage in `tests/`
- Fixture-based testing with `sample_excel_file`
- Server interface testing in `test_server.py`
- Each core module has dedicated test files

## Project-Specific Conventions

### Game Development Focus
- Excel comparison tools specialized for game configuration tables
- ID-based object tracking (new, modified, deleted objects)
- Supports Chinese worksheet names with fallback mechanisms

### Error Context Extraction
- `extract_file_context()`: Captures file path and operation context
- `extract_formula_context()`: Captures formula evaluation context

### Excel Operation Patterns
- 1-based indexing to match Excel conventions
- Preserve formulas by default (`preserve_formulas=True`)
- Support for both .xlsx and .xlsm formats

## Key Dependencies
- **FastMCP**: MCP server framework
- **openpyxl**: Core Excel file operations
- **xlcalculator/formulas**: Formula evaluation engines
- **xlwings**: Optional Excel application integration
- **pytest/pytest-asyncio**: Testing framework

## Common Operations

### File and Sheet Management
- Create files with optional sheet names
- Sheet CRUD operations with Chinese name support
- Automatic active sheet management

### Data Operations
- Range-based read/write with format preservation
- Row/column insertion and deletion
- Cell formatting with presets (title, header, data, highlight, currency)

### Search and Analysis
- Regex search across files and directories
- Game-focused Excel comparison for configuration tables
- Formula evaluation without file modification

When working with this codebase, always use the unified error handling patterns, delegate implementation to core modules, and maintain the consistent result formatting.
