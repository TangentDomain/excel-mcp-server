# 📊 Technical Architecture

## Architecture Overview

ExcelMCP is built with a layered architecture designed for high-performance Excel operations with special optimizations for game development workflows.

## System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    MCP Client Layer                          │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐ ┌─────────┐ │
│  │ Claude      │ │ Cursor      │ │ VS Code     │ │ Others  │ │
│  │ Desktop     │ │             │ │ MCP         │ │         │ │
│  └─────────────┘ └─────────────┘ └─────────────┘ └─────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   MCP Protocol Layer                         │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                 FastMCP Core                              │ │
│  └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   ExcelMCP Server Layer                    │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                 server.py                               │ │
│  │  • Tool Definition                                        │ │
│  │  • Request Routing                                        │ │
│  │  • Response Formatting                                    │ │
│  └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   Business Logic Layer                      │
│  ┌─────────────────────┐ ┌─────────────────────┐ ┌─────────┐ │
│  │   excel_operations  │ │ advanced_sql_query  │ │  utils   │ │
│  │    API Layer        │ │     Engine         │ │  Layer   │ │
│  └─────────────────────┘ └─────────────────────┘ └─────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   Core Operations Layer                      │
│  ┌─────────────────────┐ ┌─────────────────────┐ ┌─────────┐ │
│  │    openpyxl Core    │ │  python-calamine   │ │   xlcalc  │ │
│  │    • Streaming      │ │    • Fast Read     │ │ • Formulas│ │
│  │    • Memory Opt     │ │    • Large Files   │ │ • Calc   │ │
│  └─────────────────────┘ └─────────────────────┘ └─────────┘ │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   Data Layer                                │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                  Excel Files (.xlsx, .xls, .xlsm)        │ │
│  └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

## Detailed Component Breakdown

### 1. MCP Client Layer
**Purpose**: Interface with various AI clients and development environments

**Supported Clients:**
- **Claude Desktop**: Native MCP integration
- **Cursor**: VS Code-based MCP support
- **VS Code**: Continue/Cline plugin integration
- **Cherry Studio**: Direct MCP connection
- **OpenClaw**: Built-in MCP server
- **OpenAI ChatGPT Plugin**: MCP proxy support

**Client Configuration:**
```json
{
  "mcpServers": {
    "excelmcp": {
      "command": "uvx",
      "args": ["excel-mcp-server-fastmcp"]
    }
  }
}
```

### 2. MCP Protocol Layer
**Purpose**: Handle MCP protocol communication and standardization

**Components:**
- **FastMCP Core**: MCP protocol implementation
- **Request Routing**: Standard tool call routing
- **Response Formatting**: Standardized response structure
- **Connection Management**: Connection pooling and health checks

**Protocol Features:**
- Tool discovery and metadata
- Streaming response support
- Error handling and validation
- Connection multiplexing

### 3. ExcelMCP Server Layer
**Purpose**: Main server interface and tool orchestration

**Key Components:**
```python
# server.py - Main server implementation
class ExcelMCPServer:
    def __init__(self):
        self.tool_registry = ToolRegistry()
        self.request_handler = RequestHandler()
        self.response_formatter = ResponseFormatter()
    
    def register_tools(self):
        # Register all 53 ExcelMCP tools
        pass
    
    def handle_request(self, request):
        # Route MCP requests to appropriate handlers
        pass
```

**Tool Categories:**
- **File Operations**: Workbook and sheet management
- **Data Operations**: Read, write, and update operations
- **SQL Operations**: Advanced query capabilities
- **Analysis Operations**: Data analysis and statistics
- **Performance Operations**: Optimized data processing

### 4. Business Logic Layer
**Purpose**: Implement Excel-specific business logic

#### excel_operations.py
```python
class ExcelOperations:
    def __init__(self):
        self.workbook_manager = WorkbookManager()
        self.sheet_manager = SheetManager()
        self.data_manager = DataManager()
    
    def execute_operation(self, operation, params):
        # Execute Excel operations with business logic
        pass
```

**Key Features:**
- **Workbook Management**: Open, close, copy operations
- **Sheet Management**: Create, delete, rename sheets
- **Data Operations**: Cell, row, column operations
- **Data Validation**: Input validation and error handling
- **Performance Optimization**: Streaming and batch operations

#### advanced_sql_query.py
```python
class AdvancedSQLQueryEngine:
    def __init__(self):
        self.sql_parser = SQLParser()
        self.query_optimizer = QueryOptimizer()
        self.result_formatter = ResultFormatter()
    
    def execute_query(self, excel_path, query):
        # Execute SQL queries on Excel data
        pass
```

**SQL Capabilities:**
- **Complex Queries**: JOIN, GROUP BY, subqueries
- **Query Optimization**: Index-based optimization
- **Result Caching**: Query result caching
- **Error Handling**: Structured error messages

### 5. Core Operations Layer
**Purpose**: Low-level Excel file operations

#### openpyxl Integration
```python
class OpenpyxlCore:
    def __init__(self):
        self.streaming_mode = False
        self.memory_manager = MemoryManager()
    
    def create_streaming_workbook(self, path):
        # Create optimized workbook for large files
        pass
    
    def write_streaming_data(self, path, sheet, data_stream):
        # Stream data for large files
        pass
```

**Optimizations:**
- **Streaming Write**: Write-only mode for large files
- **Memory Management**: Automatic cleanup and monitoring
- **Parallel Processing**: Multi-threaded operations
- **Compression**: Optional data compression

#### python-calamine Integration
```python
class CalamineIntegration:
    def __init__(self):
        self.fast_reader = FastExcelReader()
    
    def read_large_file(self, path, sheet=None):
        # Fast reading for large Excel files
        pass
    
    def detect_format(self, path):
        # Auto-detect Excel format
        pass
```

**Features:**
- **Fast Reading**: Optimized for large files
- **Format Detection**: Auto-detect .xlsx, .xls, .xlsm
- **Memory Efficient**: Streaming read operations
- **Performance**: 10x faster than traditional methods

### 6. Data Layer
**Purpose**: File system and data storage

**Supported Formats:**
- **.xlsx**: Modern Excel format (default)
- **.xls**: Legacy Excel format
- **.xlsm**: Macro-enabled Excel files
- **.csv**: CSV import/export support
- **.json**: JSON import/export support

**File Management:**
- **File Validation**: Check file integrity
- **Format Conversion**: Convert between formats
- **Backup Support**: Automated backup capabilities
- **Version Control**: Integration with git workflows

## Performance Architecture

### Memory Management
```python
class MemoryManager:
    def __init__(self):
        self.memory_limit = 512  # MB
        self.current_usage = 0
        self.cache = QueryCache()
    
    def allocate_memory(self, size):
        # Allocate memory for operations
        pass
    
    def cleanup(self):
        # Clean up unused memory
        pass
```

**Memory Features:**
- **Memory Limits**: Configurable memory limits
- **Smart Caching**: LRU cache for query results
- **Streaming Operations**: Memory-efficient data processing
- **Automatic Cleanup**: Automatic memory management

### Performance Monitoring
```python
class PerformanceMonitor:
    def __init__(self):
        self.metrics = {
            'query_time': [],
            'memory_usage': [],
            'cache_hits': [],
            'operation_count': []
        }
    
    def record_operation(self, operation, duration, memory_used):
        # Record performance metrics
        pass
    
    def get_performance_report(self):
        # Generate performance report
        pass
```

**Metrics Tracked:**
- **Query Performance**: Response times and optimization
- **Memory Usage**: Peak memory and cleanup efficiency
- **Cache Performance**: Hit rates and effectiveness
- **Operation Statistics**: Success/failure rates

### Caching System
```python
class QueryCache:
    def __init__(self):
        self.cache = {}
        self.max_size = 1000
        self.ttl = 3600  # 1 hour
    
    def get(self, key):
        # Get cached result
        pass
    
    def set(self, key, value):
        # Set cached result
        pass
    
    def clear(self):
        # Clear all cache
        pass
```

**Cache Features:**
- **Query Result Caching**: Cache for repeated queries
- **TTL Support**: Time-based expiration
- **Size Management**: Configurable cache size
- **Invalidation**: Smart cache invalidation

## Error Handling Architecture

### Error Types
```python
class ExcelMCPErrors:
    class ExcelMCPError(Exception):
        pass
    
    class FileError(ExcelMCPError):
        pass
    
    class QueryError(ExcelMCPError):
        pass
    
    class ValidationError(ExcelMCPError):
        pass
    
    class PerformanceError(ExcelMCPError):
        pass
```

### Error Handling Flow
1. **Error Detection**: Detect errors during operations
2. **Error Classification**: Classify error type and severity
3. **Error Resolution**: Attempt automatic resolution
4. **Error Reporting**: Report errors with detailed context
5. **Error Recovery**: Recover from recoverable errors

### Error Recovery
```python
class ErrorRecovery:
    def __init__(self):
        self.recovery_strategies = {
            'file_locked': self.retry_with_delay,
            'memory_exceeded': self.reduce_memory_usage,
            'query_timeout': self.optimize_query,
            'corrupted_file': self.repair_file
        }
    
    def recover_error(self, error):
        # Attempt error recovery
        pass
```

## Configuration System

### Configuration Files
```python
# config/default.yaml
server:
  host: "localhost"
  port: 18789
  memory_limit_mb: 512
  cache_size: 1000

performance:
  streaming_enabled: true
  parallel_processing: true
  query_optimization: true

logging:
  level: "INFO"
  file_path: "logs/excel_mcp.log"
```

### Environment Variables
```bash
EXCELMCP_MEMORY_LIMIT=512
EXCELMCP_CACHE_SIZE=1000
EXCELMCP_LOG_LEVEL=INFO
EXCELMCP_PERFORMANCE_MODE=fast
```

## Security Architecture

### Input Validation
```python
class InputValidator:
    def validate_path(self, path):
        # Validate file paths
        pass
    
    def validate_query(self, query):
        # Validate SQL queries
        pass
    
    def validate_data(self, data):
        # Validate input data
        pass
```

### File Security
- **Path Validation**: Prevent directory traversal
- **File Type Checking**: Validate Excel file types
- **Size Limits**: Prevent oversized files
- **Access Control**: File permissions checking

## Testing Architecture

### Test Categories
- **Unit Tests**: Individual component testing
- **Integration Tests**: API integration testing
- **Performance Tests**: Load and stress testing
- **Game Testing**: Game-specific scenario testing

### Test Framework
```python
# tests/test_excel_operations.py
class TestExcelOperations:
    def test_write_operations(self):
        # Test writing operations
        pass
    
    def test_query_operations(self):
        # Test query operations
        pass
    
    def test_performance(self):
        # Test performance optimizations
        pass
```

## Deployment Architecture

### Installation Methods
1. **uvx Installation**: Single command installation
2. **pip Installation**: Traditional Python package
3. **Source Installation**: From source code
4. **Docker**: Containerized deployment

### System Requirements
- **Python**: 3.10+ required
- **Memory**: 512MB minimum, 1GB recommended
- **Disk**: 100MB minimum for installation
- **Network**: Internet connection for package installation

## Future Architecture Enhancements

### Planned Improvements
1. **GPU Acceleration**: GPU-based mathematical operations
2. **Distributed Processing**: Multi-machine processing for large files
3. **Cloud Integration**: Direct cloud storage integration
4. **AI Optimization**: AI-powered query optimization
5. **Mobile Support**: Mobile client integration

### Scalability Design
- **Horizontal Scaling**: Multi-instance deployment
- **Load Balancing**: Request distribution
- **Database Integration**: External database support
- **API Gateway**: Centralized API management