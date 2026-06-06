# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| 1.17.x  | ✅ |
| < 1.17  | ❌ |

## Reporting a Vulnerability

If you discover a security vulnerability in excel-mcp-server,
please report it privately via GitHub Security Advisory:

https://github.com/TangentDomain/excel-mcp-server/security/advisories/new

Please do not report security vulnerabilities via public GitHub Issues.

## Security Features

- **Path traversal protection**: All file paths are validated through `SecurityValidator` before access
- **SQL injection prevention**: All SQL queries are parsed through structured AST (sqlglot), not string interpolation
- **Safe eval**: `excel_run_python` and formula evaluation use sandboxed environments with restricted globals
- **No network access**: The MCP server operates on local files only; no outbound network requests
