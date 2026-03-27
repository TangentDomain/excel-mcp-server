- 原因：构建工具属于CI/CD环节，不是MCP查询工具
- 影响：聚焦SQL引擎

## 2026-03-26 | calamine替换openpyxl读取
- 原因：2300x性能提升
- 影响：REQ-015读取部分完成

## 2026-03-26 | 取消View/写入校验/Auto Increment
- 原因：SQL引擎已覆盖（FK用JOIN查、范围用WHERE、枚举用IN）
- 影响：需求池精简，避免过度设计
