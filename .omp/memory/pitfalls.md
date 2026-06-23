# 已知陷阱

>踩过的坑，避免重复犯错。

## P001: memory URI 误提交

- **现象**: `memory://` URI 写入代码或文档被 git 跟踪
- **根因**: memory URI 是运行时内部协议，不应出现在源代码中
- **规避**: memory URI 仅用于 agent 对话和 harness 内部通信，禁止写入 src/ tests/ 或任何会被 git add 的文件

## P002: openpyxl 大文件慢

- **现象**: 大 Excel 文件（>10MB）读写操作延迟显著
- **根因**: openpyxl 是纯 Python 实现，无 C 扩展，逐单元格操作开销大
- **规避**: 大文件场景优先使用 calamine（Rust 实现）读取；写操作尽量批量而非逐单元格

## P003: sqlglot LIKE 转换 bug

- **现象**: 含 `%` 和 `_` 的 LIKE 模式经过 sqlglot 转换后语义变化
- **根因**: sqlglot 在方言转换中未正确保留 LIKE 通配符转义
- **规避**: LIKE 模式在传入 sqlglot 前先做 workaround 处理；或直接透传到 SQLite 执行

## P004: 双行表头歧义

- **现象**: 双行表头（中文+英文）导致列名识别歧义
- **根因**: 部分工具用第 1 行（中文），部分用第 2 行（英文），不统一
- **规避**: SQL 工具（query/update/insert/delete）中英文名都支持；describe_table 返回英文名；upsert_row 自动检测双行表头

## P005: Windows 弹窗

- **现象**: Windows 上创建子进程时弹出 cmd 窗口
- **根因**: Bun.spawn / child_process.spawn 未传 windowsHide: true
- **规避**: 所有 spawn 调用无条件传入 `windowsHide: true`，非 Windows 平台此选项被忽略

## P006: update_range 覆盖模式

- **现象**: 使用 update_range 追加数据却覆盖了已有内容
- **根因**: update_range 默认 insert_mode=False（覆盖模式）
- **规避**: 追加数据时必须传 insert_mode=True + 先 find_last_row 定位末行
