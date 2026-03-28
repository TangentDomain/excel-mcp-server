# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求


**状态**: DONE
**优先级**: P1
**完成轮次**: 204
**设计稿**: `docs/readme-redesign.md`（CEO已审核确认）

**问题**:
当前README对不熟悉Python生态的用户不友好：uvx是什么不解释、没有前置条件说明、没分客户端教程、JSON配置不知道往哪放。

**改造要求**:
1. 头部简化：去掉`<div align="center">`和过多badge（只留PyPI/CI/Tests/Tools 4个）
2. 加"这是什么"段落：3句话说清楚用途、用户群、前置条件
3. 安装教程改为"5分钟上手"分4步：确认Python→装工具→配客户端→开始用
4. 分客户端教程：Claude Desktop（含配置文件路径）+ Cursor + Cherry Studio
5. uvx和pip两种方式都给，pip作为"更传统更稳定"的备选
6. 加FAQ折叠块：装uv报错、command not found、怎么确认成功、支持哪些客户端
7. 竞品对比、性能优化、SQL示例等技术细节后移（保留但不是第一屏）
8. 中文README和英文README同步改造

**不改的**:
- 后半部分的工具列表、SQL场景、技术架构等内容保持不变
- 版本号badge保留

**验收标准**:
- 一个不懂Python的人能跟着README完成安装和配置
- 中英文README同步
- 全文不出现Python API调用代码，只展示MCP对话方式（对AI说话）
- pip/uvx安装命令是唯一例外（告诉用户怎么装工具）
