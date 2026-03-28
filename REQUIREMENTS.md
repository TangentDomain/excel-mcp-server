# ExcelMCP 需求池

> 详细需求内容，需求状态变化时更新。
> 当前状态概览见 [docs/NOW.md](docs/NOW.md)，路线图见 [docs/ROADMAP.md](docs/归档见 [ARCHIVED.md](ARCHIVED.md)。

## 活跃需求

### REQ-035 [P0] CI CTE测试全平台失败

**状态: DONE
**优先级**: P0（阻断CI）
**发现轮次**: 185（MCP真实验证发现），198-200（CI持续失败确认），201（修复）

**问题描述**:
`tests/test_sql_enhanced.py::TestCTE` 3个测试在CI所有平台（Linux/macOS/Windows, Python 3.10-3.13）失败，仅macOS 3.10被skipif跳过。
本地Python 3.12 + python-calamine 0.6.2通过，但CI环境不同。

**已知信息**:
- 本地 `execute_advanced_sql_query(path, CTE_SQL)` 返回 `success: true`
- CI错误：`assert result['success'] is True` 失败，但 `--tb=short` 没有显示具体错误消息
- 错误信息被 `execute_advanced_sql_query` 的 except 块吞掉了，只返回 `{success: false, message: "..."}` 
- CI可能安装了不同版本的python-calamine（`>=0.3.0`范围太宽）
- 1155个其他测试全部通过，只有CTE 3个失败

**修复方向**:
1. CI改用 `--tb=long` 获取完整traceback，确认具体错误
2. 测试中打印 `result['message']` 看实际错误信息
3. pin python-calamine到已知工作版本，或修代码兼容所有版本
4. 可能原因：python-calamine新版本读xlsx后DataFrame结构变化，CTE依赖的双行表头检测失效

**验收标准**:
- CI所有平台（11个job）全部通过
### REQ-036 [P1] README新手友好化改造

**状态**: OPEN
**优先级**: P1
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
