# RULES — 执行细则

> 子代理可进化：发现更好的做法→更新本文件→记录到DECISIONS.md。
> 但红线（.cron-prompt.md）和文档所有权不可通过自我进化修改。

## 时间分配（每轮45分钟）
- 需求实现：25分钟
- 测试+修复：10分钟
- MCP验证+README检查+评价+合并+文档更新：10分钟

## 测试策略
- **开发中**：只跑受影响的测试文件（3-5秒）
- **全量测试仅两个时机**：第一轮评估基线 + 合并main前
- **全量命令**：`python3 -m pytest tests/ -q --tb=no -n auto --timeout=30`
- **exec timeout至少600秒**，pytest --timeout=30防止单个测试卡住
- **不要单线程跑全量**，必须用pytest-xdist并行

## 发布流程（合并main后，有代码改动时执行）
1. 更新版本号：pyproject.toml + src/excel_mcp_server_fastmcp/__init__.py
2. 构建：`rm -rf dist && ~/.local/bin/uv build`
3. 全量测试：`python3 -m pytest tests/ -q --tb=short -n auto --timeout=30`
4. 本地测试：`pip install dist/*.whl --force-reinstall --break-system-packages --no-deps && timeout 3 excel-mcp-server-fastmcp`
5. 发布：`proxychains4 ~/.local/bin/uv publish dist/*.whl dist/*.tar.gz --token <TOKEN>（token见.cron-prompt.md）`
6. 远程验证：`pip install excel-mcp-server-fastmcp --dry-run --break-system-packages`
7. tag+push：`git tag vX.Y.Z && proxychains4 git push origin main develop --tags`
- 仅纯文档改动→只合并不发布
- 版本号：功能→minor，bug修复→patch

## 分支策略
- 每个需求开独立feature分支+worktree
- `git worktree add ../wt-<REQ-ID> -b feature/<REQ-ID> develop`
- 一次只开一个worktree，完成清理后再开下一个
- 测试通过→合develop→合main→清理worktree
- 做不好→丢弃worktree，不影响develop
- 清理：`git worktree remove ../wt-<ID> && git branch -d feature/<ID>`

## 代码规范
- 不删功能、不改架构、不加依赖（CEO可覆盖）
- 所有文件读写显式指定encoding='utf-8'
- 不要修改已有的CI修复（如encoding='utf-8'显式指定）
- 需求必须有验收标准，没有验收标准不能标DONE

## 包结构
```
src/excel_mcp_server_fastmcp/
├── __init__.py              # 入口，暴露 main()
├── server.py                # MCP接口层（工具定义）
├── api/                     # API业务逻辑层
│   ├── excel_operations.py
│   └── advanced_sql_query.py
├── core/                    # 核心操作层
├── models/types.py
└── utils/                   # 工具层
```
- server.py的import使用相对导入
- grep工具数量：`grep -c "def excel_" src/excel_mcp_server_fastmcp/server.py`

## 安装方式
- uvx：`uvx excel-mcp-server-fastmcp`
- 源码：`pip install -e .`
- MCP配置：`{"command": "uvx", "args": ["excel-mcp-server-fastmcp"]}`

## 产出格式
📊 第X轮完成（共Y轮）
• 需求：[REQ-XXX] 简述 → [状态]
• README：[已同步/无需更新]
🧪 MCP：[内容] → 结果 | 准确性 X/Y
📋 评分：XX/100
🔀 合并：[状态]
💡 评价：
  功能：策划[?] AI[?] | 准确性[?]
  工程：复杂度[?] 重复[?] 测试质量[?] 性能[?] 架构[?]
  方向：短板→[?] 瓶颈→[?]
📋 需求池：X open / Y in-progress / Z done
🐛 待修：[列表]
🔄 自我进化：[有/无]
