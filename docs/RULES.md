# RULES — 执行细则

> 子代理可进化：发现更好的做法→更新本文件→记录到DECISIONS.md。
> 但红线（.cron-prompt.md）和文档所有权不可通过自我进化修改。

## 文档瘦身（每轮第0步，不消耗改进时间）
- DECISIONS.md > 40条 → 最早的10条归档到 docs/DECISIONS-ARCHIVED.md
- REQUIREMENTS.md DONE项 → 移入 ARCHIVED.md
- NOW.md > 30行 → 精简历史记录，只保留最近3轮
- 超限时不瘦身，本轮产出无效

## 自动化版本检查（新增规则）
**问题**: 文档同步依赖手动操作，容易出错且耗时
**解决**: 每轮自动检查版本一致性，异常时立即修复
**规则**:
1. 每轮开工执行 `python3 scripts/check-version-sync.py`（需创建）
2. 检查项目：pyproject.toml、__init__.py、README.md、README.en.md、CHANGELOG
3. 发现不一致 → 自动修复 + 记录到DECISIONS.md
4. 清理脚本维护：归档过期版本检查历史

🔄 **效率追踪**（2026-03-27 R131，docstring质量提升规则修改后第1轮）
**改前基线**: docstring评分 2个excellent → 6个excellent，质量提升200%
**预期效果**: AI工具使用体验提升，减少用户查询成本
**验证方式**: 每轮docstring质量统计对比用户反馈

## 时间分配（每轮45分钟）
- 需求实现：25分钟
- 测试+修复：10分钟
- MCP验证+README检查+评价+合并+文档更新：10分钟

## 项目健康度自检（每20轮至少1次）
- **目的**：自己发现代码质量问题，自己解决，不等CEO指示
- **检查项**：
  1. **根目录垃圾文件**：是否有临时脚本、测试文件散落在根目录（应在tests/内）
  2. **测试冗余**：多个测试文件测同一功能（如安全测试3个文件测同一件事）
  3. **轮次编号测试**：test_req010_r67.py这类带轮次编号的文件，应合并到功能文件后删除
  4. **文档膨胀**：DECISIONS/NOW/REQUIREMENTS是否超限（已有瘦身规则）
  5. **废弃分支**：git branch | grep feature/ 是否有过期worktree未清理
  6. **依赖变化**：pyproject.toml是否有不该引入的新依赖
- **发现问题时**：立即修复，不需要等CEO批准。修复后记入DECISIONS.md
- **产出**：本轮指标里加一行"健康度自检：发现X项，修复X项"

## MCP验证
- **开发中MCP验证**：有功能变化时，至少8项游戏场景通过MCP工具调用
- **MCP真实验证（每5轮至少1次）**：创建真实xlsx测试文件，通过MCP工具实际调用12项核心功能（list_sheets/get_range/query WHERE/query JOIN/query GROUP BY/query子查询/query FROM子查询/get_headers/find_last_row/batch_insert_rows/delete_rows/describe_table），记录通过/失败。发现的bug立即写入REQUIREMENTS.md。测试文件用后清理。 pytest只验证代码逻辑，MCP真实验证验证端到端可用性，两者不可互相替代。

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

## 产出格式（客观指标版）
📊 第X轮完成（共Y轮）
• 需求：[REQ-XXX] 简述 → [状态]
• README：[已同步/无需更新]
🧪 MCP：[内容] → 结果 | 准确性 X/Y
🔀 合并：[状态]
📊 本轮指标：
  - 改动：{files} 文件 / +{added} -{removed} 行
  - 测试：{total} 通过 / +{new} 新增
  - 需求：{done} 完成 / {moved} 推进
  - 发布：[vX.Y.Z / 无]
  - 文档：{docs} 文件更新
📋 需求池：X open / Y in-progress / Z done
🐛 待修：[列表]
🔄 自我进化：[有/无]
📊 效率追踪：[改RULES.md后的5轮内填写：测试耗时 | 有效时间 | 痛点复发]

## 指标计算方法（子代理每轮必须执行）
- 开工前：`git rev-parse HEAD` 记录起始commit
- 完工后：`git diff --stat <start>..HEAD` → 文件数/行数
- `git log --oneline <start>..HEAD` → commit数
- `python3 -m pytest tests/ -q --tb=no -n auto --timeout=30 2>&1 | tail -1` → 测试总数
- 对比REQUIREMENTS.md修改前后 → 需求完成/推进数
- `git tag --points-at HEAD` → 是否有发布tag
🔄 效率追踪（2026-03-28 R165，文档同步规则修改后第2轮）
**改前基线**: 文档版本信息不一致，用户需要手动查找解决方案
**预期效果**: 提升用户体验，减少支持请求，文档信息准确可靠
**验证方式**: 用户反馈收集 + 文档访问统计对比
