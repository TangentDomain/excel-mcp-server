# RULES — 执行细则

> 子代理可进化：发现更好的做法→更新本文件→记录到DECISIONS.md。
> 但红线（.cron-prompt.md）和文档所有权不可通过自我进化修改。
> 敏感文档（.cron-prompt.md/ROADMAP.md/REQUIREMENTS.md优先级和红线）修改时需给方案让CEO决策。其他文档自由修改。

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
   - README只写MCP使用方式，不写Python API代码（读者是策划/分析师，不是开发者）
3. 发现不一致 → 自动修复 + 记录到DECISIONS.md
4. 清理脚本维护：归档过期版本检查历史

🔄 **效率追踪**（2026-03-27 R131，docstring质量提升规则修改后第1轮）
**改前基线**: docstring评分 2个excellent → 6个excellent，质量提升200%
**预期效果**: AI工具使用体验提升，减少用户查询成本
**验证方式**: 每轮docstring质量统计对比用户反馈

## 智能文档索引系统（新增规则 - self-evolve）
**问题**: 项目文档数量庞大（29个），用户查找信息效率低下，testing-guidelines.md达63KB
**解决**: 实施按用户角色和功能模块的智能文档索引系统
**规则**:
1. 创建 `docs/INDEX.md` 文档索引，按用户角色分类：
   - **游戏策划**: 主要关注技能/装备/怪物等配置管理
   - **程序开发者**: 主要关注API调用、工具集成、技术实现  
   - **运维人员**: 主要关注安装部署、故障排查、性能优化
2. 创建 `docs/NAVIGATION.md` 导航指南，提供：
   - 文档依赖关系图
   - 查找流程算法（根据用户问题推荐3个最相关文档）
   - 快速跳转链接
3. 每轮开工执行 `python3 scripts/check-doc-index.py` 验证索引完整性
4. 发现缺失或过时的索引 → 立即更新 + 记录到DECISIONS.md
5. 大文档拆分：testing-guidelines.md > 10KB部分拆分为多个专题文档

🔄 **效率追踪**（2026-03-30 R224，文档索引系统问题已解决 - self-evolve）
**改前基线**: 用户需要浏览29个文档才能找到相关信息，平均查找时间3分钟
**预期效果**: 用户通过索引系统在30秒内找到目标文档，查找效率提升80%
**验证方式**: 用户反馈收集 + 索引使用统计对比 + 模拟用户查找任务耗时测量
**当前状态**: ✅ 已完成 - INDEX.md(3.7KB)、NAVIGATION.md(5.1KB)、testing-guidelines.md(5KB)全部达标，check-doc-index.py监控脚本运行正常

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
- **CI失败是阻断性问题**：每轮用 `gh run list --limit 1` 检查CI状态，失败则 `gh run view --log-failed` 分析→定位→修复→推送。CI红灯不能无视
- **仓库整洁**：不往根目录/子代理目录创建临时文件，内部文件必须被.gitignore覆盖。scripts/只放对用户有用的脚本，docs/只放对外文档
- **禁止合并冲突残留**：合并后必须 `grep -rn "<<<<<<" .` 检查，有冲突标记必须解决干净再commit。提交后发现冲突，当轮立即修复
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

## Harness Engineering 反哺（2026-03-30）

### 生成/评估分离
- **每轮必须有 git commit**：没有 commit = 没干活 = 本轮失败
- **不要自我评价"做得很好"**：自评不可靠，用客观指标代替（测试通过数、代码行数变更、commit message 质量）
- **代码变更必须有测试覆盖**：改了功能就必须跑测试，不能只改不测
- **禁止重做 DONE 需求**：选任务前先读 REQUIREMENTS.md，只做 status=OPEN 的。重复劳动 = 浪费轮次
- **"超额完成"是警告信号**：如果某轮声称超额完成，检查是否真的做了新东西还是重做了旧东西

### 上下文管理
- **涉及 >5 个文件操作时，用 Claude Code 外包**（见 .cron-prompt.md 上下文减压章节）
- **每轮开始先读 NOW.md**：知道上轮做到哪，接着做而不是从头开始

## Conventional Commits 规范（REQ-029）
- **格式要求**：`[REQ-XXX] type: 简述`
  - 必须包含需求编号前缀：`[REQ-XXX]`
  - type 必须是以下之一：
    - `feat`: 新功能
    - `fix`: 修复 bug
    - `refactor`: 重构（既不是新功能也不是修复 bug）
    - `docs`: 文档更新
    - `test`: 测试相关
    - `chore`: 构建或辅助工具变动
    - `perf`: 性能优化
- **示例**：
  - `[REQ-028] fix: insert_mode 默认值改为 false`
  - `[REQ-029] feat(api): 添加 docstring 契约验证脚本`
- **提交规则**：
  1. 简短描述必须以动词开头（小写）
  2. 详细描述可选，但建议超过50行代码时添加
  3. 多个段落用空行分隔
  4. 禁止使用 `Fixes #123` 或 `Closes #123`，改为使用 `Closes REQ-XXX`
- **CI 验证**：违反规范的提交信息会被 CI 检查拒绝，必须使用 `git commit --amend` 修正
