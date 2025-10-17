## MODIFIED Requirements
### Requirement: 安全默认行为
MCP服务器 SHALL 确保所有危险操作默认使用最安全的模式，防止数据意外丢失。

#### Scenario: 更新操作安全模式
- **WHEN** 调用excel_update_range时
- **THEN** 系统 SHALL 默认使用insert_mode=True的安全插入模式，而非覆盖模式

#### Scenario: 删除操作保护
- **WHEN** 执行删除操作时
- **THEN** 系统 SHALL 要求明确的确认步骤，并显示将要删除的内容

#### Scenario: 文件操作安全
- **WHEN** 操作Excel文件时
- **THEN** 系统 SHALL 检查文件锁定状态，避免并发操作冲突

### Requirement: 操作预览和确认
MCP服务器 SHALL 在执行可能影响数据的操作前提供预览和确认机制。

#### Scenario: 操作范围预览
- **WHEN** 准备执行数据操作
- **THEN** 系统 SHALL 显示操作将要影响的单元格范围和当前数据内容

#### Scenario: 数据影响评估
- **WHEN** 分析操作影响
- **THEN** 系统 SHALL 评估操作将影响多少行、列和单元格，提供明确的统计信息

#### Scenario: 危险操作警告
- **WHEN** 操作可能影响大量数据
- **THEN** 系统 SHALL 发出明确的警告，并要求特别确认

### Requirement: 自动备份和恢复
MCP服务器 SHALL 提供自动备份和恢复能力，确保误操作不会造成永久数据损失。

#### Scenario: 自动备份创建
- **WHEN** 执行重大操作前
- **THEN** 系统 SHALL 自动创建当前文件的备份，包含时间戳信息

#### Scenario: 操作撤销能力
- **WHEN** 用户需要撤销操作
- **THEN** 系统 SHALL 提供一键恢复到最近备份状态的能力

#### Scenario: 备份管理
- **WHEN** 管理备份文件
- **THEN** 系统 SHALL 自动清理过期备份，避免存储空间无限增长

## ADDED Requirements
### Requirement: 安全的LLM提示词
MCP服务器 SHALL 提供以安全为核心的LLM提示词，指导LLM进行安全操作。

#### Scenario: 安全操作原则
- **WHEN** LLM需要操作Excel数据
- **THEN** 提示词 SHALL 强调"安全第一"的原则，推荐安全的操作序列

#### Scenario: 预览优先指导
- **WHEN** LLM准备执行数据操作
- **THEN** 提示词 SHALL 指导LLM先进行预览，再执行操作

#### Scenario: 安全工具推荐
- **WHEN** LLM选择操作工具
- **THEN** 提示词 SHALL 推荐使用安全的工具组合和参数设置

### Requirement: 多层参数验证
MCP服务器 SHALL 实现严格的参数验证，防止参数错误导致的数据破坏。

#### Scenario: 范围格式严格验证
- **WHEN** 接收range参数
- **THEN** 系统 SHALL 严格验证格式，确保包含工作表名且范围合理

#### Scenario: 数据量限制检查
- **WHEN** 接收批量操作请求
- **THEN** 系统 SHALL 检查操作规模，对过大操作进行警告

#### Scenario: 危险参数检测
- **WHEN** 接收可能导致数据丢失的参数
- **THEN** 系统 SHALL 进行特别警告和确认要求

### Requirement: 操作透明化
MCP服务器 SHALL 提供透明的操作过程，让用户清楚了解每个步骤。

#### Scenario: 操作步骤展示
- **WHEN** 执行复杂操作
- **THEN** 系统 SHALL 显示详细的操作步骤和进度

#### Scenario: 操作结果验证
- **WHEN** 操作完成后
- **THEN** 系统 SHALL 提供操作结果的验证和确认

#### Scenario: 操作历史记录
- **WHEN** 用户查看操作历史
- **THEN** 系统 SHALL 提供完整的操作日志和影响范围记录
