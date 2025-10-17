# Excel MCP Server 监控和维护系统

本目录包含用于监控和维护 Excel MCP Server 项目的完整工具集。

## 📁 文件说明

### 核心脚本
- **`monitor-and-maintain.py`** - 完整的监控和维护脚本，提供全面的监控功能
- **`quick-monitor.py`** - 快速监控脚本，专注于核心功能的快速检查
- **`run-monitor.bat`** - Windows 批处理启动器，简化监控工具的使用
- **`monitor-config.json`** - 监控系统配置文件

### 功能特性

#### 1. 覆盖率监控
- 检查测试覆盖率是否达到要求 (默认 85%)
- 分析每个文件的覆盖率情况
- 识别未覆盖的文件和代码行
- 生成详细的覆盖率报告

#### 2. 性能监控
- 监控测试执行时间
- 检查测试成功率
- 识别最慢的测试用例
- 分析内存使用情况

#### 3. 内存使用监控
- 实时监控测试过程中的内存使用
- 检测内存泄漏
- 分析内存增长趋势
- 提供内存优化建议

#### 4. 测试质量评估
- 计算测试稳定性分数
- 检测不稳定的测试用例 (flaky tests)
- 评估代码质量分数
- 分析测试复杂度和重复性

#### 5. 自动报告生成
- 生成美观的 HTML 报告
- 支持文本格式报告
- 保存 JSON 格式的监控数据
- 历史报告管理和归档

#### 6. 维护建议
- 基于监控结果自动生成改进建议
- 覆盖率改进建议
- 性能优化建议
- 质量提升建议

## 🚀 快速开始

### 方式一：使用批处理启动器 (推荐)

```bash
# 快速监控 (1-2分钟)
run-monitor.bat

# 完整监控 (5-10分钟)
run-monitor.bat full

# 查看帮助
run-monitor.bat help
```

### 方式二：直接运行 Python 脚本

```bash
# 快速监控
python scripts/quick-monitor.py

# 完整监控
python scripts/monitor-and-maintain.py

# 仅覆盖率监控
python scripts/monitor-and-maintain.py --coverage-only

# 连续监控模式
python scripts/monitor-and-maintain.py --continuous --interval 5
```

### 方式三：使用 Python 模块方式

```bash
# 推荐方式 - 避免路径问题
python -m scripts.monitor-and-maintain

# 快速监控
python -m scripts.quick-monitor
```

## 📊 报告说明

### HTML 报告
完整的可视化报告，包含：
- 📊 总体评估和综合评分
- 📈 详细的覆盖率分析
- ⚡ 性能指标和趋势
- 🎯 代码质量评估
- 💡 个性化改进建议

报告位置：`reports/monitoring-report-YYYYMMDD-HHMMSS.html`

### 文本报告
简洁的文本格式报告，适合：
- 命令行快速查看
- 日志记录
- 自动化脚本集成

### JSON 数据
结构化的监控数据，适合：
- 数据分析
- 历史趋势分析
- 自动化处理
- API 集成

## ⚙️ 配置说明

编辑 `monitor-config.json` 文件来自定义监控设置：

```json
{
  "monitoring": {
    "coverage_threshold": 85.0,        // 覆盖率阈值
    "max_execution_time": 300,         // 最大执行时间(秒)
    "max_memory_usage": 1000          // 最大内存使用(MB)
  },
  "quality_thresholds": {
    "min_test_stability": 90.0,        // 最小测试稳定性
    "min_code_quality": 80.0,          // 最小代码质量分数
    "max_duplicate_coverage": 20.0     // 最大重复覆盖率
  }
}
```

## 📈 监控指标说明

### 覆盖率指标
- **总覆盖率**: 整体代码覆盖率百分比
- **文件覆盖率**: 每个文件的覆盖率详情
- **未覆盖文件**: 完全未测试的文件列表
- **缺失行**: 具体未覆盖的代码行

### 性能指标
- **执行时间**: 总测试执行时间
- **测试数量**: 总测试用例数
- **成功率**: 通过测试的百分比
- **最慢测试**: 执行时间最长的测试用例
- **内存使用**: 测试过程中的内存占用

### 质量指标
- **测试稳定性**: 多次运行结果的一致性
- **代码质量**: 基于多维度评估的质量分数
- **测试复杂度**: 测试用例的复杂度评分
- **重复覆盖率**: 重复测试用例的比例

## 🔧 高级用法

### 连续监控模式

```bash
# 每5分钟监控一次
python scripts/monitor-and-maintain.py --continuous

# 每10分钟监控一次
python scripts/monitor-and-maintain.py --continuous --interval 10
```

### 自定义报告输出

```bash
# 指定报告文件路径
python scripts/monitor-and-maintain.py --report-file /path/to/report.html

# 仅生成文本报告
python scripts/monitor-and-maintain.py --no-html --report-file report.txt
```

### 调整覆盖率阈值

```bash
# 设置覆盖率阈值为90%
python scripts/monitor-and-maintain.py --threshold 90
```

### 监控特定功能

```bash
# 仅运行覆盖率监控
python scripts/monitor-and-maintain.py --coverage-only

# 仅运行性能监控
python scripts/monitor-and-maintain.py --performance-only

# 仅运行质量评估
python scripts/monitor-and-maintain.py --quality-only
```

## 📋 监控结果解读

### 综合评分说明

| 分数范围 | 状态 | 说明 |
|---------|------|------|
| 90-100 | 优秀 | 所有指标都在理想范围内 |
| 80-89 | 良好 | 大部分指标正常，有少量改进空间 |
| 70-79 | 一般 | 某些指标需要关注和改进 |
| 60-69 | 需要改进 | 多个指标不达标，需要重点关注 |
| 0-59 | 严重问题 | 存在严重的质量问题 |

### 常见问题和解决方案

#### 覆盖率低
**问题**: 测试覆盖率低于85%
**建议**:
- 增加缺失功能的测试用例
- 覆盖边界条件和异常情况
- 关注未测试的工具函数

#### 测试执行慢
**问题**: 测试执行时间超过5分钟
**建议**:
- 优化测试数据和设置
- 使用 Mock 减少外部依赖
- 并行化独立测试用例

#### 内存使用高
**问题**: 测试内存使用超过1GB
**建议**:
- 检查测试数据清理
- 优化文件处理逻辑
- 减少测试用例中的大对象

#### 测试不稳定
**问题**: 存在 flaky tests
**建议**:
- 检查测试中的随机因素
- 增加适当的等待和重试逻辑
- 确保测试环境的稳定性

## 🔄 自动化集成

### GitHub Actions 集成

```yaml
name: Code Quality Monitor
on: [push, pull_request]

jobs:
  monitor:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Set up Python
        uses: actions/setup-python@v2
        with:
          python-version: '3.10'
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -e .
      - name: Run monitoring
        run: python scripts/monitor-and-maintain.py --no-html
      - name: Upload reports
        uses: actions/upload-artifact@v2
        with:
          name: monitoring-reports
          path: reports/
```

### 定时任务设置

```bash
# 添加到 crontab，每天早上9点运行
0 9 * * * cd /path/to/excel-mcp-server && python scripts/monitor-and-maintain.py

# Windows 任务计划程序
# 创建基本任务，每天9点运行：python scripts/monitor-and-maintain.py
```

## 🛠️ 故障排除

### 常见问题

1. **ImportError: No module named 'pytest'**
   ```bash
   pip install pytest pytest-cov
   ```

2. **PermissionError: 无法创建报告目录**
   ```bash
   mkdir reports
   chmod 755 reports
   ```

3. **覆盖率报告文件不存在**
   - 确保测试成功运行
   - 检查 `coverage.json` 文件是否生成

4. **内存监控数据异常**
   - 确保系统支持 psutil 模块
   ```bash
   pip install psutil
   ```

### 调试模式

```bash
# 启用详细日志
export EXCELMCP_DEBUG=1
python scripts/monitor-and-maintain.py --coverage-only

# Windows
set EXCELMCP_DEBUG=1
python scripts\monitor-and-maintain.py --coverage-only
```

## 📞 支持和反馈

如果遇到问题或有改进建议，请：

1. 查看日志文件：`logs/monitor.log`
2. 检查配置文件：`scripts/monitor-config.json`
3. 运行快速诊断：`python scripts/quick-monitor.py`
4. 提交 Issue 到项目仓库

---

**注意**: 首次运行监控时，请确保项目依赖已正确安装，测试环境已配置完毕。