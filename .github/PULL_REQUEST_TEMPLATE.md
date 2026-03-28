---
name: Pull Request
description: 提交代码更改到 excel-mcp-server
title: "["
labels: ["pr"]
assignees: []
body:
  - type: markdown
    attributes:
      value: |
        感谢您的贡献！请填写以下信息来帮助代码审查过程。

  - type: dropdown
    id: change-type
    attributes:
      label: 变更类型
      description: 选择此 PR 的主要变更类型
      options:
        - "feat: 新功能"
        - "fix: 修复 Bug"
        - "docs: 文档更新"
        - "style: 代码格式化"
        - "refactor: 重构"
        - "test: 测试相关"
        - "chore: 构建或工具相关"
    validations:
      required: true

  - type: textarea
    id: change-description
    attributes:
      label: 变更描述
      description: 详细描述此 PR 实现的功能或修复的问题
      placeholder: "详细描述您的变更内容..."
    validations:
      required: true

  - type: textarea
    id: motivation
    attributes:
      label: 变更动机
      description: 为什么需要进行这个变更？解决了什么问题？
      placeholder: "解释进行此变更的原因..."
    validations:
      required: true

  - type: textarea
    id: testing
    attributes:
      label: 测试情况
      description: 描述您进行的测试
      placeholder: |
        - [x] 已通过所有现有测试
        - [x] 已添加新测试（如果适用）
        - [x] 手动测试通过
        - [ ] 性能测试通过（如果适用）
        - [ ] 兼容性测试通过（如果适用）
    validations:
      required: true

  - type: dropdown
    id: breaking-changes
    attributes:
      label: 破坏性变更
      description: 此变更是否包含破坏性变更？
      options:
        - "无破坏性变更"
        - "包含破坏性变更"
        - "不确定"
    validations:
      required: true

  - type: textarea
    id: breaking-changes-description
    attributes:
      label: 破坏性变更说明
      description: 如果包含破坏性变更，请详细说明
      placeholder: "描述具体的破坏性变更和影响..."
      validations:
        required: false

  - type: checkboxes
    id: checklist
    attributes:
      label: 检查清单
      description: 请确保完成以下事项
      options:
        - label: 我的代码遵循项目的代码规范
          required: true
        - label: 我已经自测试了我的代码
          required: true
        - label: 我已经考虑了兼容性问题
          required: true
        - label: 我已经更新了相关文档
          required: false
        - label: 我已经添加了适当的测试
          required: false
    validations:
      required: true

  - type: textarea
    id: additional-context
    attributes:
      label: 附加信息
      description: 任何其他有用的信息（如截图、设计文档等）
      placeholder: "提供额外的信息、截图、或相关文档链接..."
    validations:
      required: false

  - type: markdown
    attributes:
      value: |
        ## 📋 代码审查清单
        
        审查者请检查以下项目：
        
        - [ ] 代码逻辑正确
        - [ ] 性能影响评估
        - [ ] 安全性考虑
        - [ ] 测试覆盖度
        - [ ] 文档完整性
        - [ ] 向后兼容性
        - [ ] 错误处理完善