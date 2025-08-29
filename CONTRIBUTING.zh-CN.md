# 贡献指南

首先，感谢您考虑为 ExcelMCP 做出贡献！正是因为有像您这样的人，这个工具才会变得如此出色。

<div align="center">
<a href="CONTRIBUTING.md">English</a> | <a href="CONTRIBUTING.zh-CN.md">简体中文</a>
</div>

## 我该从哪里开始？

如果您发现了错误或有功能请求，请[创建一个 Issue](https://github.com/tangjian/excel-mcp-server/issues/new)！在开始编码之前，最好通过这种方式确认您的错误或获得功能请求的批准。

### Fork 并创建分支

如果您认为自己可以解决问题，请 [Fork ExcelMCP](https://github.com/tangjian/excel-mcp-server/fork) 并创建一个具有描述性名称的分支。

一个好的分支名称应该是（其中 issue #38 是您正在处理的工单）：

```sh
git checkout -b 38-add-awesome-new-feature
```

### 运行测试套件

请确保您可以在本地运行测试套件。我们有 100% 的测试覆盖率政策，因此所有贡献都需要经过测试。

```sh
# 安装开发依赖
uv sync --dev

# 运行测试
pytest tests/
```

### 实现您的修复或功能

现在，您可以开始进行更改了！随时寻求帮助；每个人都是从初学者开始的 😸

### 创建一个拉取请求 (Pull Request)

此时，您应该切换回您的 master 分支，并确保它与 ExcelMCP 的 master 分支保持同步：

```sh
git remote add upstream git@github.com:tangjian/excel-mcp-server.git
git checkout master
git pull upstream master
```

然后从您本地的 master 更新您的功能分支，并将其推送！

```sh
git checkout 38-add-awesome-new-feature
git rebase master
git push --force-with-lease origin 38-add-awesome-new-feature
```

最后，前往 GitHub 并[创建一个拉取请求](https://github.com/tangjian/excel-mcp-server/compare)。

### 保持您的拉取请求更新

如果维护者要求您“rebase”您的 PR，这意味着代码发生了很多变化，您需要更新您的分支以便于合并。

要了解有关 rebase 和 merge 的更多信息，请查看这篇关于[同步 fork](https://help.github.com/articles/syncing-a-fork)的指南。

我们很乐意帮助您准备好您的 PR 以便合并。

感谢您的贡献！
