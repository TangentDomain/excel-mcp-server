<div align="center">
<a href="CONTRIBUTING.md">English</a> | <a href="CONTRIBUTING.zh-CN.md">简体中文</a>
</div>

# Contributing to ExcelMCP

First off, thank you for considering contributing to ExcelMCP! It's people like you that make this such a great tool.

## Where do I go from here?

If you've noticed a bug or have a feature request, [make one](https://github.com/tangjian/excel-mcp-server/issues/new)! It's generally best if you get confirmation of your bug or approval for your feature request this way before starting to code.

### Fork & create a branch

If this is something you think you can fix, then [fork ExcelMCP](https://github.com/tangjian/excel-mcp-server/fork) and create a branch with a descriptive name.

A good branch name would be (where issue #38 is the ticket you're working on):

```sh
git checkout -b 38-add-awesome-new-feature
```

### Get the test suite running

Make sure you're running the test suite locally. We have a 100% test coverage policy, so all contributions need to be tested.

```sh
# Install development dependencies
uv sync --dev

# Run tests
pytest tests/
```

### Implement your fix or feature

At this point, you're ready to make your changes! Feel free to ask for help; everyone is a beginner at first 😸

### Make a Pull Request

At this point, you should switch back to your master branch and make sure it's up to date with ExcelMCP's master branch:

```sh
git remote add upstream git@github.com:tangjian/excel-mcp-server.git
git checkout master
git pull upstream master
```

Then update your feature branch from your local copy of master, and push it!

```sh
git checkout 38-add-awesome-new-feature
git rebase master
git push --force-with-lease origin 38-add-awesome-new-feature
```

Finally, go to GitHub and [make a Pull Request](https://github.com/tangjian/excel-mcp-server/compare)

### Keeping your Pull Request updated

If a maintainer asks you to "rebase" your PR, they're saying that a lot of code has changed, and that you need to update your branch so it's easier to merge.

To learn more about rebasing and merging, check out this guide on [syncing a fork](https://help.github.com/articles/syncing-a-fork).

We're happy to help you get your PR ready for merging.

Thanks for contributing!
