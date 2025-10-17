# Makefile for Excel MCP Server Testing
# 提供便捷的测试命令

.PHONY: help test test-unit test-integration test-performance test-all
.PHONY: coverage clean lint format install-deps check-deps

# 默认目标
help:
	@echo "Available commands:"
	@echo "  test              - Run all tests"
	@echo "  test-unit         - Run unit tests only"
	@echo "  test-integration  - Run integration tests only"
	@echo "  test-performance   - Run performance tests only"
	@echo "  test-all          - Run all tests with coverage"
	@echo "  coverage          - Run tests with coverage report"
	@echo "  test-parallel     - Run tests in parallel"
	@echo "  test-slow         - Run slow tests"
	@echo "  test-failed       - Run failed tests only"
	@echo "  clean             - Clean test artifacts"
	@echo "  lint              - Run code linting"
	@echo "  format            - Format code"
	@echo "  install-deps      - Install dependencies"
	@echo "  check-deps       - Check dependency issues"

# 测试命令
test:
	python -m pytest tests/ -v

test-unit:
	python -m pytest tests/ -v -m "not integration and not performance"

test-integration:
	python -m pytest tests/ -v -m integration

test-performance:
	python -m pytest tests/ -v -s -m performance

test-all:
	python scripts/run_tests_enhanced.py all

coverage:
	python scripts/run_tests_enhanced.py coverage

test-parallel:
	python scripts/run_tests_enhanced.py parallel

test-slow:
	python scripts/run_tests_enhanced.py slow

test-failed:
	python scripts/run_tests_enhanced.py failed

# 代码质量
lint:
	@echo "Running code linting..."
	flake8 src/ tests/ --count --select=E9,F63,F7,F82 --show-source --statistics
	@echo "Running type checking..."
	mypy src/ --ignore-missing-imports

format:
	@echo "Formatting code..."
	black src/ tests/ scripts/
	isort src/ tests/ scripts/

format-check:
	@echo "Checking code formatting..."
	black --check src/ tests/ scripts/
	isort --check-only src/ tests/ scripts/

# 依赖管理
install-deps:
	@echo "Installing development dependencies..."
	pip install -e ".[dev,performance,quality]"

check-deps:
	@echo "Checking for dependency issues..."
	pip check

# 清理
clean:
	@echo "Cleaning test artifacts..."
	rm -rf .pytest_cache/
	rm -rf htmlcov/
	rm -f .coverage
	rm -f coverage.xml
	rm -rf .benchmarks/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -name "*.pyc" -delete
	find . -name "*.pyo" -delete

# 快速检查
check: format-check lint test-unit

# 完整测试套件
full-check: clean install-deps format-check lint test-all

# 性能基准
benchmark:
	python scripts/run_tests_enhanced.py benchmark

# 开发环境设置
dev-setup: install-deps
	@echo "Setting up development environment..."
	@if [ ! -f .git/hooks/pre-commit ]; then \
		echo "Setting up pre-commit hooks..."; \
		pre-commit install; \
	fi

# 持续集成支持
ci: clean install-deps format-check lint test-all coverage
	@echo "CI pipeline completed successfully"

# 开发者常用命令
dev: test-unit
	@echo "Running unit tests (development mode)"

# 发布前检查
pre-release: clean format-check lint test-all coverage
	@echo "Pre-release checks completed"

# 监控测试
watch:
	@echo "Watching for file changes (requires pytest-watch)..."
	pip install pytest-watch
	ptw tests/ --runner "python -m pytest {tests}" -- -v

# 调试测试
debug:
	python -m pytest tests/ -v -s --pdb