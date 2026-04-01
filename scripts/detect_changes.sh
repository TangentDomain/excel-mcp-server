#!/usr/bin/env bash
# CI分层策略：根据变更内容决定测试范围
# 输出：TEST_TYPE=full/targeted/none

set -e

# Get changed files since last commit (or between branches)
if [ -n "$CI_MERGE_REQUEST_TARGET_BRANCH_NAME" ]; then
    # GitLab CI
    CHANGED_FILES=$(git diff --name-only origin/$CI_MERGE_REQUEST_TARGET_BRANCH_NAME...HEAD)
elif [ -n "$GITHUB_BASE_REF" ]; then
    # GitHub Actions PR
    git fetch origin $GITHUB_BASE_REF
    CHANGED_FILES=$(git diff --name-only origin/$GITHUB_BASE_REF...HEAD)
else
    # Get changed files in last commit
    CHANGED_FILES=$(git diff --name-only HEAD~1 HEAD)
fi

echo "🔍 Changed files:"
echo "$CHANGED_FILES"
echo ""

# Detect what changed
SRC_CHANGED=0
TEST_CHANGED=0
DOCS_CHANGED=0

for file in $CHANGED_FILES; do
    if [[ $file == src/* ]]; then
        SRC_CHANGED=1
    elif [[ $file == tests/* ]]; then
        TEST_CHANGED=1
    elif [[ $file == docs/* ]] || [[ $file == *.md ]]; then
        DOCS_CHANGED=1
    elif [[ $file == scripts/* ]] || [[ $file == .github/* ]] || [[ $file == pyproject.toml ]] || [[ $file == pytest.ini ]]; then
        # Build/CI changes also require full tests
        SRC_CHANGED=1
    fi
done

# Determine test strategy
if [ $SRC_CHANGED -eq 1 ]; then
    echo "📊 src/ modified → Running full tests"
    echo "TEST_TYPE=full"
    exit 0
elif [ $TEST_CHANGED -eq 1 ] && [ $SRC_CHANGED -eq 0 ]; then
    echo "📊 tests/ modified only → Running targeted tests"
    echo "TEST_TYPE=targeted"

    # Find which test files changed
    CHANGED_TEST_FILES=$(echo "$CHANGED_FILES" | grep "^tests/" || true)

    if [ -z "$CHANGED_TEST_FILES" ]; then
        # No test files, just test files were deleted/moved
        echo "TARGET_TESTS=tests/"
    else
        echo "TARGET_TESTS=$CHANGED_TEST_FILES"
    fi
    exit 0
elif [ $DOCS_CHANGED -eq 1 ] && [ $SRC_CHANGED -eq 0 ] && [ $TEST_CHANGED -eq 0 ]; then
    echo "📊 docs/ modified only → Skipping tests"
    echo "TEST_TYPE=none"
    exit 0
else
    # Default: run full tests
    echo "📊 Default strategy → Running full tests"
    echo "TEST_TYPE=full"
    exit 0
fi
