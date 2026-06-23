#!/usr/bin/env node
// tools/harness-l5/protected-evaluator.js
// 受保护评估器 — 退化检测（封闭枚举，见 reference.md 8.5 节）

"use strict";

const { readFileSync } = require("fs");
const { resolve, join } = require("path");

// ── 受保护评估器 ID 列表 ──
const PROTECTED_EVALUATORS = new Set([
  "ruff-check",
  "ruff-format",
  "invariant-tests",
]);

// ── 受保护文件路径（相对于项目根） ──
const PROTECTED_FILES = new Set([
  "data/evaluator-registry.json",
  "tools/harness-l5/protected-evaluator.js",
  "tools/check-harness-l5.js",
  "tools/harness-l5/evolve.js",
  "data/trust-anchor.sha256",
  "data/carrier-registry.json",
  ".githooks/pre-commit",
]);

// ── 退化检测封闭枚举（8.5 节） ──
// 1. 放宽 pyproject.toml 中 ruff 配置（本项目无 tsconfig）
// 2. 删/改 tests/invariants/ 中的 assert 阈值
// 3. 删 data/carrier-registry.json 中的条目
// 4. 改 tools/harness-l5/protected-evaluator.js 自身代码
// 5. 改 pre-commit hook 中的 gate/warn_gate 调用顺序或参数

/**
 * 判断文件是否受保护
 * @param {string} file - 文件路径（相对于项目根）
 * @returns {boolean}
 */
function isProtected(file) {
  const normalized = file.replace(/\\/g, "/").replace(/^\.\/?/, "");
  return PROTECTED_FILES.has(normalized);
}

/**
 * 检测变更是否构成退化（放宽判定标准）
 * @param {{ file: string, operation: string, diff?: string, content?: string }[]} changes
 * @returns {{ isDegradation: boolean, reason: string }}
 */
function checkDegradation(changes) {
  for (const change of changes) {
    const { file, operation, diff, content } = change;
    const normalizedFile = file.replace(/\\/g, "/").replace(/^\.\/?/, "");

    // 1. 放宽 pyproject.toml 中 ruff 配置
    if (normalizedFile === "pyproject.toml") {
      const text = content || diff || "";
      if (text.includes("line-length") && /line-length\s*[:=]\s*\d+/.test(text)) {
        const match = text.match(/line-length\s*[:=]\s*(\d+)/);
        if (match && parseInt(match[1], 10) > 88) {
          return { isDegradation: true, reason: `放宽 pyproject.toml line-length: ${match[1]} > 88` };
        }
      }
      if (text.includes("ignore") && /ignore\s*[:=]\s*\[/.test(text)) {
        const ignoreAddPattern = /\+.*ignore.*\[.*"[A-Z]/;
        if (ignoreAddPattern.test(text)) {
          return { isDegradation: true, reason: "向 pyproject.toml ruff ignore 列表新增规则（放宽判定）" };
        }
      }
    }

    // 2. 删/改 tests/invariants/ 中的 assert 阈值
    if (normalizedFile.startsWith("tests/invariants/")) {
      const text = content || diff || "";
      if (operation === "delete" || /^-.*assert/.test(text)) {
        return { isDegradation: true, reason: `删除 tests/invariants/ 中的 assert 断言: ${normalizedFile}` };
      }
      const lines = text.split("\n");
      for (const line of lines) {
        if (line.startsWith("-") && /assert\s+/.test(line) && /==\s*\d+|>=\s*\d+|<=\s*\d+/.test(line)) {
          return { isDegradation: true, reason: `修改 tests/invariants/ 中的 assert 阈值: ${normalizedFile}` };
        }
      }
    }

    // 3. 删 data/carrier-registry.json 中的条目
    if (normalizedFile === "data/carrier-registry.json") {
      const text = content || diff || "";
      if (/^\-\s*\{/.test(text) && text.includes('"id"')) {
        return { isDegradation: true, reason: "删除 data/carrier-registry.json 中的载体条目" };
      }
    }

    // 4. 改 tools/harness-l5/protected-evaluator.js 自身代码
    if (normalizedFile === "tools/harness-l5/protected-evaluator.js") {
      return { isDegradation: true, reason: "修改受保护评估器自身代码（protected-evaluator.js）" };
    }

    // 5. 改 pre-commit hook 中的 gate/warn_gate 调用顺序或参数
    if (normalizedFile === ".githooks/pre-commit") {
      const text = content || diff || "";
      if (/(gate|warn_gate)/.test(text) && /(-|removed|delete)/i.test(text)) {
        return { isDegradation: true, reason: "修改 pre-commit hook 中的 gate/warn_gate 调用" };
      }
    }
  }

  return { isDegradation: false, reason: "" };
}

module.exports = { PROTECTED_EVALUATORS, PROTECTED_FILES, isProtected, checkDegradation };
