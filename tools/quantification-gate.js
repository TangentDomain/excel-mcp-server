#!/usr/bin/env node
// tools/quantification-gate.js — 量化指标退化门禁
// 在 git commit 前跑对抗评分，与上次基线对比，任何指标退化 >2% 则拒绝提交

"use strict";

const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");

const REPO_ROOT = path.resolve(__dirname, "..");
const JSONL_PATH = path.join(REPO_ROOT, "data", "adversarial-score.jsonl");
const SCRIPT_PATH = path.join(REPO_ROOT, "scripts", "adversarial-score.py");

const DEGRADATION_THRESHOLD = 0.02; // 2%

// 维度配置：key path → 中文名
const DIMENSIONS = [
  { path: "accuracy", field: "accuracy_pct", label: "准确率" },
  { path: "tool_coverage", field: "coverage_pct", label: "工具覆盖率" },
  { path: "sql_coverage", field: "coverage_pct", label: "SQL 特性覆盖率" },
  { path: "edge_coverage", field: "coverage_pct", label: "边界值覆盖率" },
  { path: "write_safety", field: null, label: "写操作安全性" }, // 聚合计算
];

/**
 * 从 jsonl 文件读最后一行 JSON
 * @param {string} filePath
 * @returns {object|null}
 */
function readLastJsonlLine(filePath) {
  if (!fs.existsSync(filePath)) return null;
  const content = fs.readFileSync(filePath, "utf-8");
  const lines = content.split("\n").filter(Boolean);
  if (lines.length === 0) return null;
  const lastLine = lines[lines.length - 1];
  try {
    return JSON.parse(lastLine);
  } catch {
    return null;
  }
}

/**
 * 从 write_safety 子维度聚合一个 overall 百分比
 * 取 4 个子维度 passed/total 的加权平均
 */
function computeWriteSafetyOverall(ws) {
  if (!ws) return null;
  const subKeys = [
    "affected_rows_accuracy",
    "readback_consistency",
    "no_match_safety",
    "row_count_consistency",
  ];
  let totalPassed = 0;
  let totalAll = 0;
  for (const key of subKeys) {
    const sub = ws[key];
    if (sub && sub.total > 0) {
      totalPassed += sub.passed;
      totalAll += sub.total;
    }
  }
  if (totalAll === 0) return null;
  return Math.round((totalPassed / totalAll) * 1000) / 10;
}

/**
 * 提取某个维度的 overall 值（百分比）
 * @param {object} report
 * @param {{ path: string, field: string|null, label: string }} dim
 * @returns {number|null}
 */
function getDimensionValue(report, dim) {
  const section = report[dim.path];
  if (!section) return null;
  if (dim.field) return section[dim.field] ?? null;
  if (dim.path === "write_safety") return computeWriteSafetyOverall(section);
  return null;
}

// ========== 主流程 ==========

// 1. 读取上次基线（脚本运行前的最后一行）
const baseline = readLastJsonlLine(JSONL_PATH);

// 2. 运行对抗评分脚本（脚本自身会追加新行到 jsonl）
console.log("🔍 运行对抗评分脚本...");
let scriptOutput;
try {
  scriptOutput = execSync(`python "${SCRIPT_PATH}"`, {
    cwd: REPO_ROOT,
    encoding: "utf-8",
    windowsHide: true,
    timeout: 60000,
    stdio: ["pipe", "pipe", "pipe"],
  });
} catch (e) {
  // 脚本 exit code != 0（有失败用例），但仍然可能有评分输出
  // 脚本失败时 stderr 可能有信息，stdout 也有报告
  console.log(e.stdout || "");
  if (e.stderr) console.error(e.stderr);
  // 不阻断 — 退化检测看 jsonl 新行即可
  console.log("⚠️ 对抗评分脚本有失败用例，继续退化检测...");
}

// 3. 读取新评分（脚本追加后的最后一行）
const current = readLastJsonlLine(JSONL_PATH);
if (!current) {
  console.log("✅ 无历史基线，首次运行，跳过退化检测");
  process.exit(0);
}

// 4. 无基线时跳过对比
if (!baseline) {
  console.log("✅ 无历史基线，首次运行，跳过退化检测");
  process.exit(0);
}

// 5. 对比各维度
const regressions = [];
for (const dim of DIMENSIONS) {
  const baselineVal = getDimensionValue(baseline, dim);
  const currentVal = getDimensionValue(current, dim);
  if (baselineVal === null || currentVal === null) continue;
  const diff = baselineVal - currentVal;
  if (diff > DEGRADATION_THRESHOLD) {
    regressions.push({
      label: dim.label,
      baseline: baselineVal,
      current: currentVal,
      diff: Math.round(diff * 10) / 10,
    });
  }
}

if (regressions.length > 0) {
  console.error("\n❌ 量化指标退化检测未通过（退化阈值 >2%）：\n");
  for (const r of regressions) {
    console.error(
      `  ⚠️ ${r.label}: ${r.baseline}% → ${r.current}%（下降 ${r.diff}%）`
    );
  }
  console.error("");
  process.exit(1);
}

console.log("✅ 量化指标退化检测通过");
process.exit(0);
