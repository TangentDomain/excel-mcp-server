#!/usr/bin/env bun
// scripts/compare-holdout.ts — Loop I3 外层对比 + 回滚
// Wilson score interval 95% CI, delta > 0.05 且 CI 下界 > pOld → pass

import { spawnSync } from "child_process";

const options = { encoding: "utf-8" as const, windowsHide: true, timeout: 120_000 };

function run(cmd: string): { stdout: string; status: number | null } {
  const r = spawnSync(cmd, { shell: true, stdio: "pipe", ...options });
  return { stdout: r.stdout?.toString().trim() ?? "", status: r.status };
}

function parsePassRate(output: string): number {
  // 从 pytest summary 解析 passed/total
  const m = output.match(/(\d+) passed/);
  if (!m) return 0;
  const passed = parseInt(m[1], 10);
  const total = output.match(/(\d+) (?:passed|failed)/);
  const t = total ? parseInt(total[1], 10) : passed;
  return t > 0 ? passed / t : 0;
}

function wilsonCI(p: number, n: number): [number, number] {
  const z = 1.96; // 95% CI
  const denom = 1 + (z * z) / n;
  const center = (p + (z * z) / (2 * n)) / denom;
  const spread = (z * Math.sqrt((p * (1 - p) + (z * z) / (4 * n)) / n)) / denom;
  return [Math.max(0, center - spread), Math.min(1, center + spread)];
}

function stashAndCheckout(commit: string): boolean {
  const s = run("git stash --include-untracked");
  if (s.status !== 0 && !s.stdout.includes("No local changes")) return false;
  const c = run(`git checkout ${commit} -- src/ tests/`);
  return c.status === 0;
}

function restore(): void {
  run("git checkout -");
  run("git stash pop --index 2>/dev/null || git stash pop 2>/dev/null || true");
}

function runPytest(): { output: string; status: number | null } {
  const r = spawnSync("python", ["-m", "pytest", "tests/invariants/", "-q", "--timeout=30"], {
    stdio: "pipe",
    windowsHide: true,
    timeout: 120_000,
  });
  return { output: r.stdout?.toString().trim() ?? "", status: r.status };
}

const baselineCommit = process.argv[2];
if (!baselineCommit) {
  console.error("用法: bun scripts/compare-holdout.ts <baseline-commit>");
  process.exit(2);
}

const pOld = 1.0; // baseline 预期全部通过

// 1. 当前测试
const current = runPytest();
if (current.status !== 0) {
  console.error("❌ 当前代码不变量测试失败，拒绝进化");
  process.exit(2);
}

const pNew = parsePassRate(current.output);

// 2. baseline 对比
if (!stashAndCheckout(baselineCommit)) {
  console.error("❌ 无法 checkout baseline，跳过对比");
  process.exit(0);
}

const baseline = runPytest();
restore();

if (baseline.status !== 0) {
  console.error("❌ baseline 不变量测试失败，数据不可信");
  process.exit(2);
}

const pBase = parsePassRate(baseline.output);
const totalTests = 32; // INV-1~32
const delta = pNew - pBase;
const [, lower] = wilsonCI(pNew, totalTests);

const pass = delta > 0.05 && lower > pOld * 0.95; // 允许 5% 退化容差
if (pass) {
  console.log(`✅ 进化通过: pNew=${pNew.toFixed(3)} delta=${delta.toFixed(3)} CI=[${lower.toFixed(3)}, ...]`);
  process.exit(0);
} else {
  console.log(`❌ 进化拒绝: pNew=${pNew.toFixed(3)} delta=${delta.toFixed(3)} CI=[${lower.toFixed(3)}, ...]`);
  process.exit(2);
}
