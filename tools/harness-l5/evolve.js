#!/usr/bin/env node
// tools/harness-l5/evolve.js
// L5 进化引擎 — 五阶段 CLI + 并发锁
// 严格按 reference.md 8.1-8.9 实现

"use strict";

const { execSync } = require("child_process");
const { createHash } = require("crypto");
const {
  existsSync,
  readFileSync,
  writeFileSync,
  unlinkSync,
  mkdirSync,
  openSync,
  closeSync,
  writeSync,
  readdirSync,
  copyFileSync,
} = require("fs");
const { resolve, join, dirname } = require("path");

// ── 项目根目录 ──
const ROOT = resolve(__dirname, "..", "..");

// ── 路径常量 ──
const PROPOSALS_DIR = join(ROOT, "tools", "harness-l5", "proposals");
const LOCK_FILE = join(PROPOSALS_DIR, ".lock");
const EVOLUTION_LOG = join(ROOT, "data", "evolution-log.json");
const EVALUATOR_REGISTRY = join(ROOT, "data", "evaluator-registry.json");
const CHECK_HARNESS_L5 = join(ROOT, "tools", "check-harness-l5.js");

// ── 受保护评估器模块 ──
const { PROTECTED_EVALUATORS, isProtected, checkDegradation } = require("./protected-evaluator");

// ── 锁 TTL（秒） ──
const LOCK_TTL = 300;

// ── 提案 TTL（秒，30 天） ──
const PROPOSAL_TTL = 2592000;

// ── 受保护文件列表（不可被进化修改，8.7 节 blocked） ──
const IMMUTABLE_FILES = new Set([
  "data/evaluator-registry.json",
  "tools/check-harness-l5.js",
  "tools/harness-l5/evolve.js",
  "data/trust-anchor.sha256",
]);

// ── 需要 human review 的变更模式（8.7 节） ──
const NEEDS_REVIEW_PATTERNS = [
  { pattern: /^\.omp\/rules\/axioms\.md/, reason: "修改公理（AX-）" },
  { pattern: /\.githooks\/pre-commit/, reason: "修改 pre-commit 脚本" },
  { pattern: /^\.omp\/extensions\//, reason: "修改 Hook 代码" },
];

// ═══════════════════════════════════════════════════════════════
// 工具函数
// ═══════════════════════════════════════════════════════════════

/** 执行 shell 命令，windowsHide: true */
function run(cmd, options = {}) {
  return execSync(cmd, {
    cwd: ROOT,
    encoding: "utf-8",
    stdio: ["pipe", "pipe", "pipe"],
    windowsHide: true,
    timeout: 60000,
    ...options,
  }).trim();
}

/** 获取当前 git HEAD hash */
function getHeadHash() {
  try {
    return run("git rev-parse HEAD");
  } catch {
    return null;
  }
}

/** 获取当前分支名 */
function getCurrentBranch() {
  try {
    return run("git symbolic-ref --short HEAD");
  } catch {
    return "main";
  }
}

/** 生成 proposalId: L5P-YYYYMMDD-NNN */
function generateProposalId() {
  const now = new Date();
  const date = now.toISOString().slice(0, 10).replace(/-/g, "");
  const prefix = `L5P-${date}`;
  const existing = listProposals().filter((p) => p.startsWith(prefix));
  let n = existing.length + 1;
  while (existing.includes(`${prefix}-${String(n).padStart(3, "0")}`)) n++;
  return `${prefix}-${String(n).padStart(3, "0")}`;
}

/** 列出 proposals 目录中的所有提案 ID */
function listProposals() {
  if (!existsSync(PROPOSALS_DIR)) return [];
  return readdirSync(PROPOSALS_DIR)
    .filter((f) => f.endsWith(".json"))
    .map((f) => f.replace(/\.json$/, ""));
}

/** 读取提案 */
function readProposal(proposalId) {
  const path = join(PROPOSALS_DIR, `${proposalId}.json`);
  if (!existsSync(path)) {
    console.error(`❌ 提案不存在: ${proposalId}`);
    process.exit(1);
  }
  return JSON.parse(readFileSync(path, "utf-8"));
}

/** 写入提案 */
function writeProposal(proposal) {
  proposal.updatedAt = new Date().toISOString();
  const path = join(PROPOSALS_DIR, `${proposal.id}.json`);
  writeFileSync(path, JSON.stringify(proposal, null, 2) + "\n");
}

/** 读取 evolution log */
function readEvolutionLog() {
  if (!existsSync(EVOLUTION_LOG)) return { log: [] };
  try {
    return JSON.parse(readFileSync(EVOLUTION_LOG, "utf-8"));
  } catch {
    return { log: [] };
  }
}

/** 追加 evolution log */
function appendEvolutionLog(entry) {
  const log = readEvolutionLog();
  log.log.push(entry);
  writeFileSync(EVOLUTION_LOG, JSON.stringify(log, null, 2) + "\n");
}

// ═══════════════════════════════════════════════════════════════
// 并发锁（8.1）
// ═══════════════════════════════════════════════════════════════

function acquireLock() {
  if (!existsSync(PROPOSALS_DIR)) {
    mkdirSync(PROPOSALS_DIR, { recursive: true });
  }

  try {
    const fd = openSync(LOCK_FILE, "wx");
    writeSync(fd, JSON.stringify({ pid: process.pid, timestamp: Date.now() }));
    closeSync(fd);
    return true;
  } catch (e) {
    if (e.code === "EEXIST") {
      try {
        const lockData = JSON.parse(readFileSync(LOCK_FILE, "utf-8"));
        const age = (Date.now() - lockData.timestamp) / 1000;

        if (age > LOCK_TTL) {
          console.error(`⚠️  锁文件 TTL 超时 (${age.toFixed(0)}s > ${LOCK_TTL}s)，强制删除`);
          unlinkSync(LOCK_FILE);
          return acquireLock();
        }

        try {
          process.kill(lockData.pid, 0);
          console.error(`❌ 另一个提案正在处理中 (pid=${lockData.pid})`);
          return false;
        } catch {
          console.error(`⚠️  锁持有进程 (pid=${lockData.pid}) 已不存在，清理 stale lock`);
          unlinkSync(LOCK_FILE);
          return acquireLock();
        }
      } catch {
        unlinkSync(LOCK_FILE);
        return acquireLock();
      }
    }
    throw e;
  }
}

function releaseLock() {
  try {
    if (existsSync(LOCK_FILE)) {
      unlinkSync(LOCK_FILE);
    }
  } catch { /* ignore */ }
}

// ═══════════════════════════════════════════════════════════════
// 变更风险分级（8.7）
// ═══════════════════════════════════════════════════════════════

function assessRisk(changes, type) {
  for (const change of changes) {
    const file = change.file.replace(/\\/g, "/").replace(/^\.\/?/, "");

    // blocked: 修改受保护文件
    if (IMMUTABLE_FILES.has(file)) {
      return { level: "blocked", reason: `受保护文件不可修改: ${file}` };
    }

    // needs_human_review: 修改公理 / pre-commit / hook
    for (const rule of NEEDS_REVIEW_PATTERNS) {
      if (rule.pattern.test(file)) {
        return { level: "needs_human_review", reason: rule.reason };
      }
    }

    // needs_human_review: 修改 Rule 的判定命令
    if (file.startsWith(".omp/rules/") && file.endsWith(".md")) {
      const text = change.diff || change.content || "";
      const lines = text.split("\n");
      for (const line of lines) {
        if (/判定[:\s]/.test(line) && (line.startsWith("+") || line.startsWith("-"))) {
          return { level: "needs_human_review", reason: "修改 Rule 的判定命令" };
        }
      }
    }
  }

  return { level: "auto_commit", reason: "" };
}

// ═══════════════════════════════════════════════════════════════
// 阶段: propose（8.3）
// ═══════════════════════════════════════════════════════════════

function propose(input) {
  const proposalId = generateProposalId();
  const headHash = getHeadHash();

  // 退化检测
  const degradation = checkDegradation(input.changes || []);
  if (degradation.isDegradation) {
    console.error(`❌ 提案被拒绝 — 退化检测: ${degradation.reason}`);
    process.exit(1);
  }

  // 风险分级
  const risk = assessRisk(input.changes || [], input.type);
  if (risk.level === "blocked") {
    console.error(`❌ 提案被拒绝 — ${risk.reason}`);
    process.exit(1);
  }

  const proposal = {
    id: proposalId,
    type: input.type || "rule",
    status: "proposed",
    propose: {
      input: input,
      output: {
        proposalId,
        status: "proposed",
        rollbackTarget: headHash,
      },
      rollbackTarget: headHash,
    },
    evaluate: { output: null },
    commit: { output: null },
    rollback: { output: null },
    review: { output: null },
    riskLevel: risk.level,
    riskReason: risk.reason,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ttl: PROPOSAL_TTL,
  };

  writeProposal(proposal);

  console.log(`✅ 提案已创建: ${proposalId}`);
  console.log(`   回滚目标: ${headHash}`);
  console.log(`   风险等级: ${risk.level}`);

  appendEvolutionLog({
    proposalId,
    action: "propose",
    timestamp: new Date().toISOString(),
    reason: input.reason || "",
    expected: input.expected || "",
    riskLevel: risk.level,
    rollbackTarget: headHash,
  });

  return proposal;
}

// ═══════════════════════════════════════════════════════════════
// 阶段: evaluate（8.3）
// ═══════════════════════════════════════════════════════════════

function evaluate(proposalId) {
  const proposal = readProposal(proposalId);

  if (!["proposed", "eval_failed", "pending_review"].includes(proposal.status)) {
    console.error(`❌ 提案 ${proposalId} 状态不可 evaluate (当前: ${proposal.status})`);
    process.exit(1);
  }

  proposal.status = "evaluating";
  writeProposal(proposal);

  const failures = [];

  // 1. 信任锚校验
  try {
    run(`node "${CHECK_HARNESS_L5}"`, { timeout: 30000 });
    console.log("  ✅ 信任锚校验通过");
  } catch (e) {
    const output = (e.stdout || "") + (e.stderr || "");
    failures.push({ guard: "trust-anchor", output });
    proposal.status = "trust_anchor_broken";
    proposal.evaluate.output = {
      passed: false,
      failures,
      needsHumanReview: false,
      trustAnchorBroken: true,
    };
    writeProposal(proposal);
    console.error("❌ 信任锚校验失败！");
    return proposal;
  }

  // 2. 运行评估器（仅运行与变更文件相关的评估器）
  let evaluators = [];
  try {
    evaluators = JSON.parse(readFileSync(EVALUATOR_REGISTRY, "utf-8"));
  } catch {
    console.error("❌ 无法读取 evaluator-registry.json");
    proposal.status = "eval_failed";
    writeProposal(proposal);
    return proposal;
  }

  for (const ev of evaluators) {
    const changes = proposal.propose.input.changes || [];
    const relevant = changes.some((c) => {
      const f = c.file.replace(/\\/g, "/");
      if ((ev.id === "ruff-check" || ev.id === "ruff-format") && f.endsWith(".py")) return true;
      if (ev.id === "invariant-tests" && f.startsWith("tests/invariants/")) return true;
      if (ev.id === "docstring-contract" && f.startsWith("src/")) return true;
      return false;
    });

    if (!relevant) {
      console.log(`  ⏭️  跳过不相关评估器: ${ev.id}`);
      continue;
    }

    try {
      console.log(`  🔍 运行评估器: ${ev.id} (${ev.command})`);
      run(ev.command, { timeout: 120000, cwd: ROOT });
      console.log(`  ✅ ${ev.id} 通过`);
    } catch (e) {
      const output = (e.stdout || "") + (e.stderr || "");
      failures.push({ guard: ev.id, output: output.substring(0, 500) });
      console.error(`  ❌ ${ev.id} 失败`);
    }
  }

  const hasProtectedFailure = failures.some((f) => {
    return evaluators.some((ev) => ev.id === f.guard && ev.protected);
  });

  const passed = failures.length === 0;
  const needsHumanReview = proposal.riskLevel === "needs_human_review" || hasProtectedFailure;

  proposal.evaluate.output = { passed, failures, needsHumanReview };

  if (!passed) {
    proposal.status = "eval_failed";
    console.error("❌ 评估未通过:");
    for (const f of failures) {
      console.error(`  - ${f.guard}: ${f.output.substring(0, 200)}`);
    }
  } else if (needsHumanReview) {
    proposal.status = "pending_review";
    console.log("✅ 评估通过，但需要人工审核");
  } else {
    proposal.status = "eval_passed";
    console.log("✅ 评估通过，可自动提交");
  }

  writeProposal(proposal);

  appendEvolutionLog({
    proposalId,
    action: "evaluate",
    timestamp: new Date().toISOString(),
    passed,
    failures: failures.map((f) => f.guard),
    needsHumanReview,
  });

  return proposal;
}

// ═══════════════════════════════════════════════════════════════
// 阶段: commit（8.3 — 原子性）
// ═══════════════════════════════════════════════════════════════

function commitProposal(proposalId) {
  const proposal = readProposal(proposalId);

  if (proposal.status !== "eval_passed") {
    console.error(`❌ 提案 ${proposalId} 状态不可 commit (当前: ${proposal.status})`);
    process.exit(1);
  }

  const branch = getCurrentBranch();
  const tmpBranch = "evolve-tmp";
  const changes = proposal.propose.input.changes || [];

  // 保存提案 JSON 内容到临时位置（branch switch 会丢失工作区文件）
  const proposalJson = JSON.stringify(proposal, null, 2) + "\n";
  const proposalBackupPath = join(ROOT, "tools", "harness-l5", "proposals", `${proposalId}.json`);

  try {
    proposal.status = "committing";
    writeProposal(proposal); // 确保状态已写

    // 1. 创建临时分支
    console.log(`  分支: ${branch} → ${tmpBranch}`);
    run(`git checkout -b ${tmpBranch}`);

    // 2. 逐文件应用变更
    for (const change of changes) {
      const filePath = join(ROOT, change.file);
      if (change.operation === "text") {
        if (change.content !== undefined) {
          const dir = dirname(filePath);
          if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
          writeFileSync(filePath, change.content);
          console.log(`  📝 应用变更: ${change.file} (完整内容)`);
        } else if (change.diff) {
          applyDiff(change.file, change.diff);
          console.log(`  📝 应用变更: ${change.file} (diff)`);
        }
      } else if (change.operation === "binary") {
        console.log(`  📦 跳过二进制变更: ${change.file}`);
      }
    }

    // 3. 暂存并提交（--no-verify: evolve 引擎自身已运行评估器，不重复跑 pre-commit gate）
    run("git add -A");
    const commitMsg = `evolve: ${proposalId} ${proposal.propose.input.reason || ""}`.trim();
    const commitHash = run(`git commit --no-verify -m ${JSON.stringify(commitMsg)}`);
    console.log(`  ✅ 提交: ${commitHash.substring(0, 8)}`);

    // 4. 合并回原分支
    run(`git checkout ${branch}`);
    run(`git merge --ff-only ${tmpBranch}`);
    console.log(`  ✅ 合并回 ${branch}`);

    // 5. 清理临时分支
    run(`git branch -d ${tmpBranch}`);

    // 6. 恢复提案 JSON（合并后工作区可能被覆盖）
    if (!existsSync(proposalBackupPath)) {
      writeFileSync(proposalBackupPath, proposalJson);
    }

    const result = {
      committed: true,
      commitHash: commitHash.substring(0, 40),
      rollbackTarget: proposal.propose.rollbackTarget,
    };

    proposal.status = "committed";
    proposal.commit.output = result;
    writeProposal(proposal);

    appendEvolutionLog({
      proposalId,
      action: "commit",
      timestamp: new Date().toISOString(),
      commitHash: result.commitHash,
      rollbackTarget: result.rollbackTarget,
    });

    console.log(`✅ 提案已提交: ${proposalId}`);
    console.log(`   Commit: ${result.commitHash.substring(0, 8)}`);
    console.log(`   回滚目标: ${result.rollbackTarget.substring(0, 8)}`);

    return proposal;
  } catch (e) {
    console.error(`❌ 提交失败: ${e.message || e}`);

    // 回滚：切回原分支，删除临时分支
    // 注意：不执行 git reset --hard，避免破坏工作区中的其他未提交文件
    try { run(`git checkout ${branch}`, { stdio: "pipe" }); } catch { /* ignore */ }
    try { run(`git branch -D ${tmpBranch}`, { stdio: "pipe" }); } catch { /* ignore */ }

    // 恢复提案 JSON
    if (!existsSync(proposalBackupPath)) {
      writeFileSync(proposalBackupPath, proposalJson);
    }

    proposal.status = "eval_passed";
    writeProposal(proposal);

    appendEvolutionLog({
      proposalId,
      action: "commit_failed",
      timestamp: new Date().toISOString(),
      error: String(e.message || e),
    });

    process.exit(1);
  }
}

// ═══════════════════════════════════════════════════════════════
// 阶段: review（8.3）
// ═══════════════════════════════════════════════════════════════

function reviewProposal(proposalId, reviewer, decision) {
  const proposal = readProposal(proposalId);

  if (proposal.status !== "pending_review") {
    console.error(`❌ 提案 ${proposalId} 状态不可 review (当前: ${proposal.status})`);
    process.exit(1);
  }

  if (decision === "approved") {
    proposal.status = "eval_passed";
    proposal.review = { reviewer, decision, timestamp: new Date().toISOString() };
    writeProposal(proposal);
    appendEvolutionLog({ proposalId, action: "review_approved", timestamp: new Date().toISOString(), reviewer });
    console.log(`✅ 提案已批准: ${proposalId} (reviewer: ${reviewer})`);
    return proposal;
  } else {
    proposal.status = "draft";
    proposal.review = { reviewer, decision, timestamp: new Date().toISOString() };
    writeProposal(proposal);
    appendEvolutionLog({ proposalId, action: "review_rejected", timestamp: new Date().toISOString(), reviewer });
    console.log(`❌ 提案已拒绝: ${proposalId} (reviewer: ${reviewer})`);
    return proposal;
  }
}

// ═══════════════════════════════════════════════════════════════
// 阶段: rollback（8.3）
// ═══════════════════════════════════════════════════════════════

function rollbackProposal(proposalId) {
  const proposal = readProposal(proposalId);

  if (proposal.status !== "committed") {
    console.error(`❌ 提案 ${proposalId} 状态不可 rollback (当前: ${proposal.status})`);
    process.exit(1);
  }

  const target = proposal.commit.output?.rollbackTarget;
  if (!target) {
    console.error("❌ 提案无回滚目标");
    process.exit(1);
  }

  try {
    proposal.status = "rolling_back";
    writeProposal(proposal);
    run(`git reset --hard ${target}`);
    console.log(`✅ 已回滚到: ${target.substring(0, 8)}`);

    proposal.status = "rolled_back";
    proposal.rollback.output = { rolledBack: true, restoredHash: target };
    writeProposal(proposal);

    appendEvolutionLog({ proposalId, action: "rollback", timestamp: new Date().toISOString(), restoredHash: target });
    console.log(`✅ 提案已回滚: ${proposalId}`);
    return proposal;
  } catch (e) {
    proposal.status = "committed";
    writeProposal(proposal);
    appendEvolutionLog({ proposalId, action: "rollback_failed", timestamp: new Date().toISOString(), error: String(e.message || e) });
    console.error(`❌ 回滚失败: ${e.message || e}`);
    process.exit(1);
  }
}

// ═══════════════════════════════════════════════════════════════
// Diff 应用工具
// ═══════════════════════════════════════════════════════════════

function applyDiff(file, diff) {
  const filePath = join(ROOT, file);
  if (!existsSync(filePath)) {
    const lines = diff.split("\n");
    const content = lines
      .filter((l) => l.startsWith("+") && !l.startsWith("+++"))
      .map((l) => l.substring(1))
      .join("\n");
    const dir = dirname(filePath);
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    writeFileSync(filePath, content + "\n");
    return;
  }

  const original = readFileSync(filePath, "utf-8");
  const origLines = original.split("\n");
  const diffLines = diff.split("\n");
  const result = [];
  let oi = 0;

  for (let i = 0; i < diffLines.length; i++) {
    const line = diffLines[i];
    if (line.startsWith("---") || line.startsWith("+++") || line.startsWith("@@")) continue;
    if (line.startsWith("-")) {
      oi++;
    } else if (line.startsWith("+")) {
      result.push(line.substring(1));
    } else {
      if (oi < origLines.length) result.push(origLines[oi]);
      oi++;
    }
  }
  while (oi < origLines.length) {
    result.push(origLines[oi]);
    oi++;
  }
  writeFileSync(filePath, result.join("\n"));
}

// ═══════════════════════════════════════════════════════════════
// CLI 入口
// ═══════════════════════════════════════════════════════════════

function parseArgs() {
  const args = process.argv.slice(2);
  const command = args[0];
  const parsed = { command, params: {} };
  let i = 1;
  while (i < args.length) {
    if (args[i] === "--input" && i + 1 < args.length) parsed.params.input = args[++i];
    else if (args[i] === "--proposalId" && i + 1 < args.length) parsed.params.proposalId = args[++i];
    else if (args[i] === "--reviewer" && i + 1 < args.length) parsed.params.reviewer = args[++i];
    else if (args[i] === "--decision" && i + 1 < args.length) parsed.params.decision = args[++i];
    i++;
  }
  return parsed;
}

function main() {
  // 8.8: TRUST_ANCHOR_SKIP 检测
  if (process.env.TRUST_ANCHOR_SKIP === "1") {
    console.error("❌ TRUST_ANCHOR_SKIP=1 环境变量存在，evolve.js 拒绝运行");
    process.exit(1);
  }

  const { command, params } = parseArgs();

  if (!command || !["propose", "evaluate", "commit", "rollback", "review"].includes(command)) {
    console.error("用法:");
    console.error("  evolve.js propose  --input @proposal.json");
    console.error("  evolve.js evaluate --proposalId L5P-YYYYMMDD-NNN");
    console.error("  evolve.js commit   --proposalId L5P-YYYYMMDD-NNN");
    console.error("  evolve.js rollback --proposalId L5P-YYYYMMDD-NNN");
    console.error("  evolve.js review   --proposalId L5P-YYYYMMDD-NNN --reviewer <name> --decision <approved|rejected>");
    process.exit(1);
  }

  if (command === "propose") {
    let inputFile = params.input;
    if (!inputFile) { console.error("❌ --input 参数必需"); process.exit(1); }
    if (inputFile.startsWith("@")) inputFile = inputFile.substring(1);
    const inputPath = resolve(ROOT, inputFile);
    if (!existsSync(inputPath)) { console.error(`❌ 输入文件不存在: ${inputFile}`); process.exit(1); }
    const input = JSON.parse(readFileSync(inputPath, "utf-8"));

    if (!acquireLock()) process.exit(1);
    try {
      const proposal = propose(input);
      console.log("\n🔄 自动流转: propose → evaluate");
      const evaluated = evaluate(proposal.id);
      if (evaluated.status === "eval_passed") {
        console.log("\n🔄 自动流转: evaluate → commit");
        commitProposal(proposal.id);
      } else if (evaluated.status === "pending_review") {
        console.log(`\n⏸️  提案需要人工审核 (pending_review)`);
        console.log(`   运行: evolve.js review --proposalId ${proposal.id} --reviewer <name> --decision approved`);
      } else if (evaluated.status === "eval_failed") {
        console.error("\n❌ 评估未通过，提案状态: eval_failed");
        process.exit(1);
      }
    } finally {
      releaseLock();
    }
  } else if (command === "evaluate") {
    if (!params.proposalId) { console.error("❌ --proposalId 参数必需"); process.exit(1); }
    if (!acquireLock()) process.exit(1);
    try {
      const evaluated = evaluate(params.proposalId);
      if (evaluated.status === "eval_passed") {
        console.log("\n🔄 自动流转: evaluate → commit");
        commitProposal(params.proposalId);
      }
    } finally { releaseLock(); }
  } else if (command === "commit") {
    if (!params.proposalId) { console.error("❌ --proposalId 参数必需"); process.exit(1); }
    if (!acquireLock()) process.exit(1);
    try { commitProposal(params.proposalId); } finally { releaseLock(); }
  } else if (command === "review") {
    if (!params.proposalId || !params.reviewer || !params.decision) {
      console.error("❌ --proposalId, --reviewer, --decision 参数必需"); process.exit(1);
    }
    reviewProposal(params.proposalId, params.reviewer, params.decision);
  } else if (command === "rollback") {
    if (!params.proposalId) { console.error("❌ --proposalId 参数必需"); process.exit(1); }
    if (!acquireLock()) process.exit(1);
    try { rollbackProposal(params.proposalId); } finally { releaseLock(); }
  }
}

main();
