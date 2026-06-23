#!/usr/bin/env node
// tools/check-harness-l5.js
// L5 自举检查 + 信任锚校验
// exit 0 = L5 健康, exit 2 = 信任锚损坏, exit 1 = 其他错误

"use strict";

const { createHash } = require("crypto");
const { existsSync, readFileSync } = require("fs");
const { resolve, join } = require("path");

const ROOT = resolve(__dirname, "..");

// ── L5 必需文件清单 ──
const REQUIRED_FILES = [
  "tools/harness-l5/evolve.js",
  "tools/harness-l5/protected-evaluator.js",
  "data/evaluator-registry.json",
  "data/trust-anchor.sha256",
  "data/carrier-registry.json",
];

// ── 信任锚包含的文件 ──
const TRUST_ANCHOR_FILES = [
  "data/evaluator-registry.json",
  "tools/harness-l5/protected-evaluator.js",
  "tools/check-harness-l5.js",
];

/**
 * 计算文件的 SHA256
 */
function sha256File(filePath) {
  const data = readFileSync(filePath);
  return createHash("sha256").update(data).digest("hex");
}

/**
 * 校验信任锚
 */
function verifyTrustAnchor() {
  const anchorPath = join(ROOT, "data", "trust-anchor.sha256");

  if (!existsSync(anchorPath)) {
    return { valid: false, errors: ["信任锚文件不存在: data/trust-anchor.sha256"] };
  }

  const anchorContent = readFileSync(anchorPath, "utf-8").trim();
  const expected = new Map();
  for (const line of anchorContent.split("\n")) {
    const match = line.match(/^([a-fA-F0-9]{64})\s{2}(.+)$/);
    if (match) {
      expected.set(match[2].trim(), match[1].toLowerCase());
    }
  }

  const errors = [];
  for (const relPath of TRUST_ANCHOR_FILES) {
    const absPath = join(ROOT, relPath);
    if (!existsSync(absPath)) {
      errors.push(`信任锚引用的文件不存在: ${relPath}`);
      continue;
    }
    const actual = sha256File(absPath);
    const exp = expected.get(relPath);
    if (!exp) {
      errors.push(`信任锚中缺少文件条目: ${relPath}`);
    } else if (actual !== exp) {
      errors.push(`SHA256 不匹配: ${relPath} (期望: ${exp.substring(0, 12)}..., 实际: ${actual.substring(0, 12)}...)`);
    }
  }

  return { valid: errors.length === 0, errors };
}

/**
 * 检查 L5 必需文件存在性
 */
function checkRequiredFiles() {
  const missing = [];
  for (const relPath of REQUIRED_FILES) {
    if (!existsSync(join(ROOT, relPath))) {
      missing.push(relPath);
    }
  }
  return { ok: missing.length === 0, missing };
}

/**
 * 检查 carrier-registry 完整性
 */
function checkCarrierRegistry() {
  const regPath = join(ROOT, "data", "carrier-registry.json");
  if (!existsSync(regPath)) return { ok: true, issues: ["carrier-registry.json 不存在，跳过检查"] };

  try {
    const data = JSON.parse(readFileSync(regPath, "utf-8"));
    const issues = [];
    for (const entry of data) {
      const isIdentityOrIndex =
        entry.variant === "identity" ||
        entry.variant === "index" ||
        entry.type === "memory";

      if (!isIdentityOrIndex && (!entry.axioms || entry.axioms.length === 0)) {
        issues.push(`载体 ${entry.id} 缺少 axioms 字段`);
      }
      if (!entry.id || !entry.path || !entry.type) {
        issues.push(`载体缺少必要字段 (id/path/type): ${JSON.stringify(entry)}`);
      }
    }
    return { ok: issues.length === 0, issues };
  } catch (e) {
    return { ok: false, issues: [`carrier-registry.json 解析失败: ${e.message}`] };
  }
}

function main() {
  const skipTrustAnchor = process.env.TRUST_ANCHOR_SKIP === "1";

  // 1. 必需文件检查
  const files = checkRequiredFiles();
  if (!files.ok) {
    console.error(`❌ L5 必需文件缺失:\n${files.missing.map(f => `  - ${f}`).join("\n")}`);
    process.exit(1);
  }

  // 2. 信任锚校验
  if (skipTrustAnchor) {
    process.stderr.write(
      "\n" +
      "╔══════════════════════════════════════════════════════════════╗\n" +
      "║  ⚠️  TRUST_ANCHOR_SKIP=1 — 信任锚校验已绕过！              ║\n" +
      "║  仅限本地调试使用。生产环境必须修复信任锚。                    ║\n" +
      "╚══════════════════════════════════════════════════════════════╝\n\n"
    );
  } else {
    const anchor = verifyTrustAnchor();
    if (!anchor.valid) {
      console.error("❌ 信任锚校验失败:");
      for (const err of anchor.errors) {
        console.error(`  - ${err}`);
      }
      process.exit(2);
    }
    console.log("✅ 信任锚校验通过");
  }

  // 3. carrier-registry 完整性检查
  const registry = checkCarrierRegistry();
  if (!registry.ok) {
    console.error("❌ carrier-registry 完整性问题:");
    for (const issue of registry.issues) {
      console.error(`  - ${issue}`);
    }
    process.exit(1);
  }
  console.log("✅ carrier-registry 完整性检查通过");

  // 4. 全部通过
  console.log("✅ L5 进化引擎健康检查通过");
  process.exit(0);
}

main();
