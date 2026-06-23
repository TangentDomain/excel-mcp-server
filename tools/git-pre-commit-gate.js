#!/usr/bin/env node
// tools/git-pre-commit-gate.js — pre-commit 门禁
// 检查: 冲突标记、敏感信息、ruff format、ruff check、pytest invariants、docstring 契约

"use strict";

const { execSync } = require("child_process");

const options = { encoding: "utf-8", windowsHide: true };

function run(cmd) {
  try {
    return execSync(cmd, { ...options, stdio: "pipe" }).trim();
  } catch (e) {
    return null;
  }
}

function runCheck(cmd) {
  try {
    execSync(cmd, { ...options, stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

// 1. 获取 staged 文件
const staged = run("git diff --cached --name-only");
if (!staged) {
  console.log("✅ 无 staged 文件，通过");
  process.exit(0);
}

const files = staged.split("\n").filter(Boolean);

// 2. 冲突标记检查（Node 原生实现，跨平台不依赖 grep）
const diffOutput = run("git diff --cached -U0");
const conflictRegex = /^\+(<{7}|={7}|>{7})/m;
if (diffOutput && conflictRegex.test(diffOutput)) {
  console.error("❌ 冲突标记未解决（发现 <<<<<<< / ======= / >>>>>>> 标记）");
  process.exit(1);
}
console.log("✅ 冲突标记检查通过");

// 3. 敏感信息检查
const secretPatterns = [
  "(?i)(password|secret|api_key|apikey|token|private_key)\\s*[:=]\\s*['\"][^'\"]+['\"]",
  "(?i)-----BEGIN (RSA |EC |DSA )?PRIVATE KEY-----",
];
for (const pat of secretPatterns) {
  try {
    const match = execSync(
      `git diff --cached | grep -cE '${pat.replace(/'/g, "'\\''")}'`,
      { ...options, stdio: "pipe" }
    );
    if (parseInt(match, 10) > 0) {
      console.error(`❌ 检测到疑似敏感信息`);
      process.exit(1);
    }
  } catch {
    // grep -c returns non-zero when count is 0, which is fine
  }
}
console.log("✅ 敏感信息检查通过");

// 4. doc-only bypass: 仅 .md 文件则跳过代码检查
const nonDocFiles = files.filter((f) => !f.endsWith(".md"));
if (nonDocFiles.length === 0) {
  console.log("✅ 仅文档变更，跳过代码检查");
  process.exit(0);
}

// 5. ruff format --check
if (!runCheck("ruff format --check src/ tests/")) {
  console.error("❌ ruff format --check 失败");
  process.exit(1);
}
console.log("✅ ruff format 检查通过");

// 6. ruff check
if (!runCheck("ruff check src/ tests/")) {
  console.error("❌ ruff check 失败");
  process.exit(1);
}
console.log("✅ ruff check 检查通过");

// 7. pytest invariants
if (!runCheck("python -m pytest tests/invariants/ -q --tb=short")) {
  console.error("❌ 不变量测试失败");
  process.exit(1);
}
console.log("✅ 不变量测试通过");

// 8. docstring 契约检查（公共函数必须有 docstring）
const pyFiles = files.filter((f) => f.endsWith(".py"));
if (pyFiles.length > 0) {
  for (const file of pyFiles) {
    const content = run(`git show :${file}`);
    if (!content) continue;
    // 检查 def 开头的公共函数（非 _ 前缀）是否紧跟 docstring
    const lines = content.split("\n");
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].match(/^def ([a-z]\w+)\(/);
      if (m) {
        const nextNonBlank = lines.slice(i + 1).find((l) => l.trim().length > 0);
        if (!nextNonBlank || (!nextNonBlank.trim().startsWith('"""') && !nextNonBlank.trim().startsWith("'"))) {
          console.error(`❌ ${file}:${i + 1} — 公共函数 ${m[1]}() 缺少 docstring`);
          process.exit(1);
        }
      }
    }
  }
  console.log("✅ docstring 契约检查通过");
}

console.log("\n✅ 所有门禁检查通过");
process.exit(0);
