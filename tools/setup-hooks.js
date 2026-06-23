#!/usr/bin/env node
// tools/setup-hooks.js — 幂等配置 core.hooksPath → .githooks
// 用法: node tools/setup-hooks.js

"use strict";

const { execSync } = require("child_process");
const { existsSync } = require("fs");
const { resolve } = require("path");

const options = { encoding: "utf-8", windowsHide: true };

const hookFile = resolve(".githooks/pre-commit");

// 1. 检查 hook 文件存在
if (!existsSync(hookFile)) {
  console.error("❌ .githooks/pre-commit 不存在");
  process.exit(1);
}

// 2. 配置 core.hooksPath
try {
  execSync("git config core.hooksPath .githooks", { ...options, stdio: "pipe" });
} catch (e) {
  console.error("❌ git config core.hooksPath 失败:", e.message);
  process.exit(1);
}

// 3. 验证
try {
  const result = execSync("git config core.hooksPath", { ...options, stdio: "pipe" }).trim();
  if (result !== ".githooks") {
    console.error(`❌ 验证失败: core.hooksPath = '${result}', 期望 '.githooks'`);
    process.exit(1);
  }
} catch (e) {
  console.error("❌ 验证失败:", e.message);
  process.exit(1);
}

console.log("✅ core.hooksPath 已配置为 .githooks");
process.exit(0);
