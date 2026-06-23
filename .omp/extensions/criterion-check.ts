#!/usr/bin/env bun
// .omp/extensions/criterion-check.ts — PostToolUse 判据 hook
// 监听 write/edit/ast_edit，跑 verify-output.py，warn-only

import { spawnSync } from "child_process";

interface PostToolUseInput {
  toolName: string;
  toolInput: Record<string, unknown>;
}

const WATCHED_TOOLS: Record<string, true> = { write: true, edit: true, ast_edit: true };
const options = { encoding: "utf-8" as const, windowsHide: true, timeout: 60_000 };

export function PostToolUse(input: PostToolUseInput): void {
  if (!WATCHED_TOOLS[input.toolName]) return;

  const r = spawnSync("python", ["scripts/verify-output.py"], {
    cwd: process.cwd(),
    stdio: "pipe",
    ...options,
  });

  const output = r.stdout?.toString().trim() ?? "";

  if (r.status !== 0) {
    console.warn(`\n⚠️  判据检查未通过 (warn-only, 不阻断):`);
    console.warn(output);
  }
}

// OMP Extension entry
const input: PostToolUseInput = JSON.parse(process.argv[2] ?? "{}");
PostToolUse(input);
