#!/usr/bin/env node
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, "..");
const args = new Set(process.argv.slice(2));

const REMOTE_MARKETPLACE = "90le/microsoft-excel-bi-agent";
const MARKETPLACE_NAME = "microsoft-excel-bi-agent";
const PLUGIN_NAME = "microsoft-excel-bi-agent-pack";

function say(message) {
  process.stdout.write(`${message}\n`);
}

function fail(message, code = 1) {
  process.stderr.write(`\n[error] ${message}\n`);
  process.exit(code);
}

function commandExists(command) {
  const probe = process.platform === "win32" ? "where" : "command";
  const probeArgs = process.platform === "win32" ? [command] : ["-v", command];
  const result = spawnSync(probe, probeArgs, { shell: process.platform !== "win32", stdio: "ignore" });
  return result.status === 0;
}

function run(command, commandArgs, options = {}) {
  say(`\n$ ${[command, ...commandArgs].join(" ")}`);
  const result = spawnSync(command, commandArgs, {
    cwd: options.cwd || root,
    stdio: "inherit",
    shell: false
  });
  if (result.error) fail(`${command} failed: ${result.error.message}`);
  if (result.status !== 0) fail(`${command} exited with code ${result.status}`);
}

function pythonCommand() {
  if (commandExists("python")) return "python";
  if (commandExists("python3")) return "python3";
  fail("Python was not found. Install Python 3, then rerun this script.");
}

function requireRepoShape() {
  const required = [
    ".codex-plugin/plugin.json",
    "skills",
    ".agents/skills",
    "tools/sync-skills.py",
    "tools/deploy-local-plugin.py",
    "tools/build_runtime_package.py"
  ];
  for (const item of required) {
    if (!existsSync(join(root, item))) fail(`Missing required project path: ${item}`);
  }
}

function runChecks(py) {
  run(py, ["tools/validate-skills.py", "."]);
  run(py, ["tools/validate_project_docs.py", "--project-root", "."]);
  run(py, ["tools/validate_github_community_health.py", "--project-root", "."]);
  run(py, ["tools/validate_task_recipes.py", "--project-root", "."]);
  run(py, ["tools/validate_official_docs_index.py", "--project-root", "."]);
  run(py, ["tools/build_artifact_hygiene_report.py", "--project-root", ".", "--require-pass"]);
  run(py, ["tools/build_goal_coverage_report.py", "--project-root", ".", "--require-pass"]);
  const runtimeCheckRoot = mkdtempSync(join(tmpdir(), "excel-bi-runtime-check-"));
  try {
    run(py, [
      "tools/build_runtime_package.py",
      "--project-root",
      ".",
      "--out-dir",
      join(runtimeCheckRoot, "runtime"),
      "--require-pass"
    ]);
  } finally {
    rmSync(runtimeCheckRoot, { recursive: true, force: true });
  }
}

function installLocal(py) {
  run(py, ["tools/deploy-local-plugin.py", "--project-root", ".", "--replace", "--install"]);
  run(py, ["tools/sync-skills.py", "--project-root", ".", "--all-project-mirrors", "--codex-user", "--replace"]);
}

function installCodexMarketplace() {
  if (!commandExists("codex")) {
    fail("Codex CLI was not found. Install Codex first, or use the local installer mode.");
  }
  run("codex", ["plugin", "marketplace", "add", REMOTE_MARKETPLACE]);
  run("codex", ["plugin", "add", `${PLUGIN_NAME}@${MARKETPLACE_NAME}`]);
}

function printHelp() {
  say(`Microsoft Excel BI Agent installer

Usage:
  node tools/install.mjs [mode]

Modes:
  --local              Install from this cloned repository and sync Codex/Claude/OpenCode mirrors. Default.
  --codex-marketplace  Install through Codex remote marketplace.
  --check              Run public structural validation only.
  --help               Show this help.

中文:
  --local              从当前仓库安装，并同步 Codex / Claude / OpenCode 技能镜像。默认模式。
  --codex-marketplace  通过 Codex 远程插件市场安装。
  --check              只运行公开结构校验，不安装。
`);
}

if (args.has("--help") || args.has("-h")) {
  printHelp();
  process.exit(0);
}

requireRepoShape();

if (args.has("--codex-marketplace")) {
  installCodexMarketplace();
  say("\n[ok] Installed through the Codex marketplace.");
  say("[完成] 已通过 Codex 插件市场安装。");
  process.exit(0);
}

const py = pythonCommand();

if (args.has("--check")) {
  runChecks(py);
  say("\n[ok] Structural checks passed.");
  say("[完成] 结构校验通过。");
  process.exit(0);
}

installLocal(py);
say("\n[ok] Local plugin install and cross-agent skill sync completed.");
say("[完成] 本地插件安装和跨 Agent 技能同步已完成。");
