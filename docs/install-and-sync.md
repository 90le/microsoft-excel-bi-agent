# Install And Sync / 安装与同步

This is the constrained installation guide for Microsoft Excel BI Agent. It keeps public installation, local development, and cross-agent skill sync separate.

这是 Microsoft Excel BI Agent 的受控安装说明，用于区分公开安装、本地开发安装和跨 Agent 技能同步。

## Recommended Order / 推荐顺序

| Scenario | Use | Command |
| --- | --- | --- |
| Codex user wants the simplest public install | Codex marketplace | `codex plugin marketplace add` + `codex plugin add` |
| Team wants to inspect or customize the repo | Local one-command installer | `node tools/install.mjs` |
| Enterprise/internal distribution needs explicit control | Manual scripts | `deploy-local-plugin.py` + `sync-skills.py` |
| Claude/OpenCode project usage | Generated mirrors | `.claude/skills/` and `.opencode/skills/` |

## Source Of Truth / 技能源头

```text
.agents/skills/
```

Generated mirrors / 生成镜像:

- `skills/` for Codex plugin packaging.
- `.claude/skills/` for Claude project usage.
- `.opencode/skills/` for OpenCode project usage.
- `~/.codex/skills/` only when user-level Codex skills are explicitly wanted.

Do not edit generated mirrors directly. Update `.agents/skills/`, then sync.

不要直接编辑生成镜像。需要修改时，先改 `.agents/skills/`，再执行同步。

## Option A: Codex Marketplace / Codex 插件市场

This follows the current Codex CLI marketplace pattern:

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

To confirm the CLI supports this pattern:

```bash
codex plugin --help
codex plugin marketplace --help
codex plugin add --help
```

中文说明：

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

第一条命令把 GitHub 仓库加入 Codex 插件市场来源，第二条命令安装该市场里的插件。

## Option B: One-Command Local Install / 本地一键安装

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Windows PowerShell:

```powershell
.\install.ps1
```

Windows CMD:

```cmd
install.cmd
```

macOS, Linux, Git Bash:

```bash
sh install.sh
```

The installer runs:

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

`deploy-local-plugin.py` refreshes the Codex `skills/` mirror, builds a compact runtime staging package, copies only that runtime package into the local Codex plugin directory, updates the personal marketplace, and optionally invokes `codex plugin add`. The second command preserves the existing source-repository behavior by syncing canonical skills to the project mirrors and user-level Codex skills.

`deploy-local-plugin.py` 会先刷新 Codex 的 `skills/` 镜像，再构建精简 runtime staging，只把 runtime 包复制到本地 Codex 插件目录，并更新个人 marketplace、按需执行 `codex plugin add`。第二条命令继续保持源码仓库原有语义：把 canonical skills 同步到项目镜像和用户级 Codex skills。

## Option C: Manual Install / 手动安装

PowerShell:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Bash:

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Use `--update-cachebuster` only when behavior changed and a new installed plugin version is required:

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install --update-cachebuster
```

## Compact Runtime Package / 精简运行时包

The source repository remains the authoring and cross-agent distribution surface. A local Codex plugin cache does not need canonical authoring sources, duplicate Claude/OpenCode mirrors, Git history, development documentation, or release-maintainer files.

源码仓库仍然是编辑和跨 Agent 分发入口。本地 Codex 插件缓存不需要 canonical 编辑源、重复的 Claude/OpenCode 镜像、Git 历史、开发文档或发布维护文件。

Build a runtime staging tree and deterministic zip without installing it:

```powershell
python tools\build_runtime_package.py --project-root . --out-dir "$env:TEMP\excel-bi-runtime" --zip "$env:TEMP\excel-bi-runtime.zip" --require-pass
```

```bash
python tools/build_runtime_package.py --project-root . --out-dir "${TMPDIR:-/tmp}/excel-bi-runtime" --zip "${TMPDIR:-/tmp}/excel-bi-runtime.zip" --require-pass
```

The runtime allowlist contains:

- `.codex-plugin/` and the current `skills/` Codex mirror.
- Runtime tools referenced by packaged skills, plus their referenced helper tools.
- Sanitized `fixtures/` and compatibility `schemas/` when present.
- `LICENSE`, a compact generated `README.md`, and `runtime-package-manifest.json`.

It excludes `.agents/`, `.claude/`, `.opencode/`, `.git/`, development docs, private workbooks, generated release evidence, caches, and lock files. The manifest records sorted relative paths, byte sizes, SHA-256 hashes, total payload bytes, source/runtime size reduction, validation errors, and mirror status. Zip entry order, timestamps, permissions, and separators are normalized for deterministic output.

`build_runtime_package.py` does not sync or edit any skill mirror. If `.agents/skills/` and `skills/` differ, the manifest reports drift as a warning and packages the existing `skills/` tree. Maintainers must run the documented sync command before a final release; runtime packaging must never silently change canonical or generated skill sources.

`--require-pass` fails for unresolved file references in packaged skills or forbidden payload artifacts. `node tools/install.mjs --check` builds and discards a temporary runtime package as part of public validation.

## Claude And OpenCode / Claude 与 OpenCode

This repo does not claim a universal Claude/OpenCode marketplace command because those ecosystems vary by version and installation surface. Instead, it provides generated project mirrors:

本项目不强行宣称所有 Claude/OpenCode 环境都有同一条插件市场命令，因为不同版本和入口存在差异。当前稳定做法是生成项目级技能镜像：

```text
.claude/skills/
.opencode/skills/
```

After sync, point the target agent to the project folder or copy those generated skill folders according to that agent's documented skill/plugin mechanism.

## npm / npx Policy / npm 与 npx 策略

High-star CLI repositories often use `npx` only after publishing an npm package. This project does not publish one yet, so this guide does not invent a non-working `npx` command.

高星 CLI 项目常用 `npx` 的前提是已经发布 npm 包。本项目当前尚未发布 npm 包，所以文档不会写一个无法执行的 `npx` 命令。

Current cross-platform entry:

```bash
node tools/install.mjs
```

Future npm wrapper shape, only after publishing:

```bash
npx microsoft-excel-bi-agent install
```

## Checks / 校验

Public checks:

```bash
python tools/validate-skills.py .
python tools/validate_project_docs.py --project-root .
python tools/validate_github_community_health.py --project-root .
python tools/validate_task_recipes.py --project-root .
python tools/validate_official_docs_index.py --project-root .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
python tools/build_goal_coverage_report.py --project-root . --require-pass
node tools/install.mjs --check
```

Maintainer release gate. This can require maintainer-only docs and local runtime evidence that are intentionally not stored in the public repository:

```bash
python tools/run_release_gate.py --project-root . --profile structural
```

Full runtime gate, Windows desktop Excel only:

```powershell
python tools\run_release_gate.py --project-root .
```

## Install Patterns We Follow / 参考的安装模式

- Codex official plugin workflow: marketplace source first, plugin install second.
- Mature Codex plugin repos: short marketplace command for users, clone/local workflow for maintainers.
- Claude-style plugin/skill repos: project-level skill folders plus optional marketplace command where supported.
- High-star CLI repos: one visible command, clear checks, no undocumented dependency on local state.

## Hard Boundaries / 硬边界

- No customer workbooks, screenshots, PDFs, generated QA reports, credentials, or private local paths in the plugin package.
- No divergent manual edits in `skills/`, `.claude/skills/`, `.opencode/skills/`, or `~/.codex/skills/`.
- No new skills or broad tools unless a repeated workflow cannot be handled by an existing skill plus script.
- Linux/macOS structural checks are not proof of Excel COM refresh, VBA execution, Data Model behavior, or rendered Excel output.
