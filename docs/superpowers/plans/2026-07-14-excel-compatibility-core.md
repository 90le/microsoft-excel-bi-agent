# Excel Compatibility Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Deliver a capability-aware, smaller, and more portable Microsoft Excel BI Agent release with compatibility evidence for legacy, desktop, offline, Mac/web, and Microsoft 365 scenarios.

**Architecture:** Keep existing public skill IDs stable. Add a Windows capability probe and a cross-platform compatibility report, route compatibility questions through the existing environment skill, and validate all behavior with synthetic fixtures plus an optional Windows live gate. Build a compact allowlisted Codex runtime package from the cross-agent source repository.

**Tech Stack:** Python 3 standard library, Windows PowerShell 5.1+, Excel COM late binding, JSON Schema Draft 2020-12, Markdown skills, GitHub Actions, Git.

## Global Constraints

- `.agents/skills/` remains the only hand-edited skill source; all other skill trees are generated mirrors.
- Existing skill IDs remain stable in this release.
- Missing Excel/provider capability is report evidence, not a package crash.
- Structural checks never claim live Excel runtime correctness.
- Windows PowerShell 5.1 and Unicode paths are supported.
- No customer workbook, local credential, or raw exception stack is committed.

---

### Task 1: Release portability and plugin product hygiene

**Files:**
- Modify: `tools/run_release_gate.py`
- Modify: `.codex-plugin/plugin.json`
- Create: `docs/privacy-policy.md`
- Create: `docs/terms-of-service.md`
- Create: `tests/test_release_gate_portability.py`
- Modify: `.github/workflows/validate.yml`

**Interfaces:**
- Produces: `run_command(command, cwd, name, ok_codes=None)` that always returns string stdout/stderr under non-UTF Windows locales.
- Produces: `find_bash()` that prefers Git for Windows Bash before WSL/system Bash on Windows.

- [ ] **Step 1: Add failing portability tests**

Create unittest cases that patch `subprocess.run` with undecodable bytes/`None` output and patch `shutil.which("bash")` to a system Bash while a Git Bash candidate exists. Assert no exception and Git Bash selection.

- [ ] **Step 2: Verify the tests fail**

Run: `python -m unittest tests.test_release_gate_portability -v`

Expected: failure from default locale decoding and/or wrong Bash candidate.

- [ ] **Step 3: Implement robust decoding and Bash selection**

Run subprocesses with explicit UTF-8 plus replacement fallback, defensively normalize `None`, and order Windows Git Bash candidates before `PATH` Bash.

- [ ] **Step 4: Complete the plugin interface**

Add `websiteURL`, `privacyPolicyURL`, and `termsOfServiceURL`. Replace nine starter prompts with exactly three prompts covering inspect/route, diagnose/evidence, and publish/verify. Add concise privacy and terms documents.

- [ ] **Step 5: Add Windows structural CI**

Add a `windows-latest` job that runs the structural gate from a Unicode-containing checkout subdirectory with UTF-8 output enabled.

- [ ] **Step 6: Run focused checks**

Run: `python -m unittest tests.test_release_gate_portability -v`

Run: `python tools/validate-skills.py .`

Expected: all pass.

### Task 2: Capability probe and compatibility report contracts

**Files:**
- Create: `tools/probe_excel_capabilities.ps1`
- Create: `tools/build_excel_compatibility_report.py`
- Create: `tools/create_excel_capability_fixture.py`
- Create: `schemas/excel-capability-probe.schema.json`
- Create: `schemas/excel-compatibility-report.schema.json`
- Create: `tests/test_excel_compatibility_report.py`

**Interfaces:**
- `probe_excel_capabilities.ps1 -OutJson <path> [-Profile inventory|runtime]`
- `build_excel_compatibility_report.py --probe-json <path> [--require-capability <id>]... [--out-json <path>] [--out-md <path>] [--require-pass]`
- Probe output kind: `excel-capability-probe`, schema version `1.0`.
- Report output kind: `excel-compatibility-report`, schema version `1.0`.

- [ ] **Step 1: Add failing report tests**

Test all-pass, required capability failure, skipped required capability, malformed status, and operation readiness. Verify `--require-pass` returns 1 only for invalid evidence or unmet explicit requirements.

- [ ] **Step 2: Verify report tests fail**

Run: `python -m unittest tests.test_excel_compatibility_report -v`

Expected: missing module/tool failures.

- [ ] **Step 3: Implement schemas and report builder**

Implement strict top-level validation, stable capability IDs, summaries, requirement evaluation, operation readiness, JSON/Markdown output, and boundary language.

- [ ] **Step 4: Implement fixture generator**

Generate `all-supported`, `core-blocked`, `partial-evidence`, and `malformed-contract` JSON fixtures without customer data.

- [ ] **Step 5: Implement Windows probe**

Use late-bound COM and temporary synthetic workbooks. Reuse `tools/probe_excel_bi_providers.ps1` rather than duplicating provider detection. Always close only Excel processes created by the probe.

- [ ] **Step 6: Run focused checks**

Run: `python -m unittest tests.test_excel_compatibility_report -v`

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tools/probe_excel_capabilities.ps1 -OutJson "$env:TEMP/excel-capabilities.json" -Profile inventory`

Expected: tests pass and probe emits valid JSON.

### Task 3: Capability-first routing and task profile integration

**Files:**
- Modify: `.agents/skills/excel-bi-router/SKILL.md`
- Modify: `.agents/skills/excel-bi-router/scripts/route_excel_bi_task.py`
- Modify: `.agents/skills/office-environment-diagnostics/SKILL.md`
- Modify: `tools/run_task_profile.py`
- Modify: `tools/build_capability_catalog.py`
- Modify: `AGENTS.md`

**Interfaces:**
- Compatibility/platform requests route to `office-environment-diagnostics`.
- DAX compatibility requests remain routed to `power-pivot-dax-modeling`.
- `env-diagnostics` task profile accepts captured probe evidence and explicit requirements.

- [ ] **Step 1: Add failing router/profile fixture cases**

Add Linux/macOS/Excel COM compatibility prompts and DAX compatibility prompts to the existing router fixture. Add profile assertions for captured probe JSON and unmet requirements.

- [ ] **Step 2: Verify the new cases fail**

Run the router script and task profile fixture commands used by `tools/run_release_gate.py`.

- [ ] **Step 3: Implement routing and skill guidance**

Add capability-first rules, three evidence levels, target-environment separation, and platform fallbacks. Do not rename skills.

- [ ] **Step 4: Integrate task profile**

Extend `env-diagnostics` to run the new probe/report while preserving provider-detail reporting and existing CLI compatibility.

- [ ] **Step 5: Update capability catalog**

Register the new tools and compatibility workflow.

### Task 4: Compact runtime package

**Files:**
- Create: `tools/build_runtime_package.py`
- Create: `tests/test_runtime_package.py`
- Modify: `tools/deploy-local-plugin.py`
- Modify: `tools/install.mjs`
- Modify: `docs/install-and-sync.md`

**Interfaces:**
- `build_runtime_package.py --project-root . --out-dir <dir> [--zip <path>] [--require-pass]`
- Output: allowlisted runtime tree plus `runtime-package-manifest.json` containing relative paths, sizes, SHA-256 hashes, and total bytes.

- [ ] **Step 1: Add failing allowlist tests**

Assert the runtime package includes `.codex-plugin/`, `skills/`, required tools/fixtures, `LICENSE`, and a compact README; excludes `.agents/`, `.claude/`, `.opencode/`, `.git/`, private artifacts, and development docs.

- [ ] **Step 2: Verify tests fail**

Run: `python -m unittest tests.test_runtime_package -v`

- [ ] **Step 3: Implement deterministic staging and zip**

Sort paths, normalize archive separators, calculate SHA-256, and fail on unresolved required skill references or forbidden files.

- [ ] **Step 4: Route local deployment through staging**

Preserve source-repository install behavior while making Codex cache deployment consume the staged runtime package.

- [ ] **Step 5: Run package checks**

Run: `python -m unittest tests.test_runtime_package -v`

Run: `python tools/build_runtime_package.py --project-root . --out-dir "$env:TEMP/excel-bi-runtime" --zip "$env:TEMP/excel-bi-runtime.zip" --require-pass`

Expected: package contains no cross-agent mirror duplication and reports a reduced size.

### Task 5: Documentation, regression cases, mirrors, and release evidence

**Files:**
- Modify: `docs/compatibility.md`
- Modify: `docs/task-recipes.md`
- Modify: `docs/current-status.md`
- Modify: `docs/release-notes.en-US.md`
- Modify: `docs/release-notes.zh-CN.md`
- Modify: `README.md`
- Modify: `README.zh-CN.md`
- Create: `fixtures/real-sanitized-cases/cases/excel-capability-routing.json`
- Modify: `fixtures/real-sanitized-cases/manifest.json`
- Modify: `tools/run_release_gate.py`

**Interfaces:**
- Public docs define structural, runtime-capability, and workbook-behavior evidence separately.
- Regression library includes an environment/compatibility case without private workbooks.

- [ ] **Step 1: Document compatibility tiers and workflows**

Cover Windows/Mac/web, legacy Office, LTSC, Microsoft 365, offline mode, third-party spreadsheet structural compatibility, authoring/automation/consumer/recipient targets, and support confidence labels.

- [ ] **Step 2: Add sanitized regression case**

Require capability probing before platform-specific implementation and preserve DAX specialist routing.

- [ ] **Step 3: Add release-gate checks**

Add structural compatibility fixture/report checks, runtime-package checks, manifest starter prompt checks, and optional Windows live probe checks.

- [ ] **Step 4: Sync mirrors**

Run: `python tools/sync-skills.py --project-root . --all-project-mirrors --replace`

- [ ] **Step 5: Run final validation**

Run: `python -m unittest discover -s tests -v`

Run: `python tools/run_release_gate.py --project-root . --profile structural`

Run on Windows: `python tools/run_release_gate.py --project-root .`

Expected: all required checks pass; unavailable optional runtime components are reported as skip/blocked evidence, not false success.

### Task 6: Publish branch and draft pull request

**Files:**
- Review all changed files.

- [ ] **Step 1: Inspect final scope**

Run: `git status -sb` and `git diff --check`.

- [ ] **Step 2: Commit intentional changes**

Commit focused implementation changes with terse messages; do not include generated local evidence.

- [ ] **Step 3: Push branch**

Run: `git push -u origin agent/excel-compatibility-core`.

- [ ] **Step 4: Open draft PR**

Create a draft PR against `main` describing architecture, compatibility impact, validation evidence, and remaining live-version test boundaries.

