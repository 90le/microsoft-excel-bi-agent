# Microsoft Excel BI Agent

![Microsoft Excel BI Agent hero](assets/readme-hero.png)

[![Release](https://img.shields.io/github/v/release/90le/microsoft-excel-bi-agent?include_prereleases&style=flat-square)](https://github.com/90le/microsoft-excel-bi-agent/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg?style=flat-square)](LICENSE)
[![Skills](https://img.shields.io/badge/skills-12_excel_bi_workflows-217346?style=flat-square)](.agents/skills)
[![Agents](https://img.shields.io/badge/agents-Codex%20%7C%20Claude%20%7C%20OpenCode-blue?style=flat-square)](docs/install-and-sync.md)

[English](#english) | [中文](#中文)

## English

**Make AI agents reliable on real Microsoft Excel BI workbooks.**

Microsoft Excel BI Agent is an open-source, cross-agent skill pack for teams that use AI agents to inspect, modify, debug, validate, and publish Excel BI workbooks. It covers **Excel VBA**, **Power Query M**, **Power Pivot DAX**, **MDX/CUBE formulas**, **ADO/SQL**, workbook QA, clean deliverables, Office diagnostics, report building, semantic model review, and sanitized testing fixtures.

It is built for the messy Excel work that generic coding agents usually mishandle: hidden sheets, macro-enabled workbooks, Power Query refresh timing, Data Model boundaries, CUBEVALUE formulas, external links, client-ready `.xlsx` publishing, and Windows Excel COM validation.

### What It Helps Agents Do

| Area | What the skill pack adds |
| --- | --- |
| Excel/VBA workbook engineering | Export/import VBA modules, bind buttons, debug compile/runtime errors, preserve workbook structure |
| Power Query M | Read, edit, refresh, diagnose errors, track dependencies, and wait for refresh completion |
| Power Pivot / DAX | Review measures, context behavior, relationships, and Excel Data Model boundaries |
| MDX / CUBE formulas | Explain and audit `CUBEVALUE`, `CUBEMEMBER`, measures, members, and helper-cell references |
| ADO / SQL | Query workbook tables, files, and Data Model sources through ADO/OLEDB/ADOMD patterns |
| Client deliverables | Freeze formulas, remove links/queries/Data Model dependencies, delete config/process sheets |
| Workbook QA | Audit formulas, hidden sheets, controls, external dependencies, and delivery risk |
| Cross-agent installation | Sync the same skills into Codex, Claude, and OpenCode style folders |

### Install

#### Option A: Codex marketplace install

This follows the Codex plugin marketplace pattern used by mature Codex plugin repos.

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

#### Option B: One-command local install

Use this when you want to clone, inspect, customize, or distribute the repo internally.

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Windows shortcuts:

```powershell
.\install.ps1
```

```cmd
install.cmd
```

macOS, Linux, and Git Bash:

```bash
sh install.sh
```

#### Option C: Manual install

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

### npm / npx Status

This repository does not publish an npm package yet, so it intentionally does **not** advertise a fake `npx` command. The current one-command installer is `node tools/install.mjs`. If the project later publishes an npm wrapper, the recommended public install command can become:

```bash
npx microsoft-excel-bi-agent install
```

### Verify

Public validation, suitable for Windows, macOS, Linux, and Git Bash:

```bash
python tools/validate-skills.py .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
node tools/install.mjs --check
```

Maintainer release gates may require maintainer-only docs and local runtime evidence. Full runtime validation requires Windows desktop Excel:

```powershell
python tools\run_release_gate.py --project-root .
```

The full gate is for Excel COM, VBA execution, Power Query refresh, Power Pivot/Data Model behavior, providers, and rendered workbook evidence.

### Included Skills

| Area | Skill |
| --- | --- |
| Routing | `excel-bi-router` |
| VBA and workbook automation | `excel-vba-workbook-engineering` |
| Power Query M | `power-query-m-engineering` |
| Power Pivot DAX | `power-pivot-dax-modeling` |
| MDX / CUBE formulas | `mdx-cubevalue-extraction` |
| ADO / SQL data access | `excel-ado-sql-data-access` |
| Clean Excel deliverables | `excel-deliverable-publisher` |
| Workbook QA | `excel-workbook-qa-auditor` |
| Report building | `excel-report-builder` |
| Office environment diagnostics | `office-environment-diagnostics` |
| Power BI semantic model context | `power-bi-semantic-model` |
| Sanitized test workbooks | `excel-testing-fixtures` |

## 中文

**让 AI Agent 更可靠地处理真实的 Microsoft Excel BI 工作簿。**

Microsoft Excel BI Agent 是一个开源、跨 Agent 的 Excel BI 技能包，用于让 Codex、Claude、OpenCode 等 AI Agent 更稳定地检查、修改、调试、验证和交付 Excel BI 工作簿。它覆盖 **Excel VBA**、**Power Query M**、**Power Pivot DAX**、**MDX/CUBE 公式**、**ADO/SQL**、工作簿 QA、纯净交付物、Office 环境诊断、报表搭建、语义模型审阅和脱敏测试样例。

它解决的是普通代码 Agent 很容易处理错的 Excel 场景：隐藏过程表、宏工作簿、Power Query 刷新时序、数据模型边界、`CUBEVALUE` 公式、外部链接、客户交付 `.xlsx` 清理，以及 Windows Excel COM 级别验证。

### 它可以让 Agent 做什么

| 场景 | 能力 |
| --- | --- |
| Excel/VBA 工程 | 导出/导入 VBA 模块、绑定按钮、定位编译和运行错误、保护工作簿结构 |
| Power Query M | 读取、编辑、插入、删除、刷新查询，等待刷新结束，定位刷新错误和依赖关系 |
| Power Pivot / DAX | 审阅度量值、筛选上下文、关系、Excel Data Model 边界 |
| MDX / CUBE 公式 | 解释和检查 `CUBEVALUE`、`CUBEMEMBER`、度量值、成员和辅助单元格引用 |
| ADO / SQL | 通过 ADO/OLEDB/ADOMD 查询工作簿表、外部文件和数据模型 |
| 客户交付 | 公式转值、删除外链、删除查询和数据模型依赖、清理配置表和过程表 |
| 工作簿 QA | 检查公式、隐藏表、控件、外部依赖和交付风险 |
| 跨 Agent 分发 | 从一份技能源同步到 Codex、Claude、OpenCode 目录 |

### 安装

#### 方式 A：Codex 插件市场安装

这是推荐给 Codex 用户的安装方式，参考了 Codex 官方插件市场和高星插件仓库的分发方式。

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

#### 方式 B：克隆后一键本地安装

适合需要内部定制、二次开发、团队分发或离线审查的场景。

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Windows:

```powershell
.\install.ps1
```

```cmd
install.cmd
```

macOS、Linux、Git Bash:

```bash
sh install.sh
```

#### 方式 C：手动安装

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

### 为什么没有直接写 npx

高星 CLI 项目常见的 `npx xxx` 安装方式，需要项目已经发布 npm 包。本项目当前还没有发布 npm 包，所以不会写一个实际上不可执行的 `npx` 命令。当前跨平台一键入口是：

```bash
node tools/install.mjs
```

后续如果发布 npm wrapper，可以再把公开安装命令升级为：

```bash
npx microsoft-excel-bi-agent install
```

### 校验

适用于 Windows、macOS、Linux、Git Bash 的公开校验：

```bash
python tools/validate-skills.py .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
node tools/install.mjs --check
```

维护者级 release gate 可能需要未放入公开仓库的维护文档和本地运行证据。完整运行时校验需要 Windows 桌面版 Excel：

```powershell
python tools\run_release_gate.py --project-root .
```

完整校验覆盖 Excel COM、VBA 执行、Power Query 刷新、Power Pivot/Data Model、Provider 环境和渲染证据。

## Documentation / 文档

- [Project overview](docs/project.md)
- [Install and sync guide](docs/install-and-sync.md)
- [Task recipes](docs/task-recipes.md)
- [Compatibility boundaries](docs/compatibility.md)
- [Distribution checklist](docs/distribution-checklist.md)
- [Chinese recipient guide](docs/recipient-guide.zh-CN.md)
- [One-click install prompt EN](prompts/one-click-install-prompt.en-US.md)
- [One-click install prompt CN](prompts/one-click-install-prompt.zh-CN.md)
- [HTML introduction page](docs/intro.html)
- [GitHub publishing notes](docs/github-publish.md)

## Boundaries / 边界

- Do not store customer workbooks, screenshots, PDFs, credentials, local machine paths, or generated QA reports inside the plugin package.
- `.agents/skills/` is the source of truth. `skills/`, `.claude/skills/`, and `.opencode/skills/` are generated mirrors.
- macOS and Linux can validate structure, prompts, OpenXML, and non-COM scripts. They do not prove Excel COM, VBA, Power Query refresh, or Power Pivot runtime behavior.
- This package improves agent operating discipline. It does not replace workbook-specific business review.

## License

MIT
