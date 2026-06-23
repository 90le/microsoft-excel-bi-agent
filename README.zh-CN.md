# Microsoft Excel BI Agent

![Microsoft Excel BI Agent 中文介绍图](assets/readme-hero.zh-CN.png)

[![Release](https://img.shields.io/github/v/release/90le/microsoft-excel-bi-agent?include_prereleases&style=flat-square)](https://github.com/90le/microsoft-excel-bi-agent/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg?style=flat-square)](LICENSE)
[![Skills](https://img.shields.io/badge/skills-12_excel_bi_workflows-217346?style=flat-square)](.agents/skills)
[![Agents](https://img.shields.io/badge/agents-Codex%20%7C%20Claude%20%7C%20OpenCode-blue?style=flat-square)](docs/install-and-sync.md)

[English](README.md) | [中文](README.zh-CN.md) | [中文站点](https://90le.github.io/microsoft-excel-bi-agent/intro.zh-CN.html)

**让 AI Agent 更可靠地处理真实的 Microsoft Excel BI 工作簿。**

Microsoft Excel BI Agent 是一个开源、跨 Agent 的 Excel BI 技能包，用于让 Codex、Claude、OpenCode 等 AI Agent 更稳定地检查、修改、调试、验证和交付 Excel BI 工作簿。它覆盖 **Excel VBA**、**Power Query M**、**Power Pivot DAX**、**MDX/CUBE 公式**、**ADO/SQL**、工作簿 QA、纯净交付物、Office 环境诊断、报表搭建、语义模型审阅和脱敏测试样例。

它解决的是普通代码 Agent 很容易处理错的 Excel 场景：隐藏过程表、宏工作簿、Power Query 刷新时序、数据模型边界、`CUBEVALUE` 公式、外部链接、客户交付 `.xlsx` 清理，以及 Windows Excel COM 级别验证。

维护者：**丘彬彬**。微信：**binstudy**。博客：**https://90le.cn**。

## 适合什么时候用

- AI Agent 需要检查或修改包含公式、VBA、Power Query、数据模型、CUBE 公式、外链或隐藏过程表的真实工作簿。
- 工作簿需要在交付客户前清理为纯净版本。
- 团队希望用同一份 Excel BI 技能源复用到 Codex、Claude、OpenCode 等 Agent。
- 维护者需要不依赖私有工作簿或桌面版 Excel 的公开结构校验。

## 它能让 Agent 做什么

| 场景 | 能力 |
| --- | --- |
| Excel/VBA 工程 | 导出/导入 VBA 模块、绑定按钮、定位编译和运行错误、保护工作簿结构 |
| Power Query M | 读取、编辑、刷新查询，诊断错误，追踪依赖，并等待刷新完成 |
| Power Pivot / DAX | 审阅度量值、上下文行为、关系和 Excel Data Model 边界 |
| MDX / CUBE 公式 | 解释和检查 `CUBEVALUE`、`CUBEMEMBER`、度量值、成员和辅助单元格引用 |
| ADO / SQL | 通过 ADO/OLEDB/ADOMD 查询工作簿表、外部文件和数据模型来源 |
| 客户交付 | 公式转值、删除外链、删除查询和数据模型依赖、清理配置表和过程表 |
| 工作簿 QA | 检查公式、隐藏表、控件、外部依赖和交付风险 |
| 跨 Agent 分发 | 从一份技能源同步到 Codex、Claude、OpenCode 目录 |

## 安装

### 方式 A：Codex 插件市场

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

### 方式 B：本地一键安装

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

Windows 快捷入口：

```powershell
.\install.ps1
```

```cmd
install.cmd
```

macOS、Linux、Git Bash：

```bash
sh install.sh
```

### 方式 C：手动安装

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## npm / npx 状态

本项目当前尚未发布 npm 包，所以不会写一个实际不可执行的 `npx` 命令。当前跨平台一键入口是：

```bash
node tools/install.mjs
```

## 校验

适用于 Windows、macOS、Linux、Git Bash 的公开校验：

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

完整运行时校验需要 Windows 桌面版 Excel：

```powershell
python tools\run_release_gate.py --project-root .
```

## 包含技能

| 场景 | 技能 |
| --- | --- |
| 路由 | `excel-bi-router` |
| VBA 与工作簿自动化 | `excel-vba-workbook-engineering` |
| Power Query M | `power-query-m-engineering` |
| Power Pivot DAX | `power-pivot-dax-modeling` |
| MDX / CUBE 公式 | `mdx-cubevalue-extraction` |
| ADO / SQL 数据访问 | `excel-ado-sql-data-access` |
| 纯净 Excel 交付物 | `excel-deliverable-publisher` |
| 工作簿 QA | `excel-workbook-qa-auditor` |
| 报表搭建 | `excel-report-builder` |
| Office 环境诊断 | `office-environment-diagnostics` |
| Power BI 语义模型上下文 | `power-bi-semantic-model` |
| 脱敏测试工作簿 | `excel-testing-fixtures` |

## 文档

- [中文项目说明](docs/project.zh-CN.md)
- [English project overview](docs/project.en-US.md)
- [安装与同步说明](docs/install-and-sync.md)
- [任务配方](docs/task-recipes.md)
- [维护目标与风险 backlog](docs/maintenance-goals.zh-CN.md)
- [公开增长目标](docs/growth-goals.zh-CN.md)
- [仓库治理目标](docs/repository-governance-goals.zh-CN.md)
- [营销文案包](docs/marketing-copy.zh-CN.md)
- [发布说明](docs/release-notes.zh-CN.md)
- [贡献说明](CONTRIBUTING.md)
- [安全策略](SECURITY.md)
- [兼容性边界](docs/compatibility.md)
- [分发检查清单](docs/distribution-checklist.md)
- [中文一键安装 Prompt](prompts/one-click-install-prompt.zh-CN.md)
- [English one-click install prompt](prompts/one-click-install-prompt.en-US.md)
- [中文站点](docs/intro.zh-CN.html)
- [English website](docs/intro.html)

## 边界

- 不要把客户工作簿、截图、PDF、凭证、本机私有路径或生成的 QA 报告放进插件包。
- `.agents/skills/` 是技能源头。`skills/`、`.claude/skills/`、`.opencode/skills/` 是生成镜像。
- macOS 和 Linux 可以验证结构、Prompt、OpenXML 和非 COM 脚本，但不能证明 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时行为。
- 本项目提升的是 Agent 操作纪律，不能替代具体工作簿的业务审核。

## License

MIT
