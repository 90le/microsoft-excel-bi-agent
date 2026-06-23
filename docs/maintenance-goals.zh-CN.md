# 维护目标

本文定义 Microsoft Excel BI Agent 的公开维护目标。它面向仓库维护、风险评估和发布判断，不包含维护者私有运行台账，也不包含本机专属证据。

## 目标

让 Microsoft Excel BI Agent 在真实 Excel BI 工作中保持可靠，同时让公开仓库更容易安装、校验、审阅，并能在 Codex、Claude、OpenCode 以及类似 Agent 中延展。

目标不是让 Agent “随便改 Excel”。目标是约束 Agent 的操作纪律：先识别工作簿表面结构，再选择正确的 BI 层，再做小范围修改，最后验证结果，并明确说明哪些运行时检查已经执行、哪些因为环境限制被跳过。

## 必须值得的优化规则

只优先处理能够降低发布风险、阻止虚假安装说明、保护用户数据、提升公开校验，或让其他 Agent/维护者更容易接手的改动。避免纯视觉微调、大范围无关重写、没有重复需求的新技能，或要求把私有证据放进公开仓库的发布产物。

## 约束

- 公开安装命令必须是当前项目真实可执行的命令。除非已经发布 npm 包，否则不要宣传 npm 或 npx 命令。
- 面向用户的中英文文档必须独立维护。不要把长篇中英文内容混排到同一个 README 或站点页面。
- `.agents/skills/` 是技能源头。`skills/`、`.claude/skills/`、`.opencode/skills/` 是生成镜像。
- Excel COM、VBA 执行、Power Query 刷新、Power Pivot 运行时行为必须用 Windows 桌面版 Excel 验证。
- macOS 和 Linux 校验只能证明结构、OpenXML、文档和非 COM 脚本行为。
- 客户文件、截图、PDF、凭证、本机私有路径、生成的 QA 报告必须留在公开包之外。

## 边界

- 本仓库发布技能源、生成镜像、脚本、文档、Prompt 和脱敏样例。
- 本仓库不发布客户工作簿、私有发布台账、本机运行报告或本地 Excel 证据。
- 结构校验不能证明某个私有工作簿计算正确。
- Release notes 可以总结运行证据，但原始运行证据应放在任务本地或发布本地目录中，除非已经明确脱敏。

## 可以做

- 优化公开文档、网站页面、安装说明和校验命令。
- 当重复出现的 Excel BI 工作流需要更清晰的 Agent 步骤时，修改 `.agents/skills/`。
- 修改技能源后同步生成镜像。
- 增加脱敏样例、回归案例和静态 OpenXML 检查。
- 增加不依赖桌面版 Excel 的公开结构校验 CI。
- 收紧 artifact hygiene 和文档一致性校验。

## 不能做

- 提交客户工作簿、截图、PDF、生成的 QA 报告、凭证或本机私有路径。
- 没有 Windows 桌面版 Excel 证据时，声称已经验证 Excel COM、VBA、Power Query 刷新或 Power Pivot 行为。
- 添加项目并不支持的 npm、npx、Claude、OpenCode 或 marketplace 命令。
- 只手动修改生成镜像，而不修改 `.agents/skills/`。
- 把结构校验当成具体工作簿的业务验证。

## 详细 Goal

| Goal | 价值 | 完成标准 |
| --- | --- | --- |
| 安装真实性 | 安装文档有假命令会直接破坏采用。 | README、网站和安装文档只暴露真实安装路径。 |
| 双语独立 | 混排页面难审阅，也会破坏搜索和用户预期。 | 英文和中文文档可以独立编辑、独立审阅。 |
| 技能源纪律 | 镜像漂移会导致不同 Agent 行为不一致。 | 先改技能源，再同步镜像，漂移检查通过。 |
| 公开校验 | 贡献者需要不依赖私有文件或 Excel COM 的快速检查。 | 公开校验能在本地和 GitHub Actions 中通过。 |
| 运行时边界清晰 | 用户必须知道 macOS/Linux 校验不能证明什么。 | 文档和发布说明明确区分结构校验和 Windows Excel 运行时校验。 |
| Artifact hygiene | 公开仓库很容易误提交敏感工作簿产物。 | hygiene 检查会阻止客户文件、本地报告、锁文件和私有路径。 |
| 风险 backlog 可见 | 维护者需要共享高价值下一步。 | 高风险、高价值优化项有明确边界并记录在案。 |

## 风险清单

| 风险 | 严重度 | 缓解方式 |
| --- | --- | --- |
| 公开文档写了不存在的命令。 | 高 | 校验安装文档，并把 `node tools/install.mjs` 作为本地一键入口。 |
| 维护者私有证据泄露到公开包。 | 高 | 私有台账和生成报告保持忽略；发布前运行 artifact hygiene。 |
| 生成镜像与 `.agents/skills/` 漂移。 | 高 | 技能修改后和发布前运行 `tools/sync-skills.py --check-drift`。 |
| Linux/macOS 用户过度信任结构校验。 | 高 | 在文档和校验输出中重复说明 Windows Excel 运行时边界。 |
| 发布信心完全依赖人工本地检查。 | 中 | 在 GitHub Actions 中对 push 和 PR 运行公开结构校验。 |
| 目标和路线图文档发布后变旧。 | 中 | 将本文纳入项目文档校验，打 tag 前复核。 |

## 优化 Backlog

| 优先级 | 事项 | 边界 |
| --- | --- | --- |
| P0 | 保持公开校验在 CI 中通过。 | CI 不能要求桌面版 Excel 或私有 artifact。 |
| P0 | 保持 README、网站和安装文档中的命令真实。 | 未发布 npm 包前，不宣传 npm/npx。 |
| P1 | 从重复工作簿问题中继续扩展脱敏回归案例。 | 案例不能泄露客户数据或本机路径。 |
| P1 | 保持网站对 release 和校验结果的可见性。 | 不嵌入原始本机证据或生成 QA 报告。 |
| P1 | 增加技能源和生成镜像的贡献说明。 | 不鼓励手动修改镜像。 |
| P2 | 增加可选 Windows Excel 运行时证据模板。 | 模板可以描述证据形状；原始证据仍留在仓库外。 |

## 必跑公开校验

修改公开文档、安装流程、校验脚本或插件包装后运行：

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

如果修改插件结构，还要运行 Codex plugin creator skill 中的插件校验脚本。
