# 仓库治理目标

本文定义 Microsoft Excel BI Agent 的公开 GitHub 仓库治理目标，覆盖 issue 入口、PR 审阅、安全报告、社区健康文件，以及影响公开信任的仓库设置。

维护者：**丘彬彬**<br>
微信：**binstudy**<br>
博客：**https://90le.cn**

## 目标

让 GitHub 仓库可以安全参与，同时不增加用户上传私有工作簿、截图、PDF、凭证、本机路径、生成 QA 报告或未脱敏运行证据的风险。

目标不是增加流程。目标是建立一条克制的公开协作路径：保护用户、保持安装说明真实，并让每个贡献都能用现有公开校验审阅。

## 约束

- 公开 issue 模板必须在用户提交前阻止客户文件和私有 artifact。
- PR 模板必须要求校验结果、双语文档意识和技能源/镜像纪律。
- 涉及安全和敏感数据的报告不能通过公开 issue 处理。
- GitHub Wiki 必须保持关闭，因为 `docs/` 和 GitHub Pages 才是维护中的文档表面。
- GitHub Discussions 可以继续关闭，直到确实有足够社区量值得投入管理。
- 新增治理文件必须简洁，不能制造与现有校验冲突的独立流程。

## 边界

- 仓库治理可以增加社区健康文件、issue 模板、PR 模板、安全指引和校验脚本。
- 仓库治理可以更新 GitHub 仓库设置，只要它能降低漂移或泄露风险。
- 仓库治理不替代具体工作簿的业务审核。
- 仓库治理不新增 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时证明。

## 可以做

- 增加要求脱敏复现步骤和环境边界的 issue 表单。
- 增加覆盖公开校验、双语文档、安装真实性和 artifact hygiene 的 PR 检查清单。
- 增加 `SECURITY.md` 和 `CONTRIBUTING.md`，帮助贡献者正确进入项目。
- 关闭 Wiki，避免出现未审阅的并行文档。
- 为社区健康文件和模板安全提示增加 CI 校验。

## 不能做

- 要求用户在公开 issue 中附加客户工作簿、截图、PDF 或生成 QA 报告。
- 接受只修改生成镜像、却不修改 `.agents/skills/` 的 PR。
- 接受写入不支持 npm/npx 命令的安装或营销改动。
- 把公开 issue 模板视为分享私有工作簿数据的许可。
- 把维护者私有运行证据放入公开仓库治理文件。

## 详细 Goal

| Goal | 价值 | 完成标准 |
| --- | --- | --- |
| 安全 issue 入口 | 公开 issue 是最容易误泄露工作簿和数据的地方。 | Issue 表单要求脱敏复现，并禁止私有 artifact。 |
| PR 审阅纪律 | 开源贡献需要与维护者变更相同的发布纪律。 | PR 模板要求校验命令、双语文档意识和技能源/镜像检查。 |
| 安全报告路由 | 敏感报告需要私密路径。 | `SECURITY.md` 明确不要在公开 issue 披露凭证或私有工作簿细节。 |
| 文档源控制 | Wiki 会绕过正常审阅并与 Pages 漂移。 | 仓库 Wiki 关闭，文档入口指向 `docs/` 和 Pages。 |
| CI 托底治理 | 治理文件不能无声退化。 | 公开校验包含社区健康文件检查。 |

## 高价值 Backlog

| 优先级 | 事项 | 边界 |
| --- | --- | --- |
| P0 | 增加 issue 表单、PR 模板、`SECURITY.md` 和 `CONTRIBUTING.md`。 | 表单保持简短，不请求客户数据。 |
| P0 | 在 CI 中校验社区健康文件。 | 校验为结构检查，不调用 GitHub API。 |
| P0 | 关闭 Wiki。 | 文档继续保存在仓库文件和 Pages。 |
| P1 | 未来只有在真实社区量出现后才开启 GitHub Discussions。 | 不提前制造无人管理的支持论坛。 |
| P2 | 后续增加贡献者 quickstart 短视频。 | 必须只使用脱敏样例。 |

## 必跑公开校验

修改治理文件、issue 模板、PR 模板、文档、安装流程或校验脚本后运行：

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
