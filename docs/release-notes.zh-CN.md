# 发布说明

## v0.1.5 - GitHub 社区健康与安全入口

发布重点：通过更安全的 issue 入口、PR 审阅、安全报告和仓库文档表面，降低公开协作风险。

### 变更

- 新增中英文仓库治理目标，包含目标、约束、边界、可以做/不能做、详细 goal 和高价值 backlog。
- 新增 `CONTRIBUTING.md`、`SECURITY.md`、issue 表单和 pull request 模板。
- 新增 `tools/validate_github_community_health.py`，并接入 `node tools/install.mjs --check` 和 GitHub Actions。
- 更新公开校验文档，使社区健康校验成为发布门禁的一部分。
- 插件 manifest 版本升至 `0.1.5+codex.20260623175347`。

### 校验

公开校验：

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

### 边界

本次发布修改 GitHub 仓库治理和公开入口安全，不声称新增 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时证明。

## v0.1.4 - 公开增长目标与营销准备

发布重点：提升公开信任、采用清晰度、社交分享和营销复用，不改变 Excel 工作簿处理行为。

### 变更

- 新增中英文公开增长目标，包含目标、约束、边界、可以做/不能做、详细 goal 和高价值优化 backlog。
- 新增中英文营销文案包，覆盖发布文案、短标语、广告方向、渠道变体和不要这样宣传规则。
- 在 README、项目文档、Pages 和插件 manifest 中增加维护者署名：丘彬彬，微信 `binstudy`，博客 `90le.cn`。
- 更新网站，增加使用场景转化卡片、证明指标、绝对 Open Graph/Twitter 图片 URL、canonical URL、作者元信息和 v0.1.4 release 可见性。
- 更新项目文档校验和 goal coverage，使公开增长、营销文案、维护者署名和社交元信息纳入 CI 覆盖。
- 插件 manifest 版本升至 `0.1.4+codex.20260623173419`。

### 校验

公开校验：

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

维护者结构门禁：

```bash
python tools/run_release_gate.py --project-root . --profile structural
```

### 边界

本次发布修改公开定位、站点布局、元信息和文档，不声称新增 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时证明。

## v0.1.3 - 公开维护目标与 CI 校验

发布重点：降低公开仓库维护风险，不改变 Excel 工作簿处理行为。

### 变更

- 新增公开维护目标、约束、边界、可以做/不能做、详细 goal、风险清单和优化 backlog。
- 新增 GitHub Actions，在 push 和 pull request 上运行公开结构校验。
- 重写项目文档校验，使其校验公开仓库文档，而不是依赖维护者本地私有台账。
- 重写 goal coverage 和 completion readiness 审计，使其围绕公开维护覆盖和活跃 backlog 状态。
- 扩展 `node tools/install.mjs --check`，使其运行完整公开结构校验集合。
- 更新网站，展示 release、公开校验和运行时边界状态。
- 更新 artifact hygiene 预期：公开包默认不包含任何 Office 工作簿。
- 更新任务配方、分发清单、开源发布说明和真实/脱敏回归文档，移除过期的本地证据假设。
- 插件 manifest 版本升至 `0.1.3+codex.20260623171436`。

### 校验

公开校验：

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

维护者结构门禁：

```bash
python tools/run_release_gate.py --project-root . --profile structural
```

### 边界

本次发布不声称新增 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时证明。这些检查仍然需要 Windows 桌面版 Excel，并且运行时证据应保留在公开仓库之外的任务或发布目录中。

## v0.1.2 - 中英文文档与站点拆分

- 将 README 和项目说明拆分为独立英文入口和中文入口。
- 将网站拆分为英文页和中文页，并保留浏览器语言跳转。
- 安装命令继续限定为已支持的 Codex marketplace 和本地安装器路径。
