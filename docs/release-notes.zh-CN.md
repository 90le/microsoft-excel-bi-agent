# 发布说明

## v0.2.1 - 触发效率与发现成本实测

当前稳定版：v0.2.1。本次发布在不改变 12 个已发布技能 ID、不增加 Excel 功能的前提下，降低插件发现成本并提高技能路由精度。

### 变更

- 将 3 条 manifest starter prompts 缩短至每条不超过 110 个字符。
- 将全部 12 个 canonical skill descriptions 改为精简的 `Use when ...` 触发条件；技能正文、路由行为、ID 与 Excel 功能范围不变。
- 新增不含客户数据的 36 条触发语料，其中 24 条 positive、12 条 confusable-negative；同时新增 3 个真实 plugin-eval 基准场景：`power-query-diagnosis`、`dax-versus-environment`、`delivery-boundary`。其输入与生成响应均为合成材料，不证明真实任务成功；observed usage（观测用量）是独立证据。
- 将 trigger validator 接入 structural release gate 与 capability catalog。
- 从 `.agents/skills/` 同步生成 `skills/`、`.claude/skills/` 和 `.opencode/skills/` 镜像。
- 插件 manifest 版本升至 `0.2.1+codex.20260714`。

### 实测对比

对新 staged runtime 的静态 plugin-eval 分析实测：`trigger_cost_tokens` 从 v0.2.0 的 1,161 降至 682，减少 41.26%；`invoke_cost_tokens` 从 15,365 降至 14,886（-479）。这些数字属于合成/生成的静态估算，不证明真实任务成功；observed usage（观测用量）是独立证据。

### 复现静态对比

```powershell
$pluginEval = '<path-to-plugin-eval.js>'
$runtime = Join-Path $env:TEMP 'excel-bi-v021-runtime'
$before = Join-Path $env:TEMP 'excel-bi-v020-plugin-eval.json'
$after = Join-Path $env:TEMP 'excel-bi-v021-plugin-eval.json'
$compare = Join-Path $env:TEMP 'excel-bi-v020-v021-compare.md'
python tools/build_runtime_package.py --project-root . --out-dir $runtime --require-pass
node $pluginEval analyze $runtime --format json --output $after
node $pluginEval compare $before $after --format markdown --output $compare
```

baseline、staged runtime、analysis JSON、comparison report、benchmark result 与 observed-usage logs 均应保存在系统临时目录。trigger validator 与三场景基准命令见 `docs/task-recipes.md`。

### 校验

```bash
python -m unittest discover -s tests -v
python tools/validate-skills.py .
python tools/run_release_gate.py --project-root . --profile structural
```

### 边界

36 条触发语料和 3 个基准场景使用脱敏合成制品。生成的场景响应不证明真实任务成功，静态 token 估算也不是 observed usage。真实工作簿成功必须另行取得 runtime 观测与代表性 workbook-behavior evidence。

## v0.2.0 - 能力感知的 Excel 兼容性与精简运行时包

发布重点：在旧版、桌面、离线、Mac/Web、Microsoft 365 和接收方环境之间建立基于证据的兼容性结论，同时缩小 Codex 安装运行时载荷。

### 变更

- 新增稳定的 Windows capability probe 合约，覆盖 Excel COM、工作簿保存/重开、VBA 项目访问、Power Query 对象模型与等待、Data Model、PDF 导出、ADO/ACE、MSOLAP 和 ADOMD。
- 新增跨平台 compatibility report，严格区分 structural evidence、runtime capability evidence 和 workbook behavior evidence；显式必需能力可阻塞交付，但可选组件不可用不会被伪装成包崩溃。
- 新增 capability-first 路由：平台、宿主、运行与可用性问题交给 `office-environment-diagnostics`；DAX/Power Pivot 公式与函数兼容性继续交给 `power-pivot-dax-modeling`。
- 新增四组合成 capability fixture 和 `excel-capability-routing` 脱敏回归案例，不需要也不发布私有工作簿。
- 新增白名单式、确定性的 runtime package builder，包含 SHA-256 清单、依赖闭包、精简 README、确定性 zip 和私有制品排除规则。
- 发布门禁新增 compatibility fixture/report、runtime package、版本 manifest 以及恰好三条 inspect/diagnose/publish starter prompts 检查；Windows full gate 额外执行可选 live capability probe。
- 文档覆盖 Windows、macOS、Excel for web、Linux、Excel 2007/2010/2013/2016/2019、Office LTSC、Microsoft 365、32/64 位、offline、WPS/LibreOffice 仅结构兼容，以及 authoring/automation/consumer/recipient 目标。
- 从 `.agents/skills/` 同步生成 `skills/`、`.claude/skills/` 和 `.opencode/skills/` 镜像。
- 插件 manifest 版本升至 `0.2.0+codex.20260714`。

### 校验

```bash
python -m unittest discover -s tests -v
python tools/validate-skills.py .
python tools/run_release_gate.py --project-root . --profile structural
```

在 Windows 桌面版 Excel 上，full gate 还会尝试 live capability probe：

```powershell
python tools\run_release_gate.py --project-root .
```

### 边界

结构证据不能证明 Excel 已执行。运行时能力证据只适用于被探测机器。高置信度兼容性需要在每个存在实质差异的目标环境中取得代表性工作簿行为证据；第三方表格软件在对应宿主完成验证前只能声称结构兼容。

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
