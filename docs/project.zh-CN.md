# Microsoft Excel BI Agent 项目说明

## 项目目标

Microsoft Excel BI Agent 是一个开源、跨 Agent 的 Excel BI 技能包，面向需要处理 Excel 工作簿、VBA、Power Query M、Power Pivot DAX、MDX/CUBE 公式、ADO/SQL、工作簿 QA、纯净交付物、报表搭建、Office 环境诊断和脱敏测试样例的 AI Agent。

项目目标不是让 Agent “随便改 Excel”，而是把反复出现的 Excel BI 风险沉淀成可复用流程：先识别工作簿表面结构，再判断所属 BI 层，再做小范围修改，最后验证结果并说明哪些内容已经验证、哪些内容因环境限制被跳过。

## 适用对象

- 使用 Codex、Claude、OpenCode 或类似 Agent 处理 Excel BI 文件的团队。
- 需要更安全工作簿修改流程的分析师和自动化工程师。
- 需要交付纯净 `.xlsx` 或 `.xlsm` 文件，且不能泄露过程表、链接、凭证、私有路径的交付团队。
- 需要重复检查 VBA、Power Query、DAX、MDX/CUBE、ADO/SQL 和 Office 运行时边界的维护者。

## 项目结构

```text
microsoft-excel-bi-agent/
  README.md                    # 英文仓库入口
  README.zh-CN.md              # 中文仓库入口
  LICENSE
  .codex-plugin/plugin.json
  marketplace.json             # Codex 插件市场元数据
  .agents/skills/              # 技能源头
  skills/                      # 生成的 Codex 插件镜像
  .claude/skills/              # 生成的 Claude 镜像
  .opencode/skills/            # 生成的 OpenCode 镜像
  docs/
  fixtures/
  prompts/
  tools/
```

## 核心能力

| 场景 | 能力 |
| --- | --- |
| Excel/VBA 工程 | 导出/导入 VBA 模块、绑定按钮、定位编译和运行错误、保护工作簿结构 |
| Power Query M | 读取、编辑、刷新查询，诊断错误，追踪依赖，并等待刷新完成 |
| Power Pivot DAX | 审阅度量值、上下文行为、关系和 Excel Data Model 边界 |
| MDX / CUBE 公式 | 检查 `CUBEVALUE`、`CUBEMEMBER`、度量值、成员和辅助单元格引用 |
| ADO / SQL | 通过 ADO/OLEDB/ADOMD 查询工作簿表、外部文件和数据模型来源 |
| 客户交付 | 公式转值、删除外链、删除查询和数据模型依赖、清理配置表和过程表 |
| 工作簿 QA | 检查公式、隐藏表、控件、外部依赖和交付风险 |
| 跨 Agent 分发 | 从同一份技能源同步到 Codex、Claude、OpenCode 目录 |

## 安装方式

Codex 插件市场：

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

本地一键安装：

```bash
git clone https://github.com/90le/microsoft-excel-bi-agent.git
cd microsoft-excel-bi-agent
node tools/install.mjs
```

手动安装：

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## 公开仓库边界

公开仓库包含技能源、脚本、测试样例、安装 Prompt 和面向使用者的文档。维护者私有的发布台账、本机运行报告和客户资料不会放入公开仓库。

Windows 桌面版 Excel 用于验证 Excel COM、VBA 执行、Power Query 刷新和 Power Pivot/Data Model 运行时行为。macOS 和 Linux 支持结构校验、OpenXML 检查、文档检查和非 COM 脚本。

## 校验方式

公开校验：

```bash
python tools/validate-skills.py .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
node tools/install.mjs --check
```

完整运行时校验，仅适用于 Windows 桌面版 Excel：

```powershell
python tools\run_release_gate.py --project-root .
```

## 相关页面

- [English project overview](project.en-US.md)
- [英文站点](intro.html)
- [中文站点](intro.zh-CN.html)
- [安装与同步说明](install-and-sync.md)
- [兼容性边界](compatibility.md)
