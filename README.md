# Microsoft Excel BI Agent Pack

这是一个可分发的 Excel BI agent 技能包，用于让 Codex、Claude、OpenCode 等 agent 更稳定地处理 Excel 工作簿、VBA、Power Query M、Power Pivot DAX、MDX/CUBE 公式、ADO/SQL、交付物清理和 workbook QA。

当前版本：`0.1.0+codex.20260622060709`

License: MIT

GitHub: https://github.com/90le/microsoft-excel-bi-agent

## 适用场景

- Excel/VBA 工作簿创建、修改、调试、按钮绑定、隐藏 sheet、公式和链接检查。
- Power Query M 查询读取、编辑、刷新、错误定位和性能记录。
- Power Pivot/Data Model/DAX/MDX/CUBE 公式分析。
- Excel 交付物转纯净版、删除外链、删除查询、冻结公式为值。
- 多 agent 共用同一套 Excel BI 操作规范。

## 安装方式

把整个 `microsoft_excel_bi_agent_pack` 文件夹解压到本机任意目录，然后在该目录下执行。

安装为 Codex 插件：

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
```

同步到 Codex / Claude / OpenCode skills：

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

macOS / Linux / Git Bash 使用斜杠路径：

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## 验证安装

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
python tools\run_release_gate.py --project-root . --profile structural
```

如果需要验证 Windows Excel COM、VBA、Power Query 刷新、Power Pivot/Data Model 运行能力，在 Windows 且已安装桌面版 Excel 的环境中执行完整门禁：

```powershell
python tools\run_release_gate.py --project-root .
```

## 使用入口

- 状态总览：`docs/current-status.md`
- 安装同步规则：`docs/install-and-sync.md`
- 任务配方：`docs/task-recipes.md`
- 分发前检查：`docs/distribution-checklist.md`
- 中文使用说明：`docs/recipient-guide.zh-CN.md`
- 一键安装 Prompt：`prompts/one-click-install-prompt.zh-CN.md`
- HTML 介绍页：`docs/intro.html`
- 开源发布说明：`docs/open-source-publishing.md`
- 发布到 GitHub：`docs/github-publish.md`

## 重要边界

- 不要手动修改 `skills/`、`.claude/skills/`、`.opencode/skills/`，这些目录由 `.agents/skills/` 同步生成。
- 不要把客户工作簿、截图、PDF、凭证、机器路径、生成报告放进插件包。
- macOS / Linux 可以做结构检查和 prompt/skill 同步，但不能证明 Excel COM、VBA、Power Query 刷新、Power Pivot 运行结果。
- 行为变更且需要让 Codex 识别新版本时，才使用 cachebuster 安装命令：

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install --update-cachebuster
```
