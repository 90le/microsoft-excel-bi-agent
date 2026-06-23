# Microsoft Excel BI Agent 使用说明

本说明面向收到插件包的使用者。拿到完整文件夹或 zip 后，按下面步骤安装即可。

## 这个包是什么

Microsoft Excel BI Agent 是一套给 AI agent 使用的 Excel BI 工作技能包。它把 Excel/VBA、Power Query M、Power Pivot DAX、MDX/CUBE、ADO/SQL、交付物清理、工作簿 QA、Office 环境诊断等常见工作流程整理成标准技能和脚本。

适用对象：

- 使用 Codex 处理 Excel 自动化、VBA、Power Query、DAX、CUBE 公式的人。
- 希望 Claude / OpenCode 也能沿用同一套 Excel 工作规范的人。
- 需要把 Excel 交付物做成可复用、可验证、可审计流程的团队。

## 安装前准备

- 已安装 Python 3.10 或更高版本。
- 如果要验证 Excel COM、VBA、Power Query 刷新、Power Pivot/Data Model，请使用 Windows，并安装桌面版 Microsoft Excel。
- macOS / Linux 可以安装 skills、运行结构检查和 OpenXML 类工具，但不能证明 Excel 桌面运行能力。

## 一键安装

在插件目录下执行：

```powershell
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

Git Bash / macOS / Linux：

```bash
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
```

## 安装后验证

```powershell
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
python tools\run_release_gate.py --project-root . --profile structural
```

如果需要验证 Windows Excel 真实运行能力：

```powershell
python tools\run_release_gate.py --project-root .
```

## 安装结果

安装完成后通常会有这些结果：

- Codex 插件安装到本机 personal marketplace。
- Codex 用户 skills 同步到 `~/.codex/skills`。
- Claude 项目 skills 同步到 `.claude/skills`。
- OpenCode 项目 skills 同步到 `.opencode/skills`。
- Codex 插件镜像同步到 `skills/`。

## 常见问题

### 提示 Python 找不到

先安装 Python，并确认命令行能执行：

```powershell
python --version
```

### 结构检查通过，但 Excel 刷新或 VBA 没有验证

这是正常边界。结构检查不等于 Excel 桌面运行验证。需要 Windows + 桌面版 Excel 才能运行完整门禁。

### 不要手改哪些目录

不要手动改：

- `skills/`
- `.claude/skills/`
- `.opencode/skills/`
- `~/.codex/skills/` 中本插件对应的 12 个 skill 文件夹

它们应从 `.agents/skills/` 统一同步生成。

## 推荐给 AI agent 的使用方式

如果让另一个 agent 帮你安装，直接把下面文件内容复制给它：

```text
prompts/one-click-install-prompt.zh-CN.md
```

如果要快速了解这个包的能力，可以打开：

```text
docs/intro.html
```
