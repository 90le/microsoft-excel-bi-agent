# 一键安装 Prompt

把下面这段完整复制给 Codex、Claude、OpenCode 或其他具备本地文件操作能力的 Agent。

```text
你是本机插件安装助手。请帮我安装 Microsoft Excel BI Agent。

目标：
1. 安装或刷新 Codex 插件。
2. 同步 Codex / Claude / OpenCode 的 skills。
3. 运行基础结构校验。
4. 汇总安装结果、跳过项和失败原因。

插件位置：
请使用当前打开的 microsoft-excel-bi-agent 目录；如果我提供的是 zip，请先解压，然后进入解压后的 microsoft-excel-bi-agent 目录。

优先安装命令：
node tools/install.mjs

如果 Node 不可用，使用手动命令。

Windows PowerShell:
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

Git Bash / macOS / Linux:
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

基础校验：
python tools/validate-skills.py .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
node tools/install.mjs --check

如果当前环境是 Windows 且安装了桌面版 Excel，再运行完整校验：
python tools\run_release_gate.py --project-root .

必须遵守：
- 不要手动编辑 generated mirrors：skills/、.claude/skills/、.opencode/skills/。
- 不要删除用户已有 ~/.codex/skills 里的其他技能。
- 不要把客户工作簿、截图、PDF、凭证、私有路径写入插件目录。
- 如果不是 Windows + 桌面版 Excel，不要声称已经验证 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时。
- 如果完整 Excel 校验被跳过，请明确说明原因。

输出要求：
- 说明插件安装是否成功。
- 说明 Codex / Claude / OpenCode skills 是否同步成功。
- 说明公开结构校验是否通过。
- 如果失败，给出失败命令、错误摘要和下一步修复建议。
```

## 最短版

```text
请进入 microsoft-excel-bi-agent 目录，执行：
node tools/install.mjs
node tools/install.mjs --check
完成后汇总结果；不要手动编辑 generated skills mirrors。
```
