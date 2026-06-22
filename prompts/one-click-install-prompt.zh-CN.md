# 一键安装 Prompt

把下面这段完整复制给 Codex、Claude、OpenCode 或其他具备本地文件操作能力的 agent。

```text
你是本机插件安装助手。请帮我安装 Microsoft Excel BI Agent Pack。

目标：
1. 安装 Codex 插件。
2. 同步 Codex / Claude / OpenCode 的 skills。
3. 运行基础验证。
4. 汇总安装结果和失败原因。

插件位置：
请使用当前打开的插件目录；如果我提供的是 zip，请先解压，然后进入解压后的 microsoft_excel_bi_agent_pack 目录。

必须遵守：
- 不要手动编辑 marketplace.json。
- 不要手动编辑 skills/、.claude/skills/、.opencode/skills/。
- 不要删除用户已有的 ~/.codex/skills 里的其他技能。
- 不要把客户工作簿、截图、PDF、凭证、私有路径写入插件目录。
- 如果不是 Windows + 桌面版 Excel，不要声称已经验证 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行。

安装命令：
Windows PowerShell:
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

Git Bash / macOS / Linux:
python tools/deploy-local-plugin.py --project-root . --replace --install
python tools/sync-skills.py --project-root . --all-project-mirrors --codex-user --replace

基础验证：
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --check-drift
python tools\validate_project_docs.py --project-root .
python tools\build_artifact_hygiene_report.py --project-root . --require-pass
python tools\run_release_gate.py --project-root . --profile structural

如果当前环境是 Windows 且安装了桌面版 Excel，再运行完整验证：
python tools\run_release_gate.py --project-root .

输出要求：
- 告诉我插件是否安装成功。
- 告诉我 Codex / Claude / OpenCode skills 是否同步成功。
- 告诉我 structural release gate 是否通过。
- 如果跳过完整 Excel 验证，请明确说明跳过原因。
- 如果失败，请给出具体失败命令、错误摘要和下一步修复建议。
```

## 最短版

```text
请进入 microsoft_excel_bi_agent_pack 目录，执行安装、skills 同步和 structural 验证：
python tools\deploy-local-plugin.py --project-root . --replace --install
python tools\sync-skills.py --project-root . --all-project-mirrors --codex-user --replace
python tools\run_release_gate.py --project-root . --profile structural
完成后汇总结果；不要手改 marketplace.json 或 generated skills mirrors。
```
