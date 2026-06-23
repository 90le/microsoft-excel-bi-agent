# Contributing

Thanks for improving Microsoft Excel BI Agent. This project handles workflows around real Excel BI workbooks, so contribution safety matters as much as code quality.

## Before You Open an Issue

- Do not upload customer workbooks, screenshots, PDFs, credentials, local private paths, generated QA reports, or unsanitized runtime evidence.
- Use a sanitized fixture or a small synthetic example when possible.
- State whether your evidence is structural-only or was verified with Windows desktop Excel.
- Use the issue forms so the report lands in the right workflow.

## Before You Open a Pull Request

- Keep English and Chinese user-facing docs independently maintained.
- Keep install commands real. Do not add npm or npx commands unless the package is actually published there.
- Edit `.agents/skills/` first for skill behavior changes, then sync generated mirrors.
- Do not manually change generated skill mirrors without changing the source skill.
- Keep private workbooks, generated reports, local paths, and runtime evidence outside the public repository.

## Required Public Checks

Run these before submitting changes that affect docs, install flow, validation, packaging, fixtures, or GitHub community files:

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

Full Excel runtime proof requires Windows desktop Excel:

```powershell
python tools\run_release_gate.py --project-root .
```

## Review Standard

A contribution is ready when it is narrow, validated, does not leak private artifacts, and states the runtime boundary clearly.

Maintainer: **Qiu Binbin (丘彬彬)**. WeChat: **binstudy**. Blog: **https://90le.cn**.
