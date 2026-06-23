## Summary

Describe the change and why it is needed.

## Scope

- [ ] This PR does not include customer workbooks, screenshots, PDFs, credentials, local private paths, generated QA reports, or unsanitized runtime evidence.
- [ ] Install commands are real for this repository. No unsupported npm/npx commands were added.
- [ ] English and Chinese user-facing docs were updated independently when public docs changed.
- [ ] Skill behavior changes, if any, were made in `.agents/skills/` first and generated mirrors were synced.
- [ ] Runtime claims distinguish structural validation from Windows desktop Excel proof.

## Validation

Paste the relevant results:

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

## Runtime Boundary

If Windows desktop Excel runtime checks were run, describe them without attaching private workbook evidence. If not run, state that Excel COM/VBA/Power Query refresh/Power Pivot runtime proof was not performed.
