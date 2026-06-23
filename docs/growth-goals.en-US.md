# Public Growth Goals

This page defines the public growth target for Microsoft Excel BI Agent. It is intentionally separate from the maintenance goals: maintenance keeps the package safe; growth makes the value easier to understand, trust, and install.

Maintainer: **Qiu Binbin (丘彬彬)**<br>
WeChat: **binstudy**<br>
Blog: **https://90le.cn**

## Objective

Make the repository credible within the first minute: a visitor should understand what the pack does, when to use it, why Excel BI work is risky for generic agents, how to install it, who maintains it, and what validation boundaries still apply.

The goal is not broad advertising. The goal is higher-trust adoption for the exact users who work with real Excel BI workbooks, workbook delivery risk, and AI-agent automation.

## Constraints

- Keep install claims true. Do not advertise npm or npx until a package is actually published.
- Keep English and Chinese public pages independently maintained.
- Do not use customer files, screenshots, workbook reports, private paths, credentials, or runtime evidence as marketing material.
- Prefer deterministic repository assets for text-heavy visuals. Use generated images only when they add real comprehension and do not risk incorrect readable text.
- Do not imply that structural validation proves Excel COM, VBA, Power Query refresh, or Power Pivot runtime behavior.
- Keep the website responsive and readable on desktop and mobile.

## Boundaries

- Public growth work may improve README copy, Pages layout, social preview metadata, screenshots, release visibility, use-case framing, and advertising copy.
- Public growth work must not alter Excel workbook behavior unless a separate product-quality goal justifies it.
- The repository can show sanitized examples and workflow shapes, but not private workbook evidence.
- The project can include contact and maintainer signature, but should not turn the README into a personal landing page.

## Can Do

- Add clear positioning, use-case sections, and trust signals to README and Pages.
- Add author/maintainer signature and contact links.
- Add reusable launch, social, and ad copy that points to real install paths.
- Improve Open Graph/Twitter metadata with absolute public image URLs.
- Add validation checks that prevent stale version, missing author, or missing growth docs from shipping.
- Keep marketing language grounded in actual skills and validation boundaries.

## Cannot Do

- Add fake package-manager commands, fake benchmark claims, fake logos, or unsupported integrations.
- Use private customer workbook images or local runtime artifacts for promotion.
- Promise Excel runtime behavior without Windows desktop Excel evidence.
- Add decorative image-heavy pages that slow the site or obscure install commands.
- Mix long-form English and Chinese content into the same user-facing document.

## Detailed Goals

| Goal | Value | Done when |
| --- | --- | --- |
| First-minute clarity | New visitors decide quickly whether the project fits their Excel BI risk. | README and Pages state the audience, use cases, install path, and runtime boundary near the top. |
| Trust and attribution | Open-source users need to know who maintains the package. | README, Pages, and metadata show Qiu Binbin, WeChat `binstudy`, and blog `90le.cn`. |
| Social sharing | Shared links should render with the right preview image and description. | HTML pages use absolute Open Graph/Twitter image URLs and canonical page URLs. |
| Advertising readiness | Maintainers can promote the repo without rewriting copy every time. | A bilingual marketing copy pack provides launch copy, short ads, and channel-specific variants. |
| Visual richness without risk | The site should look like a real product page without inventing claims. | Existing deterministic assets are reused; generated imagery is optional, not required. |
| Validation-backed growth | Marketing improvements should not bypass quality gates. | Project docs validation and goal coverage include growth docs and signature checks. |

## High-Value Optimization Backlog

| Priority | Item | Why it is worth doing | Boundary |
| --- | --- | --- | --- |
| P0 | Add maintainer signature and contact. | Improves trust and ownership with almost no runtime risk. | Keep it concise; do not make the project personal-brand first. |
| P0 | Make social metadata absolute and current. | Broken previews reduce sharing value. | Use existing public assets; no private images. |
| P0 | Add public growth goals and ad copy. | Converts vague marketing work into a reusable, reviewable contract. | Copy must only mention real features and install paths. |
| P1 | Add a use-case conversion band to Pages. | Helps users map their workbook pain to the correct skill pack value. | Keep the install commands visible and responsive. |
| P1 | Validate growth docs in CI. | Prevents stale marketing claims from drifting after releases. | CI remains structural and does not require desktop Excel. |
| P2 | Add short demo GIF or video later. | Could improve comprehension for non-technical visitors. | Only with sanitized fixtures and no customer data. |

## Required Public Checks

Run these after changing README, website, marketing copy, release notes, install docs, or validation scripts:

```bash
python tools/validate-skills.py .
python tools/validate_project_docs.py --project-root .
python tools/validate_task_recipes.py --project-root .
python tools/validate_official_docs_index.py --project-root .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
python tools/build_goal_coverage_report.py --project-root . --require-pass
node tools/install.mjs --check
```
