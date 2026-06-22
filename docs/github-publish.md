# Publish To GitHub

The repository has already been prepared as a clean public staging copy.

## Prerequisites

- Git installed.
- GitHub CLI `gh` installed.
- Authenticated GitHub CLI session:

```powershell
gh auth login
gh auth status
```

## One-Command Publish

From the repository root:

```powershell
.\tools\publish_to_github.ps1 -Owner <github-owner> -Repo microsoft-excel-bi-agent
```

If `-Owner` is omitted, the script uses the authenticated GitHub user.

Default visibility is public. To create a private repo first:

```powershell
.\tools\publish_to_github.ps1 -Owner <github-owner> -Repo microsoft-excel-bi-agent -Private
```

## Manual Equivalent

```powershell
gh repo create <github-owner>/microsoft-excel-bi-agent --public --source . --remote origin --push
```

## Current Open-Source Boundary

This public staging repository intentionally excludes maintainer-only local release ledgers and machine-specific smoke outputs. It includes source skills, generated agent mirrors, tooling, fixtures, recipient docs, install prompts, and the MIT license.
