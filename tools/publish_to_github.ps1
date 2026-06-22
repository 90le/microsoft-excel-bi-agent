param(
    [string]$Owner = "",
    [string]$Repo = "microsoft-excel-bi-agent",
    [switch]$Private
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI 'gh' is not installed or not on PATH. Install it from https://cli.github.com/ and run 'gh auth login'."
}

gh auth status | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run 'gh auth login' first."
}

if (-not (Test-Path ".git")) {
    git init -b main
}

$status = git status --porcelain
if ($status) {
    git add .
    git commit -m "Initial open source release"
}

if (-not $Owner) {
    $Owner = (gh api user --jq ".login").Trim()
}

$visibility = if ($Private) { "--private" } else { "--public" }
$repoFullName = "$Owner/$Repo"

$remote = git remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0 -or -not $remote) {
    gh repo create $repoFullName $visibility --source . --remote origin --push
} else {
    git push -u origin main
}

Write-Host "Published: https://github.com/$repoFullName"
