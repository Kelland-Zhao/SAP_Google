param(
    [Parameter(Mandatory = $true)]
    [string]$Message
)

$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

git add -A

$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "No staged changes. Nothing to commit." -ForegroundColor Yellow
    exit 0
}

Write-Host "Staged files:" -ForegroundColor Cyan
$staged | ForEach-Object { Write-Host "  $_" }

git commit -m $Message
git push origin main

Write-Host "Pushed: $Message" -ForegroundColor Green
