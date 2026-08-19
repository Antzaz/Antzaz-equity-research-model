param(
    [string]$PortfolioPath = "",
    [string]$Repository = "Antzaz/Antzaz-equity-research-model",
    [switch]$RunRefresh
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($PortfolioPath)) {
    $PortfolioPath = Join-Path $Root "institutional_research\portfolio.csv"
}

if (-not (Test-Path $PortfolioPath)) {
    throw "Portfolio file not found: $PortfolioPath"
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is not installed or is not on PATH. Install/authenticate gh first."
}

# Fail early if GitHub CLI is not authenticated.
gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run: gh auth login"
}

# Basic CSV validation before replacing the cloud secret.
$rows = Import-Csv $PortfolioPath
if (-not $rows -or $rows.Count -lt 1) {
    throw "Portfolio CSV contains no holdings."
}

$columns = @($rows[0].PSObject.Properties.Name)
if ($columns -notcontains "Ticker") {
    throw "Portfolio CSV must contain a Ticker column."
}

$tickers = @(
    $rows |
        ForEach-Object { [string]$_.Ticker } |
        ForEach-Object { $_.Trim().ToUpperInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
)

if ($tickers.Count -lt 1) {
    throw "Portfolio CSV contains no valid ticker values."
}

$duplicates = @($tickers | Group-Object | Where-Object Count -gt 1 | Select-Object -ExpandProperty Name)
if ($duplicates.Count -gt 0) {
    throw "Duplicate portfolio tickers detected: $($duplicates -join ', ')"
}

# Encode without printing the base64 value to the terminal.
$bytes = [System.IO.File]::ReadAllBytes((Resolve-Path $PortfolioPath))
$b64 = [Convert]::ToBase64String($bytes)

# gh secret set reads the value from stdin. We intentionally never Write-Host the value.
$b64 | gh secret set PORTFOLIO_CSV_B64 --repo $Repository
if ($LASTEXITCODE -ne 0) {
    throw "Failed to update PORTFOLIO_CSV_B64 in $Repository."
}

Write-Host "Cloud portfolio secret updated successfully."
Write-Host "Repository: $Repository"
Write-Host "Portfolio file: $PortfolioPath"
Write-Host "Holdings synced: $($tickers.Count)"
Write-Host "Tickers are not printed to avoid unnecessary disclosure in shared terminal output."

if ($RunRefresh) {
    Write-Host "Triggering Daily private portfolio refresh..."
    gh workflow run "Daily private portfolio refresh" --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Portfolio secret was updated, but workflow dispatch failed. Trigger the workflow manually in GitHub Actions."
    }
    Write-Host "Refresh workflow dispatched."
}
else {
    Write-Host "Next: run the 'Daily private portfolio refresh' workflow, or rerun this command with -RunRefresh."
}
