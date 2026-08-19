param(
    [string]$PortfolioPath = "",
    [string]$ThesisPath = "",
    [string]$Repository = "Antzaz/Antzaz-equity-research-model",
    [switch]$RunRefresh,
    [switch]$NoRefresh
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($PortfolioPath)) {
    $PortfolioPath = Join-Path $Root "institutional_research\portfolio.csv"
}
if ([string]::IsNullOrWhiteSpace($ThesisPath)) {
    $ThesisPath = Join-Path $Root "institutional_research\portfolio_thesis.xlsx"
}

if (-not (Test-Path $PortfolioPath)) {
    throw "Portfolio file not found: $PortfolioPath"
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is not installed or is not on PATH. Install/authenticate gh first."
}

gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run: gh auth login"
}

$csvLines = @(
    Get-Content -Path $PortfolioPath |
        Where-Object {
            $trimmed = $_.TrimStart()
            -not [string]::IsNullOrWhiteSpace($_) -and -not $trimmed.StartsWith("#")
        }
)

if (-not $csvLines -or $csvLines.Count -lt 2) {
    throw "Portfolio CSV contains no holdings after comment/template lines are ignored."
}

$rows = @($csvLines | ConvertFrom-Csv)
if (-not $rows -or $rows.Count -lt 1) {
    throw "Portfolio CSV contains no holdings."
}

$columns = @($rows[0].PSObject.Properties.Name)
if ($columns -notcontains "Ticker") {
    throw "Portfolio CSV must contain a Ticker column. Detected columns: $($columns -join ', ')"
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

$bytes = [System.IO.File]::ReadAllBytes((Resolve-Path $PortfolioPath))
$b64 = [Convert]::ToBase64String($bytes)

$b64 | gh secret set PORTFOLIO_CSV_B64 --repo $Repository
if ($LASTEXITCODE -ne 0) {
    throw "Failed to update PORTFOLIO_CSV_B64 in $Repository."
}

Write-Host "Local portfolio synced to the cloud source successfully."
Write-Host "Repository: $Repository"
Write-Host "Portfolio file: $PortfolioPath"
Write-Host "Holdings synced: $($tickers.Count)"
Write-Host "Comment/template lines were ignored during validation."
Write-Host "Tickers are not printed to avoid unnecessary disclosure in shared terminal output."

if (Test-Path $ThesisPath) {
    Write-Host "Syncing recruiter-facing investment thesis workbook..."
    python (Join-Path $Root "automation\sync_portfolio_thesis.py") --file $ThesisPath --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Portfolio was synced, but investment thesis sync failed."
    }
}
else {
    Write-Host "No portfolio_thesis.xlsx found; portfolio analytics will refresh without recruiter thesis content."
}

$shouldRefresh = -not $NoRefresh
if ($RunRefresh) {
    $shouldRefresh = $true
}

if ($shouldRefresh) {
    Write-Host "Triggering fast recruiter portfolio refresh..."
    gh workflow run "Recruiter portfolio refresh" --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Portfolio source was updated, but recruiter refresh dispatch failed. Trigger 'Recruiter portfolio refresh' manually in GitHub Actions."
    }

    Write-Host "Triggering independent public fundamentals refresh..."
    gh workflow run "Public fundamentals refresh" --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Portfolio/thesis source was updated, but public fundamentals refresh dispatch failed. Trigger 'Public fundamentals refresh' manually in GitHub Actions."
    }

    Write-Host "Recruiter analytics and public fundamentals refreshes dispatched."
}
else {
    Write-Host "Cloud portfolio/thesis source updated without triggering an immediate recruiter refresh."
}
