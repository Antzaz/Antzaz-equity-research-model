param(
    [string]$PortfolioPath = "",
    [string]$Repository = "Antzaz/Antzaz-equity-research-model",
    [switch]$RunRefresh,
    [switch]$NoRefresh
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

# Parse the CSV the same way as institutional_research/src/portfolio.py:
# ignore blank lines and template/instruction lines beginning with '#'.
# Plain Import-Csv would otherwise treat the first comment line as the header,
# making a valid Ticker column on the real header row appear to be missing.
$csvLines = @(
    Get-Content -LiteralPath $PortfolioPath -Encoding UTF8 |
        Where-Object {
            $line = [string]$_
            -not [string]::IsNullOrWhiteSpace($line) -and
            -not $line.TrimStart().StartsWith("#")
        }
)

if ($csvLines.Count -lt 2) {
    throw "Portfolio CSV does not contain a header plus at least one holding after comment lines are ignored."
}

try {
    $rows = @($csvLines | ConvertFrom-Csv)
}
catch {
    throw "Portfolio CSV could not be parsed after comment lines were ignored: $($_.Exception.Message)"
}

if (-not $rows -or $rows.Count -lt 1) {
    throw "Portfolio CSV contains no holdings."
}

$columns = @($rows[0].PSObject.Properties.Name | ForEach-Object { ([string]$_).Trim() })
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

# Encode the ORIGINAL file, including its helpful comment lines. The Python portfolio
# loader intentionally supports those comments via pd.read_csv(..., comment='#').
# The base64 value is never printed to the terminal.
$bytes = [System.IO.File]::ReadAllBytes((Resolve-Path $PortfolioPath))
$b64 = [Convert]::ToBase64String($bytes)

# gh secret set reads the value from stdin.
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

# Refresh is the default behavior. -RunRefresh remains accepted for backwards
# compatibility; use -NoRefresh only when you intentionally want to update the cloud
# portfolio composition without immediately rebuilding the online analytics.
$shouldRefresh = -not $NoRefresh
if ($RunRefresh) {
    $shouldRefresh = $true
}

if ($shouldRefresh) {
    Write-Host "Triggering Daily private portfolio refresh..."
    gh workflow run "Daily private portfolio refresh" --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Portfolio source was updated, but workflow dispatch failed. Trigger the workflow manually in GitHub Actions."
    }
    Write-Host "Refresh workflow dispatched. GitHub will rebuild and publish the sanitized online portfolio snapshot."
}
else {
    Write-Host "Cloud portfolio source updated without triggering an immediate analytics refresh."
}
