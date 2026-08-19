param(
    [string]$Destination = "$HOME\Documents\Antzaz-investment-research-showcase",
    [switch]$AllowDemo
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $Root "showcase"
$SnapshotBuilder = Join-Path $Root "automation\build_showcase_snapshot.py"
$SnapshotValidator = Join-Path $Root "automation\validate_showcase.py"
$PortfolioSummary = Join-Path $Root "institutional_research\outputs\latest\summary.json"
$ThesisWorkbook = Join-Path $Root "institutional_research\portfolio_thesis.xlsx"
$PublicThesisJson = Join-Path $Root "institutional_research\portfolio_thesis_public.json"

if (-not (Test-Path $Source)) {
    throw "Showcase source folder not found: $Source"
}

# Build the recruiter-safe thesis JSON locally whenever the private thesis workbook exists.
# This keeps local showcase exports consistent with the GitHub Actions publishing path.
if (Test-Path $ThesisWorkbook) {
    Write-Host "Building recruiter-safe investment thesis payload from local workbook..."
    $env:THESIS_WORKBOOK = $ThesisWorkbook
    $env:THESIS_PUBLIC_JSON = $PublicThesisJson
    @'
from pathlib import Path
import json
import os
from automation.sync_portfolio_thesis import build_payload

source = Path(os.environ["THESIS_WORKBOOK"])
destination = Path(os.environ["THESIS_PUBLIC_JSON"])
payload = build_payload(source)
destination.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
print(f"Published company theses prepared locally: {len(payload.get('company_theses', []))}")
print(f"Portfolio philosophy fields prepared locally: {len(payload.get('portfolio_philosophy', {}))}")
'@ | python -
    if ($LASTEXITCODE -ne 0) {
        throw "Recruiter-safe investment thesis payload build failed."
    }
}
else {
    Write-Host "No local portfolio_thesis.xlsx found; exporting portfolio analytics without recruiter thesis content."
}

# Recruiter-safe default: require a real sanitized portfolio snapshot before export.
if (Test-Path $PortfolioSummary) {
    Write-Host "Building sanitized snapshot from real portfolio analytics..."
    python $SnapshotBuilder
    if ($LASTEXITCODE -ne 0) {
        throw "Sanitized portfolio snapshot build failed."
    }

    python $SnapshotValidator
    if ($LASTEXITCODE -ne 0) {
        throw "Recruiter showcase validation failed."
    }
}
elseif (-not $AllowDemo) {
    throw "No local portfolio outputs found. Run institutional_research\run_research.py first. Use -AllowDemo only for a non-recruiter demo export."
}
else {
    Write-Warning "No local portfolio outputs found. Exporting the illustrative demo version because -AllowDemo was supplied."
}

if (Test-Path $Destination) {
    Remove-Item $Destination -Recurse -Force
}
New-Item -ItemType Directory -Path $Destination | Out-Null

Copy-Item (Join-Path $Source "app.py") $Destination
Copy-Item (Join-Path $Source "requirements.txt") $Destination
Copy-Item (Join-Path $Source "README.md") $Destination

$Snapshot = Join-Path $Source "data\portfolio_snapshot.json"
if (Test-Path $Snapshot) {
    $DataDest = Join-Path $Destination "data"
    New-Item -ItemType Directory -Path $DataDest | Out-Null
    Copy-Item $Snapshot (Join-Path $DataDest "portfolio_snapshot.json")
}

@"
__pycache__/
.streamlit/secrets.toml
.env
*.xlsx
*.csv
*.zip
"@ | Set-Content -Encoding UTF8 (Join-Path $Destination ".gitignore")

Push-Location $Destination
try {
    if (-not (Test-Path ".git")) {
        git init | Out-Null
        git branch -M main
    }

    # Configure a repository-local Git identity only when one is not already set.
    # This avoids changing the user's global Git configuration.
    $gitName = git config user.name 2>$null
    if ([string]::IsNullOrWhiteSpace($gitName)) {
        git config user.name "Antzaz"
    }
    $gitEmail = git config user.email 2>$null
    if ([string]::IsNullOrWhiteSpace($gitEmail)) {
        git config user.email "31651836+Antzaz@users.noreply.github.com"
    }

    git add app.py requirements.txt README.md .gitignore
    if (Test-Path ".\data\portfolio_snapshot.json") {
        git add .\data\portfolio_snapshot.json
    }
    $changes = git status --porcelain
    if ($changes) {
        git commit -m "Refresh investment research showcase" | Out-Null
    }
}
finally {
    Pop-Location
}

Write-Host "Public showcase prepared at: $Destination"
if (Test-Path $Snapshot) {
    Write-Host "Recruiter portfolio status: VALIDATED sanitized real analytics included."
}
else {
    Write-Host "Recruiter portfolio status: DEMO ONLY (no real sanitized snapshot included)."
}
Write-Host "Portfolio data policy: real aggregate analytics; holdings anonymized; no tickers/cost basis/portfolio value exported."
Write-Host "Run locally:"
Write-Host "  cd `"$Destination`"; python -m pip install -r .\requirements.txt; python -m streamlit run .\app.py"
Write-Host ""
Write-Host "Initial public GitHub publish (requires GitHub CLI):"
Write-Host "  cd `"$Destination`"; gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push"
