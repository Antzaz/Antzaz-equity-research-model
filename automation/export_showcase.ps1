param(
    [string]$Destination = "$HOME\Documents\Antzaz-investment-research-showcase"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $Root "showcase"
$SnapshotBuilder = Join-Path $Root "automation\build_showcase_snapshot.py"
$PortfolioSummary = Join-Path $Root "institutional_research\outputs\latest\summary.json"

if (-not (Test-Path $Source)) {
    throw "Showcase source folder not found: $Source"
}

# If real portfolio outputs exist, refresh the sanitized public snapshot before export.
if (Test-Path $PortfolioSummary) {
    Write-Host "Building sanitized snapshot from real portfolio analytics..."
    python $SnapshotBuilder
    if ($LASTEXITCODE -ne 0) {
        throw "Sanitized portfolio snapshot build failed."
    }
}
else {
    Write-Warning "No local portfolio outputs found. Showcase will use its fallback demo portfolio until run_research.py is executed."
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
Write-Host "Portfolio data policy: real aggregate analytics; holdings anonymized; no tickers/cost basis/portfolio value exported."
Write-Host "Run locally:"
Write-Host "  cd `"$Destination`"; python -m pip install -r .\requirements.txt; python -m streamlit run .\app.py"
Write-Host ""
Write-Host "If GitHub CLI is installed and authenticated, publish with:"
Write-Host "  cd `"$Destination`"; gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push"
