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

# Refresh the sanitized aggregate analytics when local production outputs exist.
if (Test-Path $PortfolioSummary) {
    Write-Host "Building sanitized snapshot from portfolio analytics..."
    python $SnapshotBuilder
    if ($LASTEXITCODE -ne 0) {
        throw "Sanitized portfolio snapshot build failed."
    }
}

if (Test-Path $Destination) {
    Remove-Item $Destination -Recurse -Force
}
New-Item -ItemType Directory -Path $Destination | Out-Null
New-Item -ItemType Directory -Path (Join-Path $Destination "data") | Out-Null

Copy-Item (Join-Path $Source "app.py") $Destination
Copy-Item (Join-Path $Source "requirements.txt") $Destination
Copy-Item (Join-Path $Source "README.md") $Destination
Copy-Item (Join-Path $Source "EDIT_PORTFOLIO.md") $Destination
Copy-Item (Join-Path $Source "data\recruiter_portfolio.json") (Join-Path $Destination "data\recruiter_portfolio.json")
Copy-Item (Join-Path $Source "data\recruiter_portfolio.example.json") (Join-Path $Destination "data\recruiter_portfolio.example.json")

$Snapshot = Join-Path $Source "data\portfolio_snapshot.json"
if (Test-Path $Snapshot) {
    Copy-Item $Snapshot (Join-Path $Destination "data\portfolio_snapshot.json")
}

@"
__pycache__/
.streamlit/secrets.toml
.env
*.xlsx
*.zip
"@ | Set-Content -Encoding UTF8 (Join-Path $Destination ".gitignore")

Push-Location $Destination
try {
    if (-not (Test-Path ".git")) {
        git init | Out-Null
        git branch -M main
    }
    git add app.py requirements.txt README.md EDIT_PORTFOLIO.md .gitignore data
    $changes = git status --porcelain
    if ($changes) {
        git commit -m "Refresh recruiter investment showcase" | Out-Null
    }
}
finally {
    Pop-Location
}

Write-Host "Recruiter showcase prepared at: $Destination"
Write-Host "Edit public theses and values in data\recruiter_portfolio.json."
Write-Host "Run locally:"
Write-Host "  cd `"$Destination`"; python -m pip install -r .\requirements.txt; python -m streamlit run .\app.py"
