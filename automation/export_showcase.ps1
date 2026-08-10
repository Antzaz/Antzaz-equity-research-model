param(
    [string]$Destination = "$HOME\Documents\Antzaz-investment-research-showcase"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $Root "showcase"

if (-not (Test-Path $Source)) {
    throw "Showcase source folder not found: $Source"
}

if (Test-Path $Destination) {
    Remove-Item $Destination -Recurse -Force
}
New-Item -ItemType Directory -Path $Destination | Out-Null

Copy-Item (Join-Path $Source "app.py") $Destination
Copy-Item (Join-Path $Source "requirements.txt") $Destination
Copy-Item (Join-Path $Source "README.md") $Destination

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
    $changes = git status --porcelain
    if ($changes) {
        git commit -m "Create investment research showcase" | Out-Null
    }
}
finally {
    Pop-Location
}

Write-Host "Public showcase prepared at: $Destination"
Write-Host "Run locally:"
Write-Host "  cd `"$Destination`"; python -m pip install -r .\requirements.txt; python -m streamlit run .\app.py"
Write-Host ""
Write-Host "If GitHub CLI is installed and authenticated, publish with:"
Write-Host "  cd `"$Destination`"; gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push"
