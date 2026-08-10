param(
    [string]$PortfolioPath = "$HOME\Documents\Antzaz-equity-research-model\institutional_research\portfolio.csv",
    [string]$Repository = "Antzaz/Antzaz-equity-research-model",
    [string]$SecUserAgent = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot

function Escape-Toml([string]$Value) {
    return $Value.Replace("\", "\\").Replace('"', '\"')
}

function New-RandomPassword {
    $bytes = New-Object byte[] 36
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($bytes)
    return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is required. Install it, run 'gh auth login', then rerun this script."
}

Write-Host "Checking GitHub CLI authentication..."
gh auth status | Out-Host

if (-not (Test-Path $PortfolioPath)) {
    throw "Portfolio file not found: $PortfolioPath"
}

$token = (gh auth token).Trim()
if (-not $token) {
    throw "GitHub CLI did not return an authentication token. Run 'gh auth login' first."
}

$portfolioBytes = [System.IO.File]::ReadAllBytes($PortfolioPath)
$portfolioB64 = [Convert]::ToBase64String($portfolioBytes)
$bundlePassword = New-RandomPassword

Write-Host "Setting repository Actions secrets..."
$portfolioB64 | gh secret set PORTFOLIO_CSV_B64 --repo $Repository
$bundlePassword | gh secret set LIVE_BUNDLE_PASSWORD --repo $Repository
if ($SecUserAgent) {
    $SecUserAgent | gh secret set SEC_USER_AGENT --repo $Repository
}

$streamlitDir = Join-Path $Root ".streamlit"
New-Item -ItemType Directory -Force -Path $streamlitDir | Out-Null
$secretsPath = Join-Path $streamlitDir "secrets.toml"
$hostedPath = Join-Path $streamlitDir "hosted_secrets_to_paste.toml"

$repoEsc = Escape-Toml $Repository
$tokenEsc = Escape-Toml $token
$passwordEsc = Escape-Toml $bundlePassword
$secrets = @"
[live_data]
repository = "$repoEsc"
github_token = "$tokenEsc"
bundle_password = "$passwordEsc"
workflow_file = "daily-portfolio-refresh.yml"
"@

$secrets | Set-Content -Encoding UTF8 $secretsPath
$secrets | Set-Content -Encoding UTF8 $hostedPath

Write-Host "Triggering first Daily private portfolio refresh workflow..."
gh workflow run daily-portfolio-refresh.yml --repo $Repository --ref main | Out-Host

Write-Host ""
Write-Host "Live dashboard GitHub setup is configured."
Write-Host "Local Streamlit secrets: $secretsPath"
Write-Host "Hosted Streamlit secrets to paste: $hostedPath"
Write-Host ""
Write-Host "The hosted secrets file contains credentials. Do not share or commit it."
Write-Host "For Streamlit Community Cloud, open App settings -> Secrets and paste the contents of hosted_secrets_to_paste.toml."
Write-Host ""
Write-Host "Check the workflow with:"
Write-Host "  gh run list --repo $Repository --workflow daily-portfolio-refresh.yml --limit 3"
Write-Host ""
Write-Host "Launch the private portal locally with:"
Write-Host "  cd `"$Root`"; python -m pip install -r .\institutional_research\requirements.txt; python -m streamlit run .\institutional_research\live_dashboard.py"
