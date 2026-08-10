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
    # Windows PowerShell 5.1 runs on .NET Framework, where the static
    # RandomNumberGenerator.Fill() API is unavailable. Use the instance API so
    # this works in both Windows PowerShell 5.1 and modern PowerShell/.NET.
    $bytes = New-Object byte[] 36
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }
    return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}

function Resolve-GitHubCli {
    $cmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $known = @(
        "$env:ProgramFiles\GitHub CLI\gh.exe",
        "$env:LOCALAPPDATA\Programs\GitHub CLI\gh.exe"
    )
    foreach ($candidate in $known) {
        if (Test-Path $candidate) { return $candidate }
    }
    return $null
}

$GhPath = Resolve-GitHubCli
if (-not $GhPath) {
    Write-Host "GitHub CLI is not installed. Installing the official GitHub CLI package with WinGet..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "WinGet is unavailable. Install GitHub CLI from https://cli.github.com/ and rerun this script."
    }
    & winget install --id GitHub.cli --source winget --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI installation failed with exit code $LASTEXITCODE."
    }
    $GhPath = Resolve-GitHubCli
    if (-not $GhPath) {
        throw "GitHub CLI installed but gh.exe could not be located. Open a new PowerShell window and rerun this script."
    }
    Write-Host "GitHub CLI installed: $GhPath"
}

# Make gh available to child commands in this PowerShell process even when the installer
# has not yet propagated PATH changes to the current terminal.
$GhDir = Split-Path -Parent $GhPath
if (($env:Path -split ';') -notcontains $GhDir) {
    $env:Path = "$GhDir;$env:Path"
}

Write-Host "Checking GitHub CLI authentication..."
# A logged-out `gh auth status` returns a non-zero exit code. That is expected during
# first-time setup, so temporarily avoid treating native stderr as a terminating error.
$oldPreference = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    & $GhPath auth status *> $null
    $authOk = ($LASTEXITCODE -eq 0)
}
finally {
    $ErrorActionPreference = $oldPreference
}

if (-not $authOk) {
    Write-Host "GitHub CLI needs authorization. A browser/device login will start now."
    Write-Host "Choose GitHub.com and complete the browser login for your Antzaz account."
    & $GhPath auth login --hostname github.com --git-protocol https --web
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI authentication did not complete successfully."
    }
}

# Verify authentication after login using the resolved executable path.
$oldPreference = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    & $GhPath auth status
    $authOk = ($LASTEXITCODE -eq 0)
}
finally {
    $ErrorActionPreference = $oldPreference
}
if (-not $authOk) {
    throw "GitHub CLI is installed but authentication is still unavailable. Rerun this script and complete the browser login."
}

if (-not (Test-Path $PortfolioPath)) {
    throw "Portfolio file not found: $PortfolioPath"
}

$token = (& $GhPath auth token).Trim()
if (-not $token) {
    throw "GitHub CLI did not return an authentication token. Rerun this script and complete GitHub authentication."
}

$portfolioBytes = [System.IO.File]::ReadAllBytes($PortfolioPath)
$portfolioB64 = [Convert]::ToBase64String($portfolioBytes)
$bundlePassword = New-RandomPassword

Write-Host "Setting repository Actions secrets..."
$portfolioB64 | & $GhPath secret set PORTFOLIO_CSV_B64 --repo $Repository
if ($LASTEXITCODE -ne 0) { throw "Failed to set PORTFOLIO_CSV_B64." }
$bundlePassword | & $GhPath secret set LIVE_BUNDLE_PASSWORD --repo $Repository
if ($LASTEXITCODE -ne 0) { throw "Failed to set LIVE_BUNDLE_PASSWORD." }
if ($SecUserAgent) {
    $SecUserAgent | & $GhPath secret set SEC_USER_AGENT --repo $Repository
    if ($LASTEXITCODE -ne 0) { throw "Failed to set SEC_USER_AGENT." }
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
& $GhPath workflow run daily-portfolio-refresh.yml --repo $Repository --ref main
if ($LASTEXITCODE -ne 0) {
    throw "Could not trigger daily-portfolio-refresh.yml. Check that Actions are enabled for the repository."
}

Write-Host ""
Write-Host "Live dashboard GitHub setup is configured."
Write-Host "Local Streamlit secrets: $secretsPath"
Write-Host "Hosted Streamlit secrets to paste: $hostedPath"
Write-Host ""
Write-Host "The hosted secrets file contains credentials. Do not share or commit it."
Write-Host "For Streamlit Community Cloud, open App settings -> Secrets and paste the contents of hosted_secrets_to_paste.toml."
Write-Host ""
Write-Host "Check the workflow with:"
Write-Host "  & `"$GhPath`" run list --repo $Repository --workflow daily-portfolio-refresh.yml --limit 3"
Write-Host ""
Write-Host "Launch the private portal locally with:"
Write-Host "  cd `"$Root`"; python -m pip install -r .\institutional_research\requirements.txt; python -m streamlit run .\institutional_research\live_dashboard.py"
