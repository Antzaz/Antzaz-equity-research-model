param(
    [string]$UserAgent = "",
    [string]$Repository = "Antzaz/Antzaz-equity-research-model",
    [switch]$RunLearning
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is not installed or is not on PATH. Install/authenticate gh first."
}

gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run: gh auth login"
}

if ([string]::IsNullOrWhiteSpace($UserAgent)) {
    $UserAgent = Read-Host "Enter SEC User-Agent (example: EquityResearch your-real-email@example.com)"
}

$UserAgent = $UserAgent.Trim()
if ($UserAgent.Length -lt 12) {
    throw "SEC User-Agent is too short. Use a descriptive application name plus a real contact email."
}
if ($UserAgent -notmatch '[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}') {
    throw "SEC User-Agent must include a real contact email, for example: EquityResearch your-real-email@example.com"
}

$env:SEC_USER_AGENT = $UserAgent
python (Join-Path $Root "automation\sec_preflight.py") --no-network --strict
if ($LASTEXITCODE -ne 0) {
    throw "SEC User-Agent validation failed."
}

$UserAgent | gh secret set SEC_USER_AGENT --repo $Repository
if ($LASTEXITCODE -ne 0) {
    throw "Failed to update SEC_USER_AGENT in $Repository."
}

Write-Host "SEC_USER_AGENT stored as a GitHub Actions secret."
Write-Host "The value was not written into the repository or printed back to the terminal."
Write-Host "Daily ML learning will now prefer SEC Company Facts for eligible US fundamentals."

if ($RunLearning) {
    gh workflow run "Daily ML learning maintenance" --repo $Repository
    if ($LASTEXITCODE -ne 0) {
        throw "Secret was saved, but the ML learning workflow could not be dispatched."
    }
    Write-Host "Daily ML learning maintenance dispatched."
}
