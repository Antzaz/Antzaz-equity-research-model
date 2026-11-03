# One-time Windows preparation: sync the repository, install portfolio dependencies,
# build private analytics and cache price history for offline use.
# No market positions, costs or credentials are printed or uploaded by this script.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw 'Git is not installed or is not on PATH.'
}

$portfolio = Join-Path $root 'institutional_research\portfolio.csv'
$cache = Join-Path $root 'institutional_research\.offline_cache'
New-Item -ItemType Directory -Path $cache -Force | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (Test-Path $portfolio) {
    $backup = Join-Path $cache "portfolio_before_setup_$stamp.csv"
    Copy-Item -LiteralPath $portfolio -Destination $backup -Force
    Write-Host "Saved a private local portfolio backup under .offline_cache (not committed)."
}

Write-Host 'Pulling latest research and portfolio code from GitHub...'
& git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) {
    & git ls-files --error-unmatch -- institutional_research/portfolio.csv *> $null
    $tracked = $LASTEXITCODE -eq 0
    & git cat-file -e origin/main:institutional_research/portfolio.csv *> $null
    $remoteHasFile = $LASTEXITCODE -eq 0

    if ((-not $tracked) -and $remoteHasFile -and (Test-Path $portfolio)) {
        $migrated = Join-Path $cache "portfolio_untracked_before_pull_$stamp.csv"
        Move-Item -LiteralPath $portfolio -Destination $migrated
        try {
            & git pull --ff-only origin main
            if ($LASTEXITCODE -ne 0) {
                throw 'Git pull failed after backing up the previous untracked portfolio.'
            }
            Write-Host 'GitHub portfolio.csv is active. Previous local copy remains backed up.'
        } catch {
            if (-not (Test-Path $portfolio)) {
                Move-Item -LiteralPath $migrated -Destination $portfolio
            }
            throw
        }
    } else {
        throw 'Git pull could not fast-forward. Existing work was preserved. Run git status and resolve any local changes.'
    }
}

$py = Get-Command py -ErrorAction SilentlyContinue
if ($py) {
    $pythonCmd = 'py'
    $pythonArgs = @('-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = 'python'
    $pythonArgs = @()
} else {
    throw 'Python 3 is not installed or not on PATH.'
}

Write-Host 'Installing portfolio runtime dependencies while online...'
& $pythonCmd @pythonArgs -m pip install -r (Join-Path $root 'institutional_research\requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Portfolio dependency installation failed.' }

Push-Location (Join-Path $root 'institutional_research')
try {
    Write-Host 'Building current portfolio outputs and persistent offline market cache...'
    & $pythonCmd @pythonArgs run_research.py
    if ($LASTEXITCODE -ne 0) {
        throw 'Portfolio research sync failed. Existing offline outputs have been preserved.'
    }

    Write-Host 'Checking that saved analytics and price histories cover every holding...'
    & $pythonCmd @pythonArgs offline_portfolio_check.py
    if ($LASTEXITCODE -ne 0) {
        throw 'Offline readiness check did not pass. See the preceding missing-data diagnostics.'
    }
} finally {
    Pop-Location
}

Write-Host ''
Write-Host 'Offline portfolio is ready on this computer.'
Write-Host 'Open Antzaz Research Hub or institutional_research\start_dashboard.bat even when disconnected.'
