# Online Portfolio Update Flow

The recruiter-facing portfolio uses two separate update layers:

```text
LOCAL PORTFOLIO COMPOSITION
institutional_research/portfolio.csv
        ↓
automation/sync_portfolio_secret.ps1
        ↓
GitHub Actions secret: PORTFOLIO_CSV_B64
        ↓
GitHub portfolio analytics refresh
        ↓
sanitized portfolio_snapshot.json
        ↓
public showcase GitHub repository
        ↓
Streamlit fetches newest snapshot when opened
```

## 1. Local portfolio.csv is the composition source

Edit this file whenever you buy, sell, add, remove or reweight holdings:

```text
C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research\portfolio.csv
```

Then run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\automation\sync_portfolio_secret.ps1
```

The sync script now does two things by default:

1. validates and copies the local portfolio composition into the private GitHub Actions source (`PORTFOLIO_CSV_B64`);
2. immediately dispatches `Daily private portfolio refresh` so the online analytics are rebuilt.

Use `-NoRefresh` only when you intentionally want to update the stored cloud portfolio without rebuilding immediately.

The old `-RunRefresh` argument is still accepted for backwards compatibility but is no longer necessary.

## 2. Automatic GitHub refresh

The full private workflow is currently scheduled at:

```text
08:17 Europe/Helsinki every day
```

It can also be triggered manually, including automatically by the local sync script above.

The workflow:

1. restores the private portfolio from `PORTFOLIO_CSV_B64`;
2. downloads fresh market/company data;
3. rebuilds portfolio analytics;
4. builds the sanitized public snapshot;
5. validates that no sensitive position data are present;
6. pushes the sanitized snapshot to the public showcase repository when publishing credentials are configured.

## 3. What happens when a recruiter opens Streamlit

The public `showcase/app.py` now attempts to fetch:

```text
Antzaz/Antzaz-investment-research-showcase/main/data/portfolio_snapshot.json
```

directly from GitHub every time a new Streamlit browser session starts.

A cache-busting request is used so the browser session asks for the current GitHub file rather than relying on the deployed app's bundled copy.

If GitHub cannot be reached temporarily, the app falls back to the last validated snapshot bundled with the deployment.

If neither source contains a valid real snapshot, the portfolio page does **not** show fake/demo portfolio figures. It shows an unavailable message instead.

## 4. What "update when opened" means

Opening Streamlit now refreshes the **displayed snapshot from GitHub**.

It does not trigger the private research workflow itself. This is intentional: allowing anonymous recruiter page loads to trigger private GitHub Actions would require a privileged credential and could be abused.

Therefore freshness is:

- after changing local `portfolio.csv`: run the sync command and GitHub rebuild starts immediately;
- without portfolio-composition changes: the scheduled GitHub workflow refreshes the analytics at 08:17 Europe/Helsinki;
- whenever Streamlit is opened/reloaded: it fetches the latest snapshot that GitHub has already produced.

## 5. One command after changing holdings

From the main repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\automation\sync_portfolio_secret.ps1
```

That is the normal portfolio-update command going forward.

## 6. Privacy boundary

The public snapshot is designed to contain aggregate/recruiter-safe data only, including:

- return and volatility
- Sharpe / Sortino
- tracking error / information ratio
- drawdown and beta
- anonymized weights and risk contributions
- alpha/factor diagnostics
- stress scenarios
- portfolio-vs-benchmark history

It excludes:

- ticker symbols
- company names
- share counts
- average cost / cost basis
- portfolio value
- unrealized P&L
- transaction history
- private workbooks
- credentials
