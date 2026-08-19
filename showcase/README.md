# Investment Research & Portfolio Analytics Showcase

This folder is the recruiter-safe public demonstration layer for the larger equity-research and portfolio-management project.

## What recruiters can see

- interactive Streamlit portfolio analytics dashboard
- real company names and current portfolio weights
- portfolio-vs-benchmark growth path
- annualized return, volatility, Sharpe/Sortino, tracking error, information ratio, drawdown and beta
- risk contributions, factor/style exposures and stress testing
- Jensen/CAPM and multi-factor alpha comparisons
- an Investment Thesis tab with portfolio philosophy and company-level reasoning
- company research scorecards, conviction, expected return, risks, sell/falsification conditions and monitoring KPIs when published from the private thesis workbook
- illustrative equity-research dashboard examples
- methodology and architecture

## Portfolio data policy

The recruiter-facing portfolio snapshot uses **real outputs from the production portfolio** while sensitive position economics remain excluded.

Public by design:

- company names
- portfolio weights
- sectors when available
- aggregate portfolio performance/risk analytics
- explicitly recruiter-safe investment philosophy and thesis fields

Deliberately excluded:

- ticker symbols from the public data payload
- share counts
- average cost / cost basis
- exact portfolio market value
- unrealized P&L
- transaction history
- Private Notes from the thesis workbook
- private Excel research models
- private GitHub Actions artifacts
- API credentials or secrets

Equity-research example figures remain illustrative unless a specific public case study is intentionally added.

## Investment thesis workbook

Keep the private workbook at:

`institutional_research/portfolio_thesis.xlsx`

Use the provided workbook template with two recruiter-input sheets:

- `Portfolio Philosophy`
- `Company Theses`

Only rows/sections marked `Publish = Yes` are eligible for the public dashboard. `Private Notes (never published)` is explicitly excluded by the sync pipeline.

The normal portfolio sync command synchronizes both `portfolio.csv` and the recruiter-safe thesis fields when the workbook exists:

```powershell
powershell -ExecutionPolicy Bypass -File .\automation\sync_portfolio_secret.ps1
```

## Recruiter-safe validation

Before public publication, run:

```powershell
python .\automation\validate_showcase.py
```

The validator requires:

- a real sanitized portfolio snapshot
- core portfolio performance/risk metrics
- multiple named holdings
- portfolio growth history
- portfolio weights that approximately sum to 100%
- no forbidden sensitive position economics, tickers, transactions or private notes

The scheduled portfolio workflow also runs this validation automatically before publishing a refreshed snapshot.

## Run locally

From the main project root:

```powershell
python -m pip install -r .\showcase\requirements.txt
python -m streamlit run .\showcase\app.py
```

## Build the recruiter version

First refresh the private analytics, then create the sanitized public export:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"
git pull
python -m pip install -r .\institutional_research\requirements.txt
python .\institutional_research\run_research.py
python .\automation\build_showcase_snapshot.py
python .\automation\validate_showcase.py
powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

## Recommended production architecture

- **Private main repository:** full research engine, portfolio inputs, thesis workbook, models and private portal
- **Public showcase repository:** only `app.py`, `requirements.txt`, `README.md`, `.gitignore`, and sanitized `data/portfolio_snapshot.json`
- **Public Streamlit app:** deployed from the public showcase repository

The intended public repository name is:

`Antzaz/Antzaz-investment-research-showcase`

## Initial public GitHub publish

After running `export_showcase.ps1`:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"
gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push
```

## Automatic portfolio refresh

The `Daily private portfolio refresh` GitHub Actions workflow rebuilds the private analytics, restores recruiter-safe thesis inputs, validates the public snapshot and can update only the sanitized JSON in the public showcase repository.

Configure the main repository with:

- Actions secret: `SHOWCASE_REPO_TOKEN`
- Actions variable: `SHOWCASE_REPOSITORY=Antzaz/Antzaz-investment-research-showcase`

The token should have the minimum permissions necessary to write repository contents to the public showcase repository.

## Streamlit hosting

Deploy `app.py` from the separate public showcase repository and make the app public. The Streamlit app fetches the newest published GitHub snapshot whenever a new session opens.

Once deployed, use the Streamlit URL—not the private research repository—as the project link in your resume and applications.
