# Investment Research & Portfolio Analytics Showcase

This folder is the recruiter-safe public demonstration layer for the larger equity-research and portfolio-management project.

## What recruiters can see

- interactive Streamlit equity-research dashboard
- historical financial progression and segment analysis
- scenario valuation examples
- real sanitized aggregate portfolio analytics when `data/portfolio_snapshot.json` is present
- anonymized portfolio weights and risk contributions
- portfolio-vs-benchmark growth path
- Jensen/CAPM alpha and multi-factor alpha comparisons
- factor/style exposures
- stress testing
- methodology and architecture

## Portfolio data policy

The recruiter-facing portfolio snapshot can use **real aggregate outputs from the production portfolio**, while sensitive position information remains excluded.

The public snapshot deliberately excludes:

- ticker symbols and company names
- share counts
- average cost / cost basis
- portfolio market value
- unrealized P&L
- transaction history
- private Excel models
- private GitHub Actions artifacts
- API credentials or secrets

The public holdings are displayed as `Holding A`, `Holding B`, etc. Their weights and aggregate risk analytics can be real.

Equity-research example figures remain illustrative unless a specific public case study is intentionally added.

## Recruiter-safe validation

Before public publication, run:

```powershell
python .\automation\validate_showcase.py
```

The validator requires:

- a real sanitized portfolio snapshot
- core portfolio performance/risk metrics
- multiple anonymous holdings
- portfolio growth history
- portfolio weights that approximately sum to 100%
- no forbidden position-level identifiers or sensitive economics

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
powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

The export script now requires validated real portfolio analytics by default. `-AllowDemo` should only be used for a non-recruiter demonstration.

## Recommended production architecture

- **Private main repository:** full research engine, portfolio inputs, models and private portal
- **Public showcase repository:** only `app.py`, `requirements.txt`, `README.md`, `.gitignore`, and the sanitized `data/portfolio_snapshot.json`
- **Public Streamlit Community Cloud app:** deployed from the public showcase repository

The intended public repository name is:

`Antzaz/Antzaz-investment-research-showcase`

## Initial public GitHub publish

After running `export_showcase.ps1`:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"
gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push
```

## Automatic portfolio refresh

The existing `Daily private portfolio refresh` GitHub Actions workflow can automatically update only the sanitized JSON in the public showcase repository.

Configure the main repository with:

- Actions secret: `SHOWCASE_REPO_TOKEN`
- Actions variable: `SHOWCASE_REPOSITORY=Antzaz/Antzaz-investment-research-showcase`

The token should have the minimum permissions necessary to write repository contents to the public showcase repository.

## Streamlit Community Cloud

Deploy `app.py` from the separate public showcase repository and make the app public. No Streamlit secrets are required for this showcase.

Once deployed, use the Streamlit URL—not the private research repository—as the project link in your resume and applications.
