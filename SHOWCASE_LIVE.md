# Public showcase with real sanitized portfolio analytics

The public showcase can display genuine aggregate portfolio analytics without publishing the portfolio universe or sensitive position economics.

## What is public

The sanitized snapshot may include:

- annualized return and volatility;
- Sharpe and Sortino ratios;
- tracking error and information ratio;
- max drawdown, beta, VaR and Expected Shortfall;
- active return;
- anonymous portfolio weights and risk contributions (`Holding A`, `Holding B`, ...);
- CAPM/Jensen, Fama-French 3, Carhart 4 and Fama-French 5 alpha statistics;
- public factor/style exposures;
- portfolio stress results;
- portfolio-vs-benchmark growth path.

## What is never exported

- ticker symbols;
- company names;
- shares;
- average cost / cost basis;
- market value or total portfolio value;
- unrealized P&L;
- transactions;
- private Excel models;
- credentials, tokens or secrets.

## Build locally

From the main repository root:

```powershell
python .\institutional_research\run_research.py
python .\automation\build_showcase_snapshot.py
python -m streamlit run .\showcase\app.py
```

Or use the exporter:

```powershell
powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

The exporter copies the sanitized snapshot into the clean public showcase folder.

## Automatic daily public refresh

The private repository workflow `.github/workflows/daily-portfolio-refresh.yml` now builds the sanitized snapshot after the portfolio analytics refresh.

To let it update the separate public showcase repository automatically:

1. Create the public repository `Antzaz/Antzaz-investment-research-showcase` and publish the exported showcase once.
2. Create a fine-grained GitHub token that has **Contents: Read and write** access only to that showcase repository.
3. In the main repository, open **Settings → Secrets and variables → Actions**.
4. Add repository secret `SHOWCASE_REPO_TOKEN` with that token.
5. Optional: add Actions variable `SHOWCASE_REPOSITORY` if the public repository is not named `Antzaz/Antzaz-investment-research-showcase`.
6. Run **Daily private portfolio refresh** manually once.

If `SHOWCASE_REPO_TOKEN` is absent, the private workflow still completes normally and simply skips the public showcase publication step.

The public showcase repository receives only `data/portfolio_snapshot.json`; it never receives the private portfolio file or encrypted live-data bundle.
