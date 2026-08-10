# Project launch commands

Assumes the repository is located at:

`C:\Users\Antza\Documents\Antzaz-equity-research-model`

## Equity research model

Replace `UBS` with any ticker:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python .\update_model.py UBS
```

## Agent-assisted ticker research — recommended full workflow

Builds the normal deterministic workbook, then runs the Filings, KPI/Earnings, Thesis Monitor and Research QA agents:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python .\research.py GOOGL
```

Enable evidence-bound OpenAI reasoning after setting your API key locally:

```powershell
$env:OPENAI_API_KEY="YOUR_KEY"; python .\research.py GOOGL --ai
```

Reuse the newest already-generated workbook without rebuilding it:

```powershell
python .\research.py GOOGL --skip-model
```

Guide: `AI_AGENTS.md`

## Portfolio research — refresh data and launch local dashboard

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research"; python -m pip install -r .\requirements.txt; python .\run_research.py; python -m streamlit run .\dashboard.py
```

## Institutional investor comparison

Creates/uses private local inputs for your company assumptions, public institutional disclosures and thesis-falsification rules, then exports Berkshire/Buffett-style screening and direct institutional comparisons.

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research"; python .\run_institutional_comparison.py
```

Guide: `institutional_research\INSTITUTIONAL_COMPARISON.md`

Private inputs created from templates:

- `company_thesis.csv`
- `institutional_views.csv`
- `thesis_risks.csv`

Outputs:

- `outputs\institutional_comparison\buffett_style_scorecard.csv`
- `outputs\institutional_comparison\institutional_comparison.csv`
- `outputs\institutional_comparison\institutional_consensus.csv`
- `outputs\institutional_comparison\thesis_risks.csv`

## Private unified live portal

Requires the `[live_data]` Streamlit secrets and a successful `Daily private portfolio refresh` GitHub Actions run. The dependency install is important because the encrypted live bundle uses `pyzipper`.

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python -m pip install -r .\institutional_research\requirements.txt; python -m streamlit run .\institutional_research\live_dashboard.py
```

The private portal contains:

- Portfolio Dashboard
- Alpha Analysis
- Company Research

## Resume-safe showcase — refresh from real portfolio outputs and launch

This first refreshes the private portfolio analytics, then builds a sanitized public snapshot with real aggregate metrics and anonymous holdings.

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python -m pip install -r .\institutional_research\requirements.txt; python .\institutional_research\run_research.py; python .\automation\build_showcase_snapshot.py; python -m pip install -r .\showcase\requirements.txt; python -m streamlit run .\showcase\app.py
```

The public snapshot can contain real:

- portfolio return, volatility, Sharpe/Sortino;
- tracking error, information ratio, drawdown, beta, VaR/ES;
- anonymous portfolio weights and risk contributions;
- CAPM/Jensen, FF3, Carhart and FF5 alpha statistics;
- factor/style exposures;
- stress results;
- portfolio-vs-benchmark growth path.

It excludes tickers, company names, shares, average cost, portfolio value, unrealized P&L, transaction history, private Excel workbooks and credentials.

## Export only the public showcase into a clean separate folder

The exporter automatically refreshes the sanitized real snapshot if `institutional_research/outputs/latest` exists.

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

The default output folder is:

`C:\Users\Antza\Documents\Antzaz-investment-research-showcase`

If GitHub CLI (`gh`) is installed and authenticated, publish the exported folder as a new public repository with:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"; gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push
```

## Recommended hosting structure

- Main repository: PRIVATE
- Streamlit private app: `institutional_research/live_dashboard.py`
- Friend access: invite as Streamlit viewer, not GitHub collaborator
- Resume/public app: deploy `app.py` from the separate public showcase repository
