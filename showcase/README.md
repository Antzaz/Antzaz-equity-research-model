# Investment Research & Portfolio Analytics Showcase

This folder is the resume-safe public demonstration layer for the larger private research project.

## What it shows

- Equity Research dashboard
- Historical financial progression
- Segment analysis
- Scenario valuation
- Portfolio risk and concentration
- Jensen/CAPM alpha
- Fama-French / Carhart style alpha comparisons
- Factor exposures
- Stress testing
- Methodology and architecture

## What it deliberately does not contain

- real portfolio holdings
- shares or market value
- average cost / cost basis
- transaction history
- private Excel models
- GitHub Actions artifacts
- API credentials or secrets
- private portfolio bundle access

All figures are synthetic/sanitized demonstration data.

## Run locally

From the main project root:

```powershell
python -m pip install -r .\showcase\requirements.txt
python -m streamlit run .\showcase\app.py
```

## Separate public repository

The recommended production architecture is:

- private main repository: full research engine and private portal
- private Streamlit portal: you + invited viewers
- separate public showcase repository: only the contents of this folder

Use `automation/export_showcase.ps1` to create a clean local folder containing only the public showcase files.

## Streamlit Community Cloud

Deploy `app.py` from the separate showcase repository and set the app to public. No Streamlit secrets are required for the showcase.
