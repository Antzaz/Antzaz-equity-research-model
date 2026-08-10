# Anton Hiltunen — Investment Research Portfolio

This folder contains a recruiter-facing Streamlit presentation layer for the broader equity-research and portfolio project.

## Recruiter experience

The app is designed so a recruiter can understand the work without reading code. It shows:

- model-portfolio overview and performance;
- published portfolio companies and weights;
- rating, fair value, modeled upside and conviction;
- concise company investment theses;
- catalysts, risks and thesis-falsification conditions;
- valuation scenarios and monitored KPIs;
- a high-level research process and methodology.

## Editing research without touching the app

All recruiter-facing thesis text and manually entered values live in:

`data/recruiter_portfolio.json`

See `EDIT_PORTFOLIO.md` and `data/recruiter_portfolio.example.json`.

Only companies with `"published": true` appear in the recruiter view.

## Privacy architecture

Recommended setup:

1. Keep the full research repository private.
2. Deploy `showcase/app.py` to Streamlit Community Cloud from that private repository.
3. Make the Streamlit app itself public and share the `*.streamlit.app` URL on the CV.
4. Do **not** put the GitHub repository URL on the recruiter-facing CV unless you intentionally want reviewers to inspect code.

The app does not need to show transaction history, personal portfolio value, cost basis, credentials, private Excel models or research-engine internals.

## Optional sanitized analytics

If `data/portfolio_snapshot.json` exists, the app can use sanitized aggregate portfolio metrics as a fallback. Values entered in `recruiter_portfolio.json` take precedence.

## Run locally

```powershell
python -m pip install -r .\showcase\requirements.txt
python -m streamlit run .\showcase\app.py
```

This showcase is for project demonstration and recruitment, not investment advice.
