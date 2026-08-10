# Equity Research & Portfolio Analytics

Personal Python/Excel/Streamlit research framework for repeatable company and portfolio analysis.

## Equity-research workflow

Run a company model with:

```powershell
python .\update_model.py TSM
```

The generated workbook includes:

- company and market data;
- normalized historical financials and full financial statements;
- segment / business-mix analysis;
- DCF and three-case scenarios;
- advanced valuation, reverse-DCF and Monte Carlo analytics;
- same-sector peer comps;
- market-share context when a comparable industry source is mapped;
- moat / competitive-advantage research;
- analyst consensus and market-implied expectations;
- ownership and holders;
- AI impact and recent-news analysis;
- leadership, workforce and governance research;
- same-sector alternative-company screening;
- research notes, checklist, journal and data-quality controls.

## Source hierarchy

The model prefers primary and issuer-owned sources where possible:

1. issuer investor-relations pages, annual reports, results and sustainability/governance disclosures;
2. SEC / regulatory filings and XBRL when available;
3. public market-data feeds as transparent fallbacks;
4. explicitly mapped specialist industry sources for comparable market-share data.

Missing or non-comparable data is left blank or marked `REVIEW` rather than silently estimated.

Foreign issuers are normalized to the traded security's quote currency before valuation and consensus comparisons. SEC outages or throttling are designed to fall through to issuer/Yahoo recovery instead of stopping the entire model.

## Workbook consolidation

The current generator removes redundant presentation sheets. `Dashboard`, `Peer Comps`, `Advanced Analytics` and `AI Impact Analysis` are the authoritative presentation layers rather than maintaining duplicate dashboards/comparison/valuation/AI tabs.

## People and leadership research

`Leadership & Culture` separates evidence from interpretation:

- employee satisfaction/engagement signals retain their exact reported scope;
- leadership scoring is explicitly a transparent research proxy, not a factual management rating;
- executive and governance sources are linked where available;
- the alternative-company screen compares validated same-sector peers on growth, margins, ROE and valuation, and only surfaces a candidate when the difference is material.

## Quality checks

GitHub Actions runs:

- a Python syntax check on core research modules; and
- an end-to-end TSM smoke test that builds a real workbook and validates currency scale, consensus alignment, market-share data, consolidated tabs and broken formula references.

Generated workbooks, private portfolio files, local Streamlit secrets and runtime caches are excluded from Git tracking.
