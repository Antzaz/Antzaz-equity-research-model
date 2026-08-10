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
- comparable industry market share plus selected-peer-set market-value share;
- source-scoped business market-position snapshots where a reputable specialist source is mapped;
- moat / competitive-advantage research;
- analyst consensus and market-implied expectations;
- ownership and holders;
- AI impact and recent-news analysis;
- leadership, workforce and governance research;
- same-sector alternative-company screening;
- research notes, checklist, journal and data-quality controls.

## Agent-assisted research

Use the agent orchestrator when you want the normal company model plus evidence monitoring, source checks and research QA:

```powershell
python .\research.py GOOGL
```

Phase 1 adds five specialist agents around the deterministic workbook:

- Source Health Agent;
- Filings & Financials Agent;
- KPI / Earnings Agent with durable KPI history;
- Thesis Monitor Agent;
- Research QA / Skeptic Agent.

The normal command uses no OpenAI API tokens. Optional evidence-bound LLM reasoning can be enabled with `python .\research.py GOOGL --ai` after setting `OPENAI_API_KEY`. AI narrative never overwrites model inputs or executes trades. See `AI_AGENTS.md` for governance, commands and the roadmap.

## Source hierarchy

The model prefers primary and issuer-owned sources where possible:

1. issuer investor-relations pages, annual reports, results and sustainability/governance disclosures;
2. SEC / regulatory filings and XBRL when available;
3. public market-data feeds as transparent fallbacks;
4. explicitly mapped specialist industry sources for comparable market-position data.

`source_registry.py` is the shared catalog for issuer-owned and specialist research pages used by the newer research layers. Missing or non-comparable data is left blank or marked `REVIEW` rather than silently estimated.

Foreign issuers are normalized to the traded security's quote currency before valuation and consensus comparisons. SEC outages or throttling are designed to fall through to issuer/Yahoo recovery instead of stopping the entire model.

## Workbook consolidation

The generator removes redundant presentation sheets. `Dashboard`, `Peer Comps`, `Advanced Analytics` and `AI Impact Analysis` are the authoritative presentation layers rather than maintaining duplicate dashboards/comparison/valuation/AI tabs.

The removed legacy presentation tabs are:

- `Visual Dashboard`;
- `Comparative Analysis`;
- `Valuation Cross-Checks`;
- `AI Analysis`;
- `AI Valuation`.

People/leadership blocks in `Dashboard` and `Investment Summary` are written idempotently so the same section is not repeatedly appended within an existing workbook.

## People and leadership research

`Leadership & Culture` separates evidence from interpretation:

- employee satisfaction/engagement signals retain their exact reported scope;
- missing company-wide worker-happiness evidence is shown as `REVIEW`, not fabricated;
- leadership scoring is explicitly a transparent research proxy, not a factual management rating;
- executive and governance sources are linked where available;
- source-scoped market-position snapshots preserve their own market definition and period;
- the alternative-company screen compares validated same-sector peers on growth, margins, ROE and valuation, and also uses industry market share only when that field is comparable across multiple peers;
- a better-company candidate is surfaced only with adequate metric coverage and a material score gap.

## Quality checks

GitHub Actions runs:

- a Python syntax check on the core research modules and agent framework; and
- end-to-end TSM and Alphabet smoke tests that build real workbooks and validate consolidated tabs, source coverage, peer percentages, leadership/workforce sections, market-position data, agent output and broken formula references.

Generated workbooks, agent run reports, private portfolio files, local Streamlit secrets and runtime caches are excluded from Git tracking. Durable public KPI history under `research_data/` is retained for longitudinal thesis monitoring.
