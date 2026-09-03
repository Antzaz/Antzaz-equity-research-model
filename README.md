# Equity Research & Portfolio Analytics

Personal Python/Excel/Streamlit research framework for repeatable company and portfolio analysis.

## Equity-research workflow

Run the **guarded production company model** with:

```powershell
python .\commodity_safe_runner.py TSM
```

Or use the agent orchestrator below, which builds the same guarded workbook first.

`update_model.py` is the low-level deterministic generator used internally by the guarded runtime. For new cross-sector investment work, use `commodity_safe_runner.py` or `research.py` so business-model routing, statement-profile selection, valuation/score gates, commodity normalization, verified segment adapters and final Data Quality controls are installed. See `CROSS_SECTOR_MODELING.md`.

The generated workbook includes:

- company and market data;
- normalized historical financials and full financial statements;
- segment / business-mix analysis;
- DCF and three-case scenarios;
- advanced valuation, reverse-DCF and Monte Carlo analytics where economically appropriate;
- expanded peer comps with target + up to nine peers;
- direct/exact/strategic peer labels and per-company data-coverage percentages;
- calculated fallback peer metrics when live summary fields are missing and statement data can support the calculation;
- comparable industry market share plus selected-peer-set market-value share;
- source-scoped business market-position snapshots where a reputable specialist source is mapped;
- moat / competitive-advantage research;
- analyst consensus and market-implied expectations;
- ownership and holders;
- AI impact and recent-news analysis;
- leadership, workforce and governance research;
- same-sector alternative-company screening;
- a 10-lens institutional investment-style comparison;
- a consolidated research workbench and data-quality controls.

### Cross-sector modeling contract

The production runner classifies the issuer by business model rather than forcing every ticker through an industrial-company template. Banks, insurers, REITs, commodity producers, utilities, software/cloud, semiconductors, pharma/biotech, industrials, consumer businesses, payments networks and digital platforms receive explicit modeling policy. Inappropriate metrics are marked `REVIEW`, `N/M` or excluded from scoring rather than being silently treated as valid. Full details and the test matrix are in `CROSS_SECTOR_MODELING.md`.

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

## Source hierarchy and missing-data policy

The model prefers primary and issuer-owned sources where possible:

1. issuer investor-relations pages, annual reports, results and sustainability/governance disclosures;
2. SEC / regulatory filings and XBRL when available;
3. public market-data feeds as transparent fallbacks;
4. explicitly mapped specialist industry sources for comparable market-position data.

`source_registry.py` is the shared catalog for issuer-owned and specialist research pages used by the newer research layers. Missing or non-comparable data is left blank or marked `REVIEW` rather than silently estimated.

For peer analysis, missing live summary fields can now be recovered from public statement/market data where the calculation is mechanically supportable. Examples include revenue growth, operating margin, ROE, enterprise value, EV/Revenue, EV/EBITDA and forward P/E. The Peer Comps sheet labels calculated fallback usage in `Metric Source / Fallback Notes` and reports a `Data Coverage %` for every company.

Foreign issuers are normalized to the traded security's quote currency before valuation and consensus comparisons. SEC outages or throttling are designed to fall through to issuer/Yahoo recovery instead of stopping the entire model.

## Workbook consolidation

The generator removes redundant presentation sheets. `Dashboard`, `Peer Comps`, `Advanced Analytics` and `AI Impact Analysis` are the authoritative presentation layers rather than maintaining duplicate dashboards/comparison/valuation/AI tabs.

The removed legacy presentation tabs are:

- `Visual Dashboard`;
- `Comparative Analysis`;
- `Valuation Cross-Checks`;
- `AI Analysis`;
- `AI Valuation`.

The research/admin layer is also consolidated:

- `Research Notes` + `Research Checklist` + `Research Journal` become `Research Workbench` when no formula dependency prevents safe consolidation;
- `Data Dictionary` content is preserved inside `Data Quality`, then the standalone dictionary tab is removed when safe.

People/leadership blocks in `Dashboard` and `Investment Summary` are written idempotently so the same section is not repeatedly appended within an existing workbook.

## Institutional comparison

`Institutional Comparison` applies ten transparent public-style investment lenses to the same reusable company-quality dimensions:

1. Berkshire Hathaway / Buffett-style;
2. Fundsmith;
3. Dodge & Cox;
4. Davis Advisors;
5. Pershing Square;
6. Baillie Gifford — Long Term Growth;
7. Akre Capital;
8. Polen Capital;
9. Capital Group;
10. T. Rowe Price.

Each row shows the public investment lens, a transparent fit score, data coverage, the company characteristics that fit the lens, the characteristics that deserve more diligence, and a public source describing the firm's investment approach. These are research comparisons only: they are not proprietary institutional models, holdings claims, target prices, endorsements or predictions that a firm would buy the company.

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

GitHub Actions runs syntax/policy tests plus representative live workbook builds. In addition to issuer-specific tests, the cross-sector contract validates software, semiconductors, banks, insurance, REITs, energy, utilities, pharma, industrials and consumer companies for core-sheet presence, statement-profile suitability, valuation gates, final Data Quality state and broken formula references.

Generated workbooks, agent run reports, private portfolio files, local Streamlit secrets and runtime caches are excluded from Git tracking. Durable public KPI history under `research_data/` is retained for longitudinal thesis monitoring.
