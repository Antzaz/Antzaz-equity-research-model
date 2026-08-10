# AI Research Agents

This project uses agents as an evidence and reasoning layer around the existing deterministic Python/Excel model. Agents must not silently change valuation assumptions or execute trades.

## Phase 1 — implemented

Run:

```powershell
python .\research.py GOOGL
```

This builds the normal research workbook and then runs five specialist agents:

1. **Source Health Agent** — checks the centralized issuer/specialist source registry and performs a small bounded reachability probe. It is deterministic and uses no AI tokens.
2. **Filings & Financials Agent** — inventories primary filing references already collected by the workbook and flags coverage gaps.
3. **KPI / Earnings Agent** — extracts the `AI Impact Analysis` evidence table, writes an auditable run snapshot, and maintains `research_data/<TICKER>/kpi_history.json`.
4. **Thesis Monitor Agent** — converts evidence signals into a transparent triage score and, when AI reasoning is enabled, assesses growth, margins, capital intensity, moat, disruption and falsification questions.
5. **Research QA / Skeptic Agent** — checks source URLs, disclosure types, dates, duplicate/stale KPI evidence and separation between evidence and inference.

Generated run artifacts are stored under `research_runs/<TICKER>/<timestamp>/` and are ignored by Git. Public KPI history under `research_data/` is intentionally durable.

## Source discipline

`source_registry.py` is the shared catalog for issuer-owned research pages and specialist market-position sources. It exists to avoid maintaining the same URL lists in multiple research layers.

- Prefer issuer filings, earnings, annual reports, governance and workforce disclosures for company facts.
- Regulator/XBRL data remains the next financial-data layer where applicable.
- Transparent market-data fallbacks may fill gaps but must remain labeled as fallbacks.
- Specialist sources such as TrendForce, Synergy Research, StatCounter or IDC are used only for the market definitions they actually measure.
- Industry market-share percentages may enter `Peer Comps` only when one comparable source/market definition covers the peer set.
- `Peer-Set Market Cap %` is a separate calculated concentration measure and must never be labeled as industry market share.
- Employee-review platforms may be used for manual corroboration, not silently treated as audited company-wide employee happiness.

## Optional AI reasoning

The deterministic agents run without an API key. To enable evidence-bound LLM review:

```powershell
$env:OPENAI_API_KEY="YOUR_KEY"
python .\research.py GOOGL --ai
```

Optional model override:

```powershell
$env:OPENAI_RESEARCH_MODEL="gpt-5.5"
python .\research.py GOOGL --ai
```

The OpenAI layer receives only the locally extracted evidence payload. It has no trading permissions and no browsing tools. AI text is stored as narrative and is never treated as a primary source or a DCF input.

## Useful modes

Reuse the newest workbook instead of rebuilding it:

```powershell
python .\research.py GOOGL --skip-model
```

Fail automation only when Research QA reaches `FAIL`:

```powershell
python .\research.py GOOGL --strict
```

## Governance rules

- Python/Excel calculates; agents research, classify and challenge.
- Company filings and issuer material remain preferred primary evidence.
- Every KPI should retain an as-of date, source URL and disclosure/data type.
- AI must say when evidence is insufficient.
- No agent silently changes WACC, growth, margins, capex, terminal assumptions or position sizes.
- No agent executes trades.
- Thesis-monitor scores are triage aids, not investment ratings.

## Roadmap

The same orchestrator is designed to add these specialist agents next:

- Institutional Investor Comparison Agent
- Competitive Intelligence Agent
- Valuation Challenge / Reverse-DCF Agent
- Portfolio Manager Agent
- Research Journal / Decision Review Agent

Those agents should reuse the repository's existing institutional-comparison, peer, valuation and portfolio modules rather than duplicate calculations.
