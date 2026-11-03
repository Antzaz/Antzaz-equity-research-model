# ChatGPT Project Context — Equity Research & Portfolio Management

Use this file as durable context when working with ChatGPT on this repository.

## Project objective

Build a professional-style personal equity-research and portfolio-management system that combines:

- single-company fundamental research
- valuation and expected-return analysis
- portfolio construction and risk analytics
- factor and alpha diagnostics
- institutional-investor comparison
- investment-thesis monitoring and falsification
- decision journaling

The goal is not to imitate an institutional data terminal. The goal is to apply institutional-quality research discipline using auditable public data and clearly labeled assumptions.

## Core workflow

```text
Screen
  -> Fundamental research
  -> Forecasts
  -> Valuation
  -> Expected return
  -> Institutional comparison
  -> Buffett-style quality/value screen
  -> Thesis falsification conditions
  -> Position sizing
  -> Portfolio construction
  -> Risk / factor / alpha analysis
  -> Monitoring
  -> Thesis review
  -> Decision journal
```

## Equity-research standard

For each company, try to maintain a standardized scorecard covering:

- Business Quality
- Competitive Moat
- Management / Capital Allocation
- Balance Sheet
- Growth
- Valuation
- Expected Return / IRR
- Margin of Safety
- Key Risks
- Falsification Conditions
- Institutional Agreement / Disagreement
- Buffett-Style Screen

Do not turn qualitative judgments into false precision. Scores should be accompanied by the evidence and assumptions that produced them.

## Institutional comparison standard

Compare the user's analysis with high-quality public material from professional investors when available, including examples such as:

- Berkshire Hathaway / Warren Buffett
- Davis Funds
- Fundsmith
- Pershing Square
- Dodge & Cox
- other relevant long-only, value, quality or sector-specialist managers

The manager set should depend on the company and strategy. Berkshire is not automatically the best benchmark for every security.

### Evidence rules

Prefer primary sources:

- regulatory filings / 13F
- annual reports
- shareholder letters
- fund letters
- official investor presentations
- official interviews or transcripts

Use secondary research only when needed and label it clearly.

Never infer that a manager agrees with the user's fair value, growth assumptions, WACC or expected IRR merely because the manager owns the stock.

For every institutional data point, preserve:

- institution
- metric
- value
- as-of date
- source type
- source
- disclosure status (`Disclosed`, `CalculatedFromDisclosure`, or `Inferred`)
- notes

## Berkshire / Buffett standard

Treat the Buffett-style score as an independent public-data framework inspired by recurring Berkshire/Buffett themes.

It is not:

- Berkshire Hathaway's private model
- Buffett's disclosed scoring formula
- Berkshire's target price
- Berkshire's intrinsic-value estimate
- proof that Buffett would buy the security

When Berkshire owns a security, distinguish reported ownership from an inferred investment thesis.

## Alpha-thesis discipline

For each company, identify where the user's forecast differs from market or institutional expectations.

Useful questions:

- Which assumption creates most of the valuation upside/downside?
- Is the user more optimistic on revenue, margins, capital intensity or terminal economics?
- What does the current market price imply?
- What would need to happen for the user's expected return to disappear?
- Which observable evidence would falsify the thesis?

## "What could prove me wrong?" rule

Every material position should have at least one pre-defined falsification condition.

A good condition is:

- observable
- specific
- connected to the original thesis
- defined before the evidence arrives

Avoid using share-price declines alone as thesis falsification unless price itself changes financing, solvency or another fundamental condition.

## Portfolio comparison

When institutional portfolio data is available, compare the user's portfolio with suitable benchmarks/managers on metrics such as:

- concentration
- sector exposure
- factor/style exposure
- weighted valuation
- weighted growth
- weighted quality
- volatility
- beta
- Sharpe / Sortino
- tracking error
- information ratio
- maximum drawdown
- active share when valid benchmark weights exist

Do not present portfolio metrics as comparable when methodologies, dates or data coverage differ materially.

## Agent-assisted research layer

`research.py` is the recommended orchestration entry point when agent monitoring is wanted around the normal company model.

```powershell
python .\research.py GOOGL
```

Phase 1 specialist agents:

- Filings & Financials Agent
- KPI / Earnings Agent
- Thesis Monitor Agent
- Research QA / Skeptic Agent

The agent layer must remain evidence-bound. Python/Excel remains the deterministic calculation engine; agents research, classify, compare, challenge and identify missing evidence. AI narrative must never silently overwrite WACC, revenue growth, margins, capex, terminal assumptions, position sizes or other model inputs.

Durable public KPI history is stored under `research_data/<TICKER>/kpi_history.json`. Generated per-run reports under `research_runs/` are runtime artifacts and are not committed.

Optional OpenAI reasoning is explicitly enabled with `--ai` and requires a locally supplied API key. The LLM receives locally extracted evidence only and has no trading permissions. See `AI_AGENTS.md` for detailed governance and the roadmap.

Planned specialist extensions should reuse existing project modules rather than duplicate calculations:

- Institutional Investor Comparison Agent
- Competitive Intelligence Agent
- Valuation Challenge / Reverse-DCF Agent
- Portfolio Manager Agent
- Research Journal / Decision Review Agent

## Finland / Nasdaq Helsinki source policy

For Yahoo-style Helsinki tickers ending in `.HE`, use the dedicated free-source route before generic market-data fallbacks:

1. Finnish OAM / Nasdaq Helsinki regulated disclosures and ESEF annual reports.
2. XBRL International `filings.xbrl.org` machine-readable ESEF xBRL-JSON sourced from OAM filings.
3. Issuer investor-relations material.
4. SEC annual XBRL only when applicable.
5. Yahoo annual statements as a transparent fallback.

Additional free Finland sources registered for research context include Nasdaq Nordic Company News, PRH Open Data, Statistics Finland PxWeb, Bank of Finland / FIN-FSA Open Data, Fingrid Open Data (free registration/API key required), and issuer-linked Vara Research consensus where the issuer provides the public link.

Helsinki builds should be Finland-native across the workbook, not merely at the historical-financial fallback layer:
- feed Finnish OAM/ESEF facts into the full IFRS statement renderer;
- recover dimensioned ESEF operating segments when reported;
- keep issuer-defined KPIs and latest interim/results tables in the dedicated `Issuer KPIs` sheet;
- prefer issuer/Euroclear registered-shareholder tables and legal share counts where available;
- use issuer stock-exchange releases for deal/event monitoring;
- use issuer-linked Vara Research for free consensus, estimate ranges and target-price context when available;
- keep adjusted/comparable issuer metrics separate from statutory IFRS values;
- keep currency/source labels Finland-aware through the final source-audit/polish passes;
- use the ECB AAA euro-area 10-year spot curve as the EUR risk-free proxy for Helsinki/EUR WACC rather than a US Treasury default; use Finnish 20% corporate tax as the 2026 fallback only when a reported effective tax rate is unavailable.
- prefer Nordic listed comparables before broad US fallbacks for Finnish issuers; for YIT-style Engineering & Construction businesses use the Nordic contractor/residential/infrastructure set as discovery candidates and still apply the normal sector, size, de-duplication and data-coverage gates.

The ESEF parser may use recognized consolidated standard IFRS concepts plus a narrow conservative allow-list of unadjusted issuer-extension concepts for core statutory metrics. It must not map adjusted/comparable/underlying KPIs into IFRS accounting rows. Unsupported concepts remain blank rather than inferred or fabricated. Numeric-coverage reporting should distinguish true extraction failures from analyst inputs, derived fields, optional provider coverage and legitimate IFRS presentation/non-disclosure.

## Public showcase refresh policy

The public repository `Antzaz/Antzaz-investment-research-showcase` must not rely on an old sanitized base snapshot for market freshness. The public repo owns a daily market-data overlay that runs after the U.S. cash close (currently 22:30 UTC every calendar day) and atomically refreshes:

- `data/portfolio_snapshot.json` as the canonical website payload;
- `data/public_analytics.json`;
- `data/public_fundamentals.json`;
- `data/short_term_performance.json`;
- `data/stock_quotes.json`.

The public workflow recalculates market-derived holdings characteristics, portfolio performance/risk, attribution, factors, public alpha proxies, stress scenarios, reverse DCF, timeseries, latest closes and standard return windows. It validates freshness and complete holding coverage before publishing.

Private-to-public synchronization from the main repository remains the source of truth for actual portfolio composition, published weights and recruiter-safe thesis/philosophy inputs. Those inputs only need to change when the user changes the private portfolio or thesis workbook. A missing/expired cross-repository token must not freeze daily public market data.

Public payload privacy remains strict: no ticker/symbol fields, shares, average cost, cost basis, exact market value, unrealized P&L, transactions, private notes or credentials. Company names, public weights, market prices, return windows and aggregate analytics may be public.

## Windows offline portfolio contract

The private portfolio input is currently tracked in the main repository at
`institutional_research/portfolio.csv` and has 15 holdings. The repo is still publicly
visible as of 2026-11-03, so quantities and average costs in that tracked file are
exposed: `.gitignore` cannot undo past tracking. Explicitly warn the user to set the
main repository private and address Git history; never copy these economics into the
public showcase.

Offline startup **must not** call GitHub, pip, Yahoo or the research generator.
The Windows launchers now open `desktop_research_hub.py` without an automatic Git pull,
and `institutional_research/start_dashboard.bat` sets `ANTZAZ_OFFLINE=1`. The Hub
checks `outputs/latest/summary.json` plus the portfolio's offline readiness before
starting the local Streamlit server, with startup logs under
`institutional_research/.offline_cache/`.

The shared market-data adapter `institutional_research/src/data.py` persists adjusted
closes and public issuer snapshots to the gitignored private local directory
`institutional_research/.offline_cache/`. The Stock Prices & Purchase Costs page
loads these cached closes when offline, showing the cached market date instead of a
fake current quote. `institutional_research/src/export.py` stages a complete new
snapshot and preserves the previous `outputs/latest` if refresh fails.

For a new Windows checkout, do one connected run of
`prepare_offline_portfolio.ps1` to install dependencies, sync GitHub safely while
backing up prior local portfolio CSVs, generate private research outputs and populate
the offline cache. Then run `python institutional_research/offline_portfolio_check.py`
and confirm 15 holdings and complete quote cache before disconnecting.
The portfolio can be launched offline afterwards, but GitHub alone cannot supply
the uncommitted local analytics/prices; a first successful connected sync is required.

## Current repository areas

### Single-company model

Repository root, including `update_model.py` and generated research workbooks.

### Agent-assisted single-company research

Repository root `research.py`, `agent_research/`, `research_data/` and `AI_AGENTS.md`.

### Institutional portfolio layer

`institutional_research/`

Core portfolio analytics are run with:

```powershell
python run_research.py
```

### Institutional comparison layer

Guide:

`institutional_research/INSTITUTIONAL_COMPARISON.md`

Run:

```powershell
python run_institutional_comparison.py
```

The comparison layer uses private local files created from committed templates:

- `company_thesis.csv`
- `institutional_views.csv`
- `thesis_risks.csv`

These private working files are ignored by Git.

## ChatGPT working instructions

When asked to analyze a company or portfolio in this project:

1. Use the latest available data when the answer depends on current holdings, prices, filings, fund positions or institutional opinions.
2. Prefer primary sources for professional-investor comparisons.
3. Separate disclosed facts from calculated values and inference.
4. Compare the user's assumptions directly with institutional/market assumptions where a true comparable exists.
5. Highlight the assumptions that drive the valuation gap.
6. Include a thesis-falsification section for material investment recommendations.
7. Keep outputs compatible with the repository's existing templates and research workflow when possible.
8. Do not invent private institutional valuation assumptions that are not publicly disclosed.
9. When agent-assisted monitoring is relevant, prefer `research.py TICKER` and preserve the rule that AI reasoning proposes or challenges assumptions but the deterministic model calculates them.
