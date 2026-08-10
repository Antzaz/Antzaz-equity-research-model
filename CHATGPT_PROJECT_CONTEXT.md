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

## Current repository areas

### Single-company model

Repository root, including `update_model.py` and generated research workbooks.

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
