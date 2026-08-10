# Institutional Comparison & Buffett-Style Research Layer

This layer compares your own company assumptions with public information disclosed by large investment firms, funds and other institutional investors.

It is designed to answer three different questions without mixing them together:

1. **What does my analysis imply?**
2. **What have professional investors actually disclosed?**
3. **How does the company look under a Buffett-style quality/value screen?**

The distinction matters. A fund holding a stock does **not** prove that the manager agrees with your target price, growth assumptions or expected return.

## Files

### Private working inputs

These are automatically created from templates and are ignored by Git:

- `company_thesis.csv` — your valuation and operating assumptions
- `institutional_views.csv` — values supported by public institutional disclosures
- `thesis_risks.csv` — thesis, key risk and falsification conditions

### Version-controlled templates

- `company_thesis_template.csv`
- `institutional_views_template.csv`
- `thesis_risks_template.csv`

### Analysis code

- `src/institutional_comparison.py`
- `run_institutional_comparison.py`

## Run

From `institutional_research`:

```powershell
python run_institutional_comparison.py
```

Outputs are written to:

```text
outputs/institutional_comparison/
```

The runner also imports `ExpectedReturn` automatically from `expected_returns.csv` when that file exists.

## Your thesis input

Use one row per ticker and metric.

```csv
Ticker,Metric,UserValue,Unit,AsOfDate,Notes
XYZ,FairValue,150,USD,2026-08-10,Base-case DCF
XYZ,RevenueCAGR,0.10,decimal,2026-08-10,Five-year forecast
XYZ,OperatingMargin,0.30,decimal,2026-08-10,Steady-state margin
XYZ,ExpectedReturn,0.13,decimal,2026-08-10,Annualized base case
```

Recommended metric names include:

- `FairValue`
- `TargetPrice`
- `ExpectedReturn`
- `RevenueCAGR`
- `EPSCAGR`
- `OperatingMargin`
- `FCFMargin`
- `ROIC`
- `WACC`
- `TerminalGrowth`
- `PortfolioWeight`

Use the same metric name and unit in the institutional file when you want a direct numerical comparison.

## Institutional view input

Only enter a numerical value when a source genuinely supports that value.

```csv
Ticker,Institution,Metric,InstitutionValue,Unit,AsOfDate,SourceType,Source,DisclosureStatus,Notes
XYZ,Example Fund,RevenueCAGR,0.08,decimal,2026-06-30,FundLetter,Primary-source URL,Disclosed,Manager explicitly discussed the growth assumption
```

Useful `SourceType` values:

- `13F`
- `AnnualReport`
- `ShareholderLetter`
- `FundLetter`
- `InvestorPresentation`
- `EarningsCall`
- `RegulatoryFiling`
- `Interview`
- `SecondaryResearch`

Useful `DisclosureStatus` values:

- `Disclosed` — the institution explicitly stated the value
- `CalculatedFromDisclosure` — calculated directly from disclosed figures
- `Inferred` — interpretation rather than a stated figure

Prefer `Disclosed` and `CalculatedFromDisclosure`. Keep `Inferred` rows clearly labeled.

## Berkshire Hathaway comparison

Berkshire is especially useful for evaluating business quality, capital allocation and valuation discipline, but the comparison must stay evidence-based.

### What can be compared directly

Depending on what is publicly available:

- whether Berkshire reports a position
- disclosed holding size / reported market value
- Berkshire annual-report and shareholder-letter commentary
- business economics discussed publicly by Buffett or Berkshire
- observable quality metrics such as returns on capital, margins, leverage and cash generation

### What usually cannot be compared directly

Berkshire generally does not publish its complete internal model for an individual public-equity investment. Do not treat any of the following as Berkshire's view unless a primary source explicitly supports it:

- exact intrinsic value
- exact DCF forecast
- exact WACC
- exact terminal growth assumption
- exact expected IRR
- exact required margin of safety

A 13F is evidence of a reportable holding at the filing date. It is **not** evidence of Berkshire's target price or intrinsic-value estimate.

## Buffett-style scorecard

`buffett_style_scorecard.csv` is a systematic public-data screen inspired by recurring Buffett/Berkshire themes. It is not a disclosed Berkshire formula.

The current screen evaluates ten observable areas:

1. Return on equity
2. Operating margin
3. Free-cash-flow margin
4. FCF-to-net-income cash conversion
5. Debt-to-equity
6. Net debt / EBITDA
7. Revenue growth
8. Earnings growth
9. FCF yield
10. EV / EBITDA

The output includes:

- `BuffettStyleScore10`
- `DataCoveragePct`
- the underlying metrics
- an explicit framework disclaimer

A high score means the company looks attractive under this particular screen. It does **not** mean Buffett would buy it.

## Direct institutional comparison

`institutional_comparison.csv` compares each institution separately:

- your value
- institutional value
- absolute difference
- percentage difference
- institution and source
- disclosure status

This is useful for seeing exactly where your model disagrees with a named investor or research source.

## Institutional consensus comparison

`institutional_consensus.csv` aggregates all entered institutional observations for the same ticker/metric and reports:

- number of institutional observations
- mean
- median
- minimum
- maximum
- your difference versus the mean

This should not be described as Wall Street consensus unless the input dataset is actually a representative analyst-consensus dataset.

## What could prove me wrong?

Use `thesis_risks.csv` before entering or increasing a position.

```csv
Ticker,Thesis,KeyRisk,FalsificationCondition,MonitoringMetric,Threshold,ReviewDate,Status,Notes
XYZ,Cloud growth supports valuation,Growth slows structurally,Two consecutive quarters below threshold,CloudRevenueGrowth,<0.10,2026-11-01,Open,Reassess long-term revenue CAGR
```

The purpose is to separate a changing stock price from a changing investment thesis.

Good falsification conditions should be:

- observable
- specific
- linked to the original thesis
- defined before the evidence arrives

## Recommended professional workflow

```text
Screen
  -> Company research
  -> Forecasts and valuation
  -> Your thesis inputs
  -> Institutional-source comparison
  -> Buffett-style quality screen
  -> Falsification conditions
  -> Position sizing
  -> Portfolio risk analysis
  -> Monitoring and thesis review
  -> Decision journal
```

## Source discipline

For institutional comparisons, use primary sources whenever possible. Examples include regulatory filings, Berkshire annual reports, shareholder letters, fund letters and official investor materials.

Do not manufacture a comparable number because a manager owns the stock. If the professional investor has not disclosed an assumption, leave that assumption absent and compare only what is actually observable.

This makes the comparison more useful: disagreement is meaningful only when both sides of the comparison are based on real information.
