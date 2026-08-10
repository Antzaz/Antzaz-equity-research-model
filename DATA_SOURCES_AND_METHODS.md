# Data Sources & Professional Research Methods

This project separates **facts**, **consensus observations**, **model calculations**, **analyst assumptions**, and **professional-framework reference anchors**. That distinction is intentional: a workbook should never look more certain than the evidence supports.

## 1. Source hierarchy

### Reported company facts — highest priority
1. Issuer investor-relations pages, annual/quarterly reports and earnings materials.
2. SEC EDGAR / Company Facts or equivalent regulator filings.
3. Transparent market-data fallbacks when primary sources are unavailable.

### Market / industry context
Specialist sources are used only where their market definition is comparable. Examples in the registry include TrendForce (foundry), Synergy Research Group (cloud), IDC (devices/PCs), StatCounter (search/OS usage) and WSTS (semiconductor market context).

### Analyst consensus
The project supports a provider stack rather than a single consensus feed:

| Provider | Default? | What it can add | Environment variable |
|---|---:|---|---|
| Yahoo Finance / yfinance | Yes | Revenue, EPS, analyst ranges, EPS revisions | none |
| Financial Modeling Prep | Optional | Revenue, EPS, EBIT, EBITDA and other financial estimates; price-target consensus | `FMP_API_KEY` |
| Alpha Vantage | Optional | Annual/quarterly revenue and EPS estimates, analyst count/revision evidence | `ALPHAVANTAGE_API_KEY` |
| Finnhub | Optional / plan-dependent | Revenue, EPS, EBIT, EBITDA, FCF, OCF and capex estimates | `FINNHUB_API_KEY` |

**Important:** aggregator analyst universes can overlap. The project therefore does **not** add analyst counts across providers. For the same metric/year it uses the median of available provider consensus means as a cross-provider reference, while preserving every provider row in `Expectations & Consensus`.

Optional keys can be set in PowerShell for a session, for example:

```powershell
$env:FMP_API_KEY="your-key"
$env:ALPHAVANTAGE_API_KEY="your-key"
$env:FINNHUB_API_KEY="your-key"
python .\research.py GOOGL
```

If a key or paid endpoint is unavailable, the provider is shown as **not configured / unavailable** rather than silently replaced with invented data.

## 2. Market Expectations vs Consensus

These are different questions and should not be mixed:

- **Consensus** asks: what are analysts currently forecasting for revenue, EPS, EBIT margin, FCF, capex, etc.?
- **Market Expectations** asks: what growth, margin, capital intensity and competitive-advantage duration are required to justify the current share price?

The workbook now places external consensus beside reverse-DCF hurdles. This makes variant perception explicit: your opportunity is often not that a company is objectively good, but that your justified forecast differs from what analysts and/or the market price already imply.

## 3. Institutional Comparison — where the scores come from

The named institutional scores are **not proprietary scores from Berkshire Hathaway, Fundsmith, Dodge & Cox, Davis Advisors, Pershing Square, Baillie Gifford, Akre, Polen, Capital Group or T. Rowe Price**.

They are transparent project-created lenses built in three steps:

1. The workbook creates reusable 0–100 company dimensions from its own evidence: Growth, Profitability, FCF Quality, Balance Sheet, Absolute Valuation, Relative Valuation, Stress Robustness, Leadership and Moat/Position.
2. Each institutional style receives an explicit weight map derived from its publicly described investment philosophy.
3. `Fit Score = sum(company dimension score × lens weight) / sum(available weights)`.

If a company dimension is unavailable, it is excluded and the sheet reports the **Data Coverage** actually represented. Fit labels and thresholds are also this project's rubric, not labels used by the named firms.

The workbook's `Institutional Comparison` tab now contains:
- the score formula and caveat;
- the provenance/construction of every company dimension;
- the full institution-by-dimension weight matrix;
- the public philosophy source for every lens;
- the strongest dimensions and challenge dimensions.

## 4. Professional analysis methods incorporated

### Competitive Advantage Period (CAP)
Morgan Stanley Investment Management / Counterpoint Global emphasizes that value creation depends on the spread between ROIC and the cost of capital, how much a business can reinvest, and how long it can sustain that spread. The project uses this in `Market Expectations` to avoid treating terminal growth as an arbitrary plug.

### Bayes + reference-class base rates
Start with a historical reference-class prior and update it with company-specific evidence. This is especially useful for extreme growth forecasts and AI narratives, where extrapolation can otherwise dominate judgment.

### Reverse DCF / expectations investing
Instead of only asking “what is my fair value?”, solve for the revenue growth, FCF growth, margins, capex intensity or competitive-advantage duration required by the current price. Compare those hurdles with your own forecast and external consensus.

### Incremental ROIC and capital allocation
Historic ROIC is useful, but the economics of **new** reinvestment drive future value creation. The project therefore links growth, capex, cash generation and incremental returns rather than rewarding growth on its own.

### Consensus dispersion and revisions
A consensus average without its low/high range, analyst coverage, provider disagreement and revision direction hides information. These are now explicit in `Expectations & Consensus`.

### Segment / SOTP analysis
Where businesses have materially different growth, margins, capital intensity or competitive position, consolidated multiples can obscure value. Segment economics should be underwritten separately before rolling into consolidated valuation.

### Earnings quality / cash conversion
Net income and EPS are reconciled against operating cash flow, capex, SBC and working capital. FCF economics receive explicit weight in both the investment score and institutional lenses.

## 5. AI Impact Analysis — professional estimation logic

The AI sheet should answer **how much shareholder value AI can create or destroy**, not whether AI is strategically exciting.

The framework separates:

1. **Exposure** — what share of revenue/cost base is actually addressable by AI?
2. **Adoption / attach** — what share of users, workloads, seats or transactions use the AI product or process?
3. **Monetization** — price, ARPU, conversion, share gains, attach revenue or AI-related demand.
4. **Gross-margin capture** — AI revenue less inference/serving/model/data costs.
5. **Productivity capture** — labor/opex exposure × automation potential × adoption × realized savings.
6. **Capital burden** — incremental compute/data-center/power/network capex and depreciation.
7. **Financing / scarcity** — whether capital needs increase leverage or the cost of capital, and where bottlenecks shift bargaining power.
8. **Incremental AI FCF / ROIC** — whether the combined revenue and cost benefits earn an attractive return after the capital burden.
9. **Expectations gap** — whether the outcome is already embedded in the Base DCF and market price.

Professional reference methods shown in the workbook include:
- McKinsey's bottom-up use-case approach to measurable revenue/cost outcomes;
- Goldman Sachs Research's labor-intensity × AI-task-exposure productivity framework and capex-to-revenue investor focus;
- BlackRock Investment Institute's AI value-capture, scarcity, financing and capex-vs-eventual-revenue framework;
- Goldman Sachs' translation of productivity scenarios into EPS growth and valuation effects.

These are **reference anchors, not automatic company assumptions**. Company-specific adoption, pricing, value capture and capital intensity remain explicit analyst inputs unless the company reports them.

## 6. Rule for adding future sources

A new source should be classified before it enters a model:

- `PRIMARY`: issuer/regulator fact.
- `CONSENSUS`: analyst aggregation or estimate feed.
- `SPECIALIST`: industry/market measurement.
- `PROFESSIONAL_FRAMEWORK`: research method or benchmark.
- `FALLBACK`: market-data source used when primary evidence is unavailable.

The source must also record its market definition, period, units/currency and whether the number is observed, calculated or assumed. This is the best defense against false precision as the project grows.
