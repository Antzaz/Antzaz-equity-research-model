# Machine Learning Research Layer

The ML layer is a second-opinion system around the deterministic equity-research model. It never overwrites DCF assumptions, reported financials, consensus inputs, thesis decisions, or portfolio trades.

## Run

Build the normal research workbook and then run all six ML models:

```powershell
python .\research.py GOOGL --ml
```

Reuse an existing workbook and run only the ML layer:

```powershell
python .\ml_research.py GOOGL
```

The ML runner uses no OpenAI API calls or tokens.

## Six models

### 1. Expected 12-month excess return

- Ensemble: histogram gradient boosting (65%) + Elastic Net (35%).
- Target: 12-month stock return minus benchmark return.
- Features: revenue growth, operating/net/FCF margins, capex and R&D intensity, ROE, leverage, momentum, volatility and drawdown.
- Historical fundamentals are assigned a conservative 120-day availability lag after fiscal year-end before they may enter training.
- Validation: expanding walk-forward evaluation only; no random train/test split.

### 2. Consensus / earnings surprise

- Random Forest regression.
- Target: next reported EPS surprise versus the prevailing EPS estimate.
- Features use only prior surprise history, pre-earnings momentum/volatility and the contemporaneous estimate.
- Requires enough historical earnings observations; otherwise returns `INSUFFICIENT_DATA`.

### 3. Financial anomaly detection

- Isolation Forest on the company's historical operating profile.
- Features: growth, operating/net/FCF margins, capex, R&D and SBC intensity.
- Output is an anomaly percentile versus the company's available history, plus the largest standardized deviations.
- An anomaly is a diligence flag, not proof of accounting misconduct or deterioration.

### 4. Market regime classifier

- Five-cluster K-Means model using monthly market features.
- Inputs: equity, duration, credit, commodity and dollar momentum plus equity volatility.
- Regimes are labeled after clustering from the observed cluster characteristics.
- Distance-based cluster weights are reported as context, not as calibrated probabilities or a market-timing instruction.

### 5. AI impact ML

- Ridge regression over the project's accumulated company-specific AI KPI history.
- The model will not train until there are at least eight dated KPI snapshots and enough usable forward observations.
- Before that threshold, the deterministic `AI Impact Analysis` economics bridge remains authoritative.
- This prevents one or two management KPIs from being turned into a false machine-learning forecast.

### 6. Portfolio ML / position sizing

- Uses ML expected-return evidence where available.
- Risk is estimated with Ledoit-Wolf shrinkage covariance from daily returns.
- Long-only constrained optimizer translates expected-return and covariance inputs into suggested weights subject to maximum-position and risk-aversion constraints.
- No trades are sent or executed.
- If the private local `institutional_research/portfolio.csv` is absent, the model reports readiness instead of inventing a portfolio.

## Outputs

Each run writes:

- `ml_runs/<TICKER>/<timestamp>/ml_results.json` (gitignored)
- one consolidated `ML & Quantitative Research` workbook tab

The workbook dashboard shows all six models, status, prediction/state, validation evidence, top drivers, intended use and caveats. Detailed sections remain on the same sheet to avoid tab proliferation.

## Validation and leakage controls

- Point-in-time dates are explicit.
- Forward targets must occur after the feature observation date.
- Expected-return and earnings models use expanding walk-forward testing.
- Training-set size controls reduce false confidence from tiny samples.
- AI impact refuses to train before the KPI-history threshold.
- Deterministic CI uses generated fixtures so repository health does not depend on Yahoo availability.

## Interpretation

ML output is evidence, not an investment recommendation. A strong prediction should still be challenged against:

- DCF / scenario valuation
- reverse DCF and market-implied expectations
- consensus and revisions
- balance-sheet and accounting quality
- competitive advantage / moat
- management and capital allocation
- AI / technology disruption
- institutional comparison
- portfolio concentration and risk

A model that looks strong in-sample but fails walk-forward validation should not influence a decision.
