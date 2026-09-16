# Machine Learning Research Layer

The ML layer is a second-opinion system around the deterministic equity-research model. It never overwrites DCF assumptions, reported financials, consensus inputs, thesis decisions, or portfolio trades.

## Run

Build the normal research workbook and then run the core company ML models:

```powershell
python .\research.py GOOGL --ml
```

Reuse an existing workbook and run only the company ML layer:

```powershell
python .\ml_research.py GOOGL
```

Run the persistent continual-learning, short-horizon, risk-forecast and research-loop stack:

```powershell
python -m machine_learning.learning_runner
```

The standard ML runner uses no OpenAI API calls or tokens.

## Visual ML Research Lab

Launch the existing institutional-research Streamlit app:

```powershell
python -m streamlit run institutional_research/app.py
```

Then open **ML Research Lab** from the sidebar. The page reads the same local `ml_data/ml_history.sqlite` evidence used by the scheduled learning workflow and exposes:

- current **1D / 1W / 1M / 3M / 6M / 12M** excess-return forecasts versus SPY;
- per-company horizon charts and plain-English forecast explanations;
- expected **1W realized volatility** and **1M forward drawdown**;
- live matured-vs-pending forecast counts;
- champion/challenger governance, MAE skill versus baseline, directional accuracy, IC, calibration and drift;
- stored walk-forward evidence for the component estimators;
- autonomous research hypotheses, challenger experiments, skeptic decisions and promotion candidates;
- point-in-time database health and freshness.

The page also provides two explicit research controls: **Refresh forecasts & learning** and **Run challenger research now**. These actions update research evidence only; they do not execute trades, rewrite production model code, or modify DCF assumptions.

If the local database has not been initialized yet, run:

```powershell
python ml_history.py daily-refresh --universe sp500 --limit 500 --years 1 --deep-years 20 --deep-batch 25
python -m machine_learning.learning_runner
```

The dashboard leaves unavailable evidence blank rather than fabricating forecasts or realized outcomes.

## Core company models

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
- Distance-based cluster weights are context, not calibrated probabilities or a market-timing instruction.

### 5. AI impact ML

- Ridge regression over the project's accumulated company-specific AI KPI history.
- The model will not train until there are at least eight dated KPI snapshots and enough usable forward observations.
- Before that threshold, the deterministic `AI Impact Analysis` economics bridge remains authoritative.

### 6. Portfolio ML / position sizing

- Uses ML expected-return evidence where available.
- Risk is estimated with Ledoit-Wolf shrinkage covariance from daily returns.
- Long-only constrained optimizer translates expected-return and covariance inputs into suggested weights subject to maximum-position and risk-aversion constraints.
- No trades are sent or executed.

## Multi-horizon return research

The continual-learning layer now adds five research-only forward excess-return horizons:

- **1D**: next trading-day excess return.
- **1W**: next five trading-day excess return.
- **1M**: forward 30-calendar-day excess return.
- **3M**: forward 91-calendar-day excess return.
- **6M**: forward 182-calendar-day excess return.

The 1D and 1W models use a fast point-in-time market feature panel rather than annual fundamentals alone. Features include:

- 1D / 5D / 21D stock returns;
- 1D / 5D / 21D excess returns;
- 5D / 21D / 63D realized volatility;
- 21D / 63D drawdown;
- 63D beta;
- 21D volume z-score;
- market momentum and volatility context.

The fast-horizon ensemble is histogram gradient boosting + Extra Trees + Elastic Net. The 1M/3M/6M models retain the existing point-in-time fundamental feature family and histogram-gradient-boosting/Elastic-Net architecture.

Short-horizon predictions are journaled into the persistent SQLite history and are closed only after the correct future trading/calendar horizon is observable. They do not feed the production portfolio optimizer automatically.

## Forward risk models

Two additional context-only models complement expected-return forecasts:

### Expected 1W Realized Volatility

Predicts annualized realized volatility over the next five trading days from the fast market feature panel.

### Expected 1M Forward Drawdown

Predicts the worst price decline from the current price over the next 21 trading days.

Both use a three-model ensemble (histogram gradient boosting, Extra Trees and Elastic Net), purged expanding walk-forward validation and explicit future target dates. They are research evidence, not automatic position-size overrides.

## Autonomous research loop

`machine_learning/research_loop.py` adds a bounded agent-style research cycle around the existing models.

The roles are:

1. **Research agent** — reads the live champion/challenger registry, drift and realized forecast skill.
2. **Experiment agent** — proposes approved estimator and feature-set challengers.
3. **Simulation agent** — runs leakage-safe expanding walk-forward experiments.
4. **Skeptic agent** — rejects weak sample sizes, baseline underperformance, poor directional skill and strongly negative out-of-sample fit.
5. **Promotion gate** — records strong challengers as promotion candidates for review.

The research loop currently experiments against the existing 12M return model and all 1D/1W/1M/3M/6M return horizons. Approved challenger families include:

- histogram gradient boosting;
- Extra Trees;
- Random Forest;
- Elastic Net;
- Ridge.

Experiments and hypotheses are stored in the local SQLite `research_experiments` table. The loop runs on a weekly cadence by default while daily prediction/maturation continues every day.

**Important governance rule:** the research loop never rewrites production model code, changes DCF assumptions, promotes a challenger by itself, or executes trades. Numerical gates identify candidates; production changes remain explicit and auditable.

## Persistent learning cycle

The scheduled daily workflow performs the following sequence:

```text
refresh point-in-time market/fundamental history
        -> generate 1D/1W/1M/3M/6M predictions
        -> generate forward risk forecasts
        -> mature forecasts whose outcomes are now observable
        -> score live forecast errors against baselines
        -> refresh champion/challenger governance
        -> run bounded weekly challenger research when due
        -> preserve encrypted learning state
```

Heavy holding-level model retraining remains governed separately on the monthly schedule.

## Outputs

Core company ML runs write:

- `ml_runs/<TICKER>/<timestamp>/ml_results.json` (gitignored)
- one consolidated `ML & Quantitative Research` workbook tab

Persistent learning additionally stores:

- predictions and realized outcomes in `ml_data/ml_history.sqlite`;
- training runs and model-performance history;
- model registry / influence multipliers;
- research experiments and skeptic decisions;
- daily short-horizon and risk-forecast state.

## Validation and leakage controls

- Point-in-time dates are explicit.
- Forward targets must occur after the feature observation date.
- 1D/1W targets are based on future trading observations, not calendar shortcuts.
- Expected-return models use expanding walk-forward testing.
- Training-set size controls reduce false confidence from tiny samples.
- Fast panels cap research-row counts to keep scheduled learning computationally bounded.
- AI impact refuses to train before its KPI-history threshold.
- Deterministic CI uses generated fixtures so repository health does not depend on Yahoo availability.
- Research challengers cannot silently replace production models.

## Interpretation

ML output is evidence, not an investment recommendation. A strong prediction should still be challenged against:

- DCF / scenario valuation;
- reverse DCF and market-implied expectations;
- consensus and revisions;
- balance-sheet and accounting quality;
- competitive advantage / moat;
- management and capital allocation;
- AI / technology disruption;
- institutional comparison;
- forward volatility / drawdown risk;
- portfolio concentration and risk.

A model that looks strong in-sample but fails walk-forward or live matured-forecast validation should not influence a decision.
