# Institutional Research Lab

A portfolio research, construction and validation layer designed to sit beside the single-company equity-research model.

The project now follows a professional-style workflow:

**investment intent → exposures → risk budget → constraints → portfolio construction → trade sizing → monitoring → attribution**

It uses public data, so it should be treated as an institutional-style research framework rather than an institutional data stack.

## Current analytics

### Portfolio and benchmark
- holdings, weights and unrealized P&L
- benchmark-relative return and beta
- annualized volatility, Sharpe and Sortino
- tracking error and information ratio
- daily active hit rate
- upside and downside capture
- maximum drawdown
- historical VaR and Expected Shortfall
- correlations and covariance

### Risk budgeting and concentration
- marginal and component contribution to risk
- capital weight vs risk contribution
- risk-to-capital ratio by holding
- top-1 / top-3 / top-5 concentration
- Herfindahl concentration index
- effective number of holdings
- sector concentration and effective number of sectors
- configurable portfolio-constraint monitor

### Factor and style diagnostics
- value, quality, growth, momentum and low-volatility holding-relative scores
- public ETF proxy sensitivities for market, size, value, growth, momentum, quality and low volatility

The ETF sensitivities are diagnostics, not a replacement for a commercial Barra/Axioma/MSCI factor risk model.

### Portfolio construction
The engine can calculate:
- minimum-variance portfolio
- equal-risk-contribution portfolio
- expected-return / max-Sharpe portfolio when `expected_returns.csv` is supplied
- current vs target weights
- one-way turnover
- target market values
- estimated trade values and share changes

Position and other risk limits are controlled in `config.json`.

### Liquidity / capacity
- average daily share volume
- average daily dollar volume
- position as % of ADV
- estimated days to liquidate at a configurable participation rate

This is a first-pass public-data liquidity screen, not an institutional transaction-cost model.

### Stress and scenario analysis
- configurable beta-based forward shocks
- configurable historical stress windows
- rolling 1M / 3M / 6M / 1Y returns, volatility and tracking error
- bootstrap Monte Carlo

### Attribution and decision discipline
- static-weight arithmetic return contribution by security
- benchmark-relative portfolio metrics
- analyst forecast-error tracking
- optional portfolio decision journal template

Static-weight attribution remains available as a research diagnostic.

## Realized portfolio & decision loop

The offline/private portfolio can now close the loop from research to realized outcomes:

- optional `transaction_ledger.csv` reconstructs daily positions, cash and point-in-time weights;
- true TWR removes explicit deposits/withdrawals and MWR/XIRR uses dated external cash flows;
- raw closes plus provider dividends/splits are used for transaction accounting when available;
- realized security contribution uses actual historical weights rather than today's weights;
- optional `benchmark_sector_history.csv` unlocks Brinson allocation / selection / interaction attribution without backfilling today's benchmark sector weights;
- decision-journal entries are evaluated at 3M / 6M / 12M and summarized by decision and conviction;
- the latest local equity-research workbook for each holding feeds a confidence-shrunk expected return into the optimizer unless a manual expected return overrides it;
- position-sizing ranges compare capital weight, risk contribution, model expected alpha, downside and research confidence;
- thesis-budget diagnostics compare capital weight, risk weight and expected-alpha weight;
- bottom-up portfolio scenarios aggregate company-model bear/base/P90 valuations rather than applying one portfolio beta;
- expected return is decomposed into fundamental growth, dividend yield, net buyback yield and valuation/other residual;
- transaction-cost-aware rebalancing compares expected annual benefit with an explicit spread + ADV impact estimate.

Copy the templates rather than inventing missing history:

```powershell
Copy-Item transaction_ledger_template.csv transaction_ledger.csv
Copy-Item portfolio_decision_journal_template.csv portfolio_decision_journal.csv
Copy-Item fundamental_scenarios_template.csv fundamental_scenarios.csv
# Only when you have point-in-time benchmark sector data:
Copy-Item benchmark_sector_history_template.csv benchmark_sector_history.csv
```

If these files are absent, the related analytics remain REVIEW/unavailable rather than substituting current holdings or current benchmark weights. Position-sizing ranges also apply an explicit liquidity haircut when estimated liquidation days exceed the configured limit, while RiskContributionPct already carries the covariance/correlation effect of the holding inside the current portfolio.

### Valuation
- simplified reverse DCF / market-implied FCF growth

## Private portfolio inputs

Copy the templates below to the private filenames shown. The private files are ignored by Git.

### Holdings

```powershell
Copy-Item portfolio_template.csv portfolio.csv
```

Use shares:

```csv
Ticker,Shares,AverageCost,ManualWeight,Notes
GOOGL,10,200,,Core position
MSFT,5,350,,Quality compounder
```

or manual weights:

```csv
Ticker,Shares,AverageCost,ManualWeight,Notes
GOOGL,,,0.60,Core position
MSFT,,,0.40,Quality compounder
```

### Expected returns and conviction

```powershell
Copy-Item expected_returns_template.csv expected_returns.csv
```

`ExpectedReturn` should be your forward annual expected return, not a historical average. `MinWeight` and `MaxWeight` are optional security-specific sizing constraints.

Max-Sharpe optimization is disabled unless expected returns are supplied for every holding. This avoids silently optimizing on historical returns as though they were forecasts.

### Active Share

```powershell
Copy-Item benchmark_weights_template.csv benchmark_weights.csv
```

Populate it with actual benchmark constituent weights. Active Share is not calculated unless this file is present and valid.

### Decision journal

```powershell
Copy-Item portfolio_decision_journal_template.csv portfolio_decision_journal.csv
```

Use it before trades to record the sizing decision, expected return, conviction, key risk, falsification condition and review date.

## Install

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research"
python -m pip install -r requirements.txt
```

## Run

```powershell
python run_research.py
```

Outputs are written to:

```text
outputs/latest/
outputs/snapshots/YYYYMMDD_HHMMSS/
```

Snapshots preserve how the portfolio analysis looked at different dates.

## Dashboard

```powershell
python -m streamlit run dashboard.py
```

or double-click `start_dashboard.bat`.

The dashboard now contains dedicated views for:
- portfolio
- risk
- construction / constraints
- factors
- attribution
- liquidity
- Monte Carlo
- reverse DCF
- forecast accuracy

## Important configuration

`config.json` contains:
- benchmark
- history period
- risk-free rate
- factor weights and public factor proxies
- maximum position / sector / risk-contribution limits
- beta and tracking-error limits
- liquidity participation assumptions
- stress scenarios
- historical crisis windows

The defaults are examples, not recommendations. Set them to match your own investment policy and risk capacity.

## Professional-quality limitations

Yahoo Finance is convenient but is not point-in-time institutional data. This project should not be used to claim a clean historical factor backtest or precise institutional attribution.

A professional-grade next step would require point-in-time data with:
- delisted companies
- historical index membership and benchmark constituent weights
- filing-availability dates and restatements
- corporate actions
- historical analyst estimates and revisions
- transaction history and actual portfolio weights through time
- bid/ask spreads, market impact and transaction costs
- security-level currencies and FX hedges
- tax lots and tax-aware optimization where relevant

Commercial or academic-quality sources may include CRSP/Compustat, FactSet, Bloomberg, LSEG, MSCI/Barra, Axioma or equivalent point-in-time databases.

## Future roadmap

### Phase 2 — point-in-time factor backtesting
- monthly cross-sectional universe
- value / quality / growth / momentum signals
- walk-forward testing
- sector-neutral portfolios
- transaction costs and turnover
- bootstrap confidence intervals
- parameter perturbation
- sub-period and regime testing

### Phase 3 — richer market expectations
- point-in-time consensus estimates
- estimate revisions and breadth
- earnings surprises
- options-implied volatility / skew
- positioning / short-interest data

### Phase 4 — advanced portfolio construction
- sector-aware constrained optimizer
- benchmark-factor exposure constraints
- expected-return confidence / Bayesian shrinkage
- Black-Litterman-style view integration
- transaction-cost-aware rebalancing
- tax-aware optimization
- currency and FX-risk budgeting

### Phase 5 — realized portfolio attribution
- transaction ledger
- time-weighted and money-weighted returns
- realized security and sector attribution
- allocation / selection / interaction attribution
- realized turnover and implementation shortfall
- decision-journal outcome analytics


### Research → portfolio decision loop

The offline dashboard now exposes two additional private pages:

- **Realized Portfolio** — reconstructs point-in-time positions from `transaction_ledger.csv`, calculates TWR/MWR, realized security contribution and optional Brinson allocation/selection/interaction.
- **Research Decision Loop** — links the newest local company workbooks to expected returns, confidence, position-size ranges, capital/risk/alpha budgets, fundamental scenarios, transaction-cost-aware rebalance gates and decision-journal learning.

Optional private inputs:

```powershell
Copy-Item transaction_ledger_template.csv transaction_ledger.csv
Copy-Item portfolio_decision_journal_template.csv portfolio_decision_journal.csv
Copy-Item fundamental_scenarios_template.csv fundamental_scenarios.csv
Copy-Item benchmark_sector_history_template.csv benchmark_sector_history.csv
```

The decision journal supports optional `Sector`, `ThesisCategory` and `Catalyst` fields. Mature 3M/6M/12M outcomes are summarized by decision type, conviction bucket, sector and thesis category.

The automatic research expected-return bridge decomposes expected return into:
- annualized convergence from current price toward the latest local base fair value;
- dividend yield;
- bounded net share-reduction / buyback yield from the Capital Allocation sheet;
- confidence shrinkage based on data quality, model score and historical forecast accuracy.

Manual `expected_returns.csv` remains authoritative when you provide it.


## Offline investment-decision loop

The private portfolio layer can now close the loop from company research to realized outcomes. These files stay local and are ignored by Git:

```powershell
Copy-Item transaction_ledger_template.csv transaction_ledger.csv
Copy-Item portfolio_decision_journal_template.csv portfolio_decision_journal.csv
Copy-Item fundamental_scenarios_template.csv fundamental_scenarios.csv
Copy-Item benchmark_sector_history_template.csv benchmark_sector_history.csv
```

After populating the private files, run:

```powershell
python run_research.py
python -m streamlit run app.py
```

The **Realized Portfolio** page shows point-in-time positions, TWR, MWR/XIRR, realized security contribution and Brinson attribution when historical benchmark-sector data are supplied. The **Research Decision Loop** page links private company workbooks to confidence-adjusted expected returns, position-sizing ranges, capital/risk/expected-alpha budgets, bottom-up scenarios, transaction-cost-aware rebalance gates and matured decision-journal outcomes.

The company workbook itself contains private/offline accountability layers such as **Forecast Accountability**, **Earnings & Revisions**, **Capital Allocation**, **Valuation History**, **SOTP Framework** and **Thesis Timeline**. These are not exported to the public recruiter showcase.
