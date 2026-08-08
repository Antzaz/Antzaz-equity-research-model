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

Static-weight attribution is a research diagnostic. True realized attribution requires transaction history and point-in-time weights.

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
