# Institutional Research Lab

A portfolio research and validation layer designed to sit beside the existing single-company equity-research model.

The first version focuses on analyses that are practical with public data and useful for an individual investor:

- portfolio holdings, weights and unrealized P&L
- benchmark-relative return and beta
- volatility, Sharpe, Sortino and alpha
- historical VaR and Expected Shortfall
- maximum drawdown
- correlation and covariance
- contribution to portfolio volatility
- portfolio-relative factor scores
- beta-based stress scenarios
- historical-bootstrap Monte Carlo
- simplified reverse DCF / market-implied FCF growth
- analyst forecast-error tracking
- flat CSV exports for Power BI
- interactive Streamlit dashboard

## 1. Add your portfolio

Copy `portfolio_template.csv` to `portfolio.csv`, then edit your local `portfolio.csv`.

You can use **either** shares:

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

If every row has positive `Shares`, the engine calculates market-value weights.
Otherwise every row must have a positive `ManualWeight`.

`portfolio.csv` and generated `outputs/` are ignored by Git so your real holdings stay local by default. If you intentionally want them in GitHub, remove those entries from `.gitignore`.

## 2. Install

From PowerShell:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research"
python -m pip install -r requirements.txt
```

## 3. Run the research engine

```powershell
python run_research.py
```

Outputs are written to:

```text
outputs/latest/
outputs/snapshots/YYYYMMDD_HHMMSS/
```

The snapshots let you preserve how the analysis looked at different dates.

## 4. Launch the interactive dashboard

```powershell
python -m streamlit run dashboard.py
```

Or double-click `start_dashboard.bat`.

The dashboard opens in your browser.

## 5. Power BI

The same analysis is exported as flat CSV files. See `powerbi/README.md`.

Power BI should be treated as the visualization layer; Python remains the analytics engine.

## Forecast tracking

Use `forecasts.csv` to record forecasts **before** the outcome is known.

Example:

```csv
Ticker,ForecastDate,FiscalYear,Metric,Forecast,Actual,Notes
GOOGL,2026-08-08,2027,Revenue,500,,Base-case forecast
```

After results are reported, enter `Actual` and rerun the engine. The dashboard will calculate
signed and absolute forecast errors and indicate whether you systematically over- or under-forecast.

## Stress scenarios

Edit `stress_scenarios` inside `config.json`.

Each scenario uses:

`estimated holding return = beta × benchmark shock + idiosyncratic shock`

This is deliberately transparent and is a first-pass portfolio stress framework, not a full nonlinear risk model.

## Reverse DCF

The first reverse-DCF implementation solves the constant annual FCF growth rate needed to reconcile
a company's current market cap with:

- current free cash flow
- cash and debt
- configured WACC
- configured terminal growth
- configured explicit forecast period

Edit these assumptions in `config.json`.

## Important limitations

This project is an **institutional-style research framework**, not an institutional data stack.

Yahoo Finance is convenient for live/public analysis, but it is not point-in-time fundamentals data.
Therefore this version should not be used to claim an unbiased historical factor backtest.

For true historical cross-sectional backtesting, the next phase should use point-in-time data with:
- delisted companies
- historical index membership
- filing availability dates
- restatement handling
- corporate actions
- realistic transaction costs

Typical professional-quality sources include CRSP/Compustat or a point-in-time commercial equity database.

## Roadmap

### Phase 2 — backtesting
- monthly cross-sectional universe
- value / quality / growth / momentum signals
- walk-forward testing
- sector-neutral portfolios
- transaction costs and turnover
- bootstrap confidence intervals
- parameter perturbation
- sub-period and regime testing

### Phase 3 — market expectations
- richer reverse DCF
- consensus estimates
- estimate revisions
- earnings surprises
- options-implied volatility / skew

### Phase 4 — portfolio construction
- optimizer with position/sector constraints
- factor exposure constraints
- expected return vs risk
- risk budgeting
- rebalancing / turnover controls
