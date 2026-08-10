# Alpha Analysis

The Institutional Research Lab now separates raw benchmark outperformance from risk- and factor-adjusted residual return.

## Models

### CAPM / Jensen alpha
Uses the configured benchmark (SPY by default) and configured annual risk-free rate.

`portfolio excess return = alpha + beta * benchmark excess return + residual`

This replaces the previous simplified beta-adjusted return calculation in the portfolio summary.

### Fama-French 3
Uses Kenneth French daily U.S. research factors:
- Mkt-RF
- SMB
- HML

### Carhart 4
Adds the Kenneth French daily momentum factor (Mom) to FF3.

### Fama-French 5
Uses:
- Mkt-RF
- SMB
- HML
- RMW
- CMA

### Public Style Proxy
Uses the project's configured ETF proxies as a practical style diagnostic. The broad-market proxy is used as the market excess-return control; small-cap, value, growth, momentum, quality and low-volatility ETFs are expressed relative to the market proxy.

This model is useful for asking whether apparent alpha is really exposure to familiar public styles, but it is not a substitute for a commercial Barra/Axioma/MSCI risk model.

## Statistical diagnostics

Each full-sample model exports:
- annualized alpha
- annualized alpha standard error
- alpha t-statistic
- two-sided p-value
- 5% significance flag
- R-squared
- residual volatility
- observations and sample dates
- factor betas and factor beta t-statistics

A positive alpha is not automatically evidence of persistent skill. A weak t-statistic means the estimate is noisy.

## Rolling analysis

Configured in `config.json` under `alpha_analysis.rolling_windows`.

Defaults:
- 1Y = 252 trading days
- 3Y = 756 trading days

Rolling CAPM and Carhart alpha are sampled weekly in the output to keep files compact.

## Outputs

`outputs/latest/` now includes:
- `alpha_summary.csv`
- `alpha_factor_loadings.csv`
- `rolling_alpha.csv`
- `alpha_return_decomposition.csv`

The Streamlit multipage dashboard also includes `Alpha Analysis` under the `pages/` directory.

## Interpretation hierarchy

A useful way to read the project is:

1. **Raw active return** — did the portfolio beat the benchmark?
2. **CAPM alpha** — did it beat what broad-market beta would imply?
3. **FF3 / Carhart / FF5 alpha** — does the residual survive size, value, momentum, profitability and investment controls?
4. **Public Style Proxy alpha** — does the residual survive practical growth, momentum, quality and low-volatility proxy controls?
5. **Rolling alpha** — has the residual been persistent or concentrated in one period?
6. **Alpha t-statistic** — is the estimate large relative to its statistical uncertainty?

## Critical limitation

The current portfolio engine applies today's portfolio weights to historical asset returns unless point-in-time weights or transaction history are supplied. Therefore the alpha analysis is a **current-weight backcast diagnostic**, not a claim about realized historical manager alpha.

True realized alpha attribution requires historical portfolio weights or transactions through time, corporate-action-aware performance accounting, and preferably point-in-time factor / benchmark data.
