from __future__ import annotations

import numpy as np
import pandas as pd


def max_drawdown(returns: pd.Series) -> float:
    wealth = (1 + returns.fillna(0)).cumprod()
    peak = wealth.cummax()
    drawdown = wealth / peak - 1
    return float(drawdown.min())


def portfolio_risk(
    asset_returns: pd.DataFrame,
    weights: pd.Series,
    benchmark_returns: pd.Series,
    risk_free_rate: float = 0.04,
    trading_days: int = 252,
):
    common = asset_returns.index.intersection(benchmark_returns.index)
    r = asset_returns.loc[common].dropna(how="any")
    b = benchmark_returns.loc[r.index].dropna()
    r = r.loc[b.index]

    w = weights.reindex(r.columns).astype(float)
    w = w / w.sum()
    portfolio_returns = r.mul(w, axis=1).sum(axis=1)

    ann_return = (1 + portfolio_returns).prod() ** (trading_days / max(len(portfolio_returns), 1)) - 1
    ann_vol = portfolio_returns.std(ddof=1) * np.sqrt(trading_days)
    downside = portfolio_returns[portfolio_returns < 0].std(ddof=1) * np.sqrt(trading_days)

    sharpe = (ann_return - risk_free_rate) / ann_vol if ann_vol > 0 else np.nan
    sortino = (ann_return - risk_free_rate) / downside if downside and downside > 0 else np.nan

    var95 = float(portfolio_returns.quantile(0.05))
    tail = portfolio_returns[portfolio_returns <= var95]
    es95 = float(tail.mean()) if not tail.empty else np.nan

    beta = np.cov(portfolio_returns, b, ddof=1)[0, 1] / np.var(b, ddof=1) if np.var(b, ddof=1) > 0 else np.nan
    # Jensen/CAPM alpha: portfolio excess return minus beta times benchmark excess return.
    rf_daily = (1 + float(risk_free_rate)) ** (1 / trading_days) - 1
    alpha_daily = (
        (portfolio_returns - rf_daily).mean() - beta * (b - rf_daily).mean()
        if np.isfinite(beta) else np.nan
    )
    alpha_ann = alpha_daily * trading_days if np.isfinite(alpha_daily) else np.nan

    covariance = r.cov() * trading_days
    cov_arr = covariance.to_numpy()
    w_arr = w.to_numpy()
    port_variance = float(w_arr.T @ cov_arr @ w_arr)
    port_vol = np.sqrt(max(port_variance, 0.0))
    marginal = cov_arr @ w_arr / port_vol if port_vol > 0 else np.full(len(w_arr), np.nan)
    component = w_arr * marginal
    pct_contribution = component / port_vol if port_vol > 0 else np.full(len(w_arr), np.nan)

    contribution = pd.DataFrame({
        "Ticker": r.columns,
        "Weight": w.values,
        "MarginalRisk": marginal,
        "ComponentRisk": component,
        "RiskContributionPct": pct_contribution,
    })

    summary = {
        "annualized_return": float(ann_return),
        "annualized_volatility": float(ann_vol),
        "sharpe": float(sharpe) if np.isfinite(sharpe) else None,
        "sortino": float(sortino) if np.isfinite(sortino) else None,
        "beta": float(beta) if np.isfinite(beta) else None,
        "annualized_alpha": float(alpha_ann) if np.isfinite(alpha_ann) else None,
        "alpha_method": "Jensen/CAPM alpha vs configured benchmark",
        "max_drawdown": max_drawdown(portfolio_returns),
        "daily_var_95": var95,
        "daily_expected_shortfall_95": es95,
        "observations": int(len(portfolio_returns)),
    }
    return summary, portfolio_returns, covariance, contribution


def per_asset_risk(
    prices: pd.DataFrame,
    benchmark: str,
    trading_days: int = 252,
) -> pd.DataFrame:
    returns = prices.pct_change().dropna(how="all")
    b = returns[benchmark].dropna()
    rows = []
    for ticker in [c for c in returns.columns if c != benchmark]:
        x = returns[ticker].dropna()
        common = x.index.intersection(b.index)
        x = x.loc[common]
        bb = b.loc[common]
        vol = x.std(ddof=1) * np.sqrt(trading_days)
        beta = np.cov(x, bb, ddof=1)[0, 1] / np.var(bb, ddof=1) if len(common) > 2 and np.var(bb, ddof=1) > 0 else np.nan
        if len(x) > 21:
            total = (1 + x).prod() - 1
        else:
            total = np.nan
        rows.append({
            "Ticker": ticker,
            "AnnualizedVolatility": vol,
            "Beta": beta,
            "HistoricalReturn": total,
            "MaxDrawdown": max_drawdown(x) if not x.empty else np.nan,
        })
    return pd.DataFrame(rows)
