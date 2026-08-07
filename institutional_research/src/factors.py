from __future__ import annotations

import numpy as np
import pandas as pd


def _zscore(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    if s.notna().sum() < 2 or s.std(ddof=0) == 0:
        return pd.Series(0.0, index=s.index)
    return (s - s.mean()) / s.std(ddof=0)


def _momentum(prices: pd.Series) -> float:
    p = prices.dropna()
    if len(p) >= 252:
        return float(p.iloc[-21] / p.iloc[-252] - 1)
    if len(p) >= 126:
        return float(p.iloc[-1] / p.iloc[-126] - 1)
    if len(p) >= 63:
        return float(p.iloc[-1] / p.iloc[-63] - 1)
    return np.nan


def build_factor_scores(
    holdings: pd.DataFrame,
    prices: pd.DataFrame,
    info: dict[str, dict],
    factor_weights: dict[str, float],
    trading_days: int = 252,
) -> pd.DataFrame:
    rows = []
    for ticker in holdings["Ticker"]:
        d = info.get(ticker) or {}
        market_cap = d.get("marketCap")
        fcf = d.get("freeCashflow")
        fwd_pe = d.get("forwardPE")
        earnings_yield = 1 / fwd_pe if isinstance(fwd_pe, (int, float)) and fwd_pe > 0 else np.nan
        fcf_yield = fcf / market_cap if market_cap and fcf is not None else np.nan
        op_margin = d.get("operatingMargins")
        roe = d.get("returnOnEquity")
        revenue_growth = d.get("revenueGrowth")
        earnings_growth = d.get("earningsGrowth")
        momentum = _momentum(prices[ticker]) if ticker in prices.columns else np.nan
        vol = prices[ticker].pct_change().std(ddof=1) * np.sqrt(trading_days) if ticker in prices.columns else np.nan

        rows.append({
            "Ticker": ticker,
            "EarningsYield": earnings_yield,
            "FCFYield": fcf_yield,
            "OperatingMargin": op_margin,
            "ROE": roe,
            "RevenueGrowth": revenue_growth,
            "EarningsGrowth": earnings_growth,
            "Momentum": momentum,
            "AnnualizedVolatility": vol,
        })

    df = pd.DataFrame(rows).set_index("Ticker")

    df["ValueScore"] = (_zscore(df["EarningsYield"]) + _zscore(df["FCFYield"])) / 2
    df["QualityScore"] = (_zscore(df["OperatingMargin"]) + _zscore(df["ROE"])) / 2
    df["GrowthScore"] = (_zscore(df["RevenueGrowth"]) + _zscore(df["EarningsGrowth"])) / 2
    df["MomentumScore"] = _zscore(df["Momentum"])
    df["LowVolatilityScore"] = _zscore(-df["AnnualizedVolatility"])

    mapping = {
        "value": "ValueScore",
        "quality": "QualityScore",
        "growth": "GrowthScore",
        "momentum": "MomentumScore",
        "low_volatility": "LowVolatilityScore",
    }
    total_weight = sum(float(factor_weights.get(k, 0)) for k in mapping)
    if total_weight <= 0:
        total_weight = 1.0

    composite = 0.0
    for key, col in mapping.items():
        composite = composite + df[col].fillna(0) * float(factor_weights.get(key, 0)) / total_weight
    df["CompositeScore"] = composite

    weights = holdings.set_index("Ticker")["Weight"]
    df["PortfolioWeight"] = weights.reindex(df.index)
    df["WeightedCompositeContribution"] = df["CompositeScore"] * df["PortfolioWeight"]

    return df.reset_index()


def factor_exposure(factor_scores: pd.DataFrame) -> pd.DataFrame:
    factor_cols = ["ValueScore", "QualityScore", "GrowthScore", "MomentumScore", "LowVolatilityScore"]
    rows = []
    w = factor_scores["PortfolioWeight"].fillna(0)
    for col in factor_cols:
        rows.append({
            "Factor": col.replace("Score", ""),
            "PortfolioExposure": float((factor_scores[col].fillna(0) * w).sum()),
        })
    return pd.DataFrame(rows)
