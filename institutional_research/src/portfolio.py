from __future__ import annotations

from pathlib import Path
import numpy as np
import pandas as pd


REQUIRED_COLUMNS = ["Ticker", "Shares", "AverageCost", "ManualWeight", "Notes"]


def load_portfolio(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    df = pd.read_csv(path, comment="#")

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan

    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    df = df[df["Ticker"].ne("") & df["Ticker"].ne("NAN")].copy()

    for col in ["Shares", "AverageCost", "ManualWeight"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if df.empty:
        raise ValueError(
            f"{path} has no holdings. Add at least one ticker and either Shares "
            "or ManualWeight."
        )

    if df["Ticker"].duplicated().any():
        duplicates = ", ".join(df.loc[df["Ticker"].duplicated(), "Ticker"].tolist())
        raise ValueError(f"Duplicate tickers in portfolio.csv: {duplicates}")

    return df


def build_holdings(
    portfolio: pd.DataFrame,
    current_prices: pd.Series,
    info: dict[str, dict],
) -> pd.DataFrame:
    df = portfolio.copy()
    df["CurrentPrice"] = df["Ticker"].map(current_prices.to_dict())
    df["MarketValue"] = df["Shares"] * df["CurrentPrice"]

    use_market_weights = df["Shares"].notna().all() and (df["Shares"] > 0).all()
    if use_market_weights:
        total_value = df["MarketValue"].sum()
        if not np.isfinite(total_value) or total_value <= 0:
            raise ValueError("Could not calculate positive portfolio market value.")
        df["Weight"] = df["MarketValue"] / total_value
        weight_method = "market_value"
    else:
        if df["ManualWeight"].isna().any() or (df["ManualWeight"] <= 0).any():
            raise ValueError(
                "Either enter positive Shares for every holding, or enter a positive "
                "ManualWeight for every holding."
            )
        df["Weight"] = df["ManualWeight"] / df["ManualWeight"].sum()
        weight_method = "manual"

    df["CostValue"] = df["Shares"] * df["AverageCost"]
    df["UnrealizedPnL"] = df["MarketValue"] - df["CostValue"]
    df["UnrealizedPnLPct"] = np.where(
        df["CostValue"] > 0,
        df["UnrealizedPnL"] / df["CostValue"],
        np.nan,
    )

    def get_info(ticker: str, field: str):
        return (info.get(ticker) or {}).get(field)

    df["Company"] = df["Ticker"].map(lambda t: get_info(t, "longName") or t)
    df["Sector"] = df["Ticker"].map(lambda t: get_info(t, "sector"))
    df["Industry"] = df["Ticker"].map(lambda t: get_info(t, "industry"))
    df["Country"] = df["Ticker"].map(lambda t: get_info(t, "country"))
    df["Currency"] = df["Ticker"].map(lambda t: get_info(t, "currency"))
    df["MarketCap"] = df["Ticker"].map(lambda t: get_info(t, "marketCap"))
    df["ForwardPE"] = df["Ticker"].map(lambda t: get_info(t, "forwardPE"))
    df["RevenueGrowth"] = df["Ticker"].map(lambda t: get_info(t, "revenueGrowth"))
    df["OperatingMargin"] = df["Ticker"].map(lambda t: get_info(t, "operatingMargins"))
    df["ROE"] = df["Ticker"].map(lambda t: get_info(t, "returnOnEquity"))
    df.attrs["weight_method"] = weight_method
    return df
