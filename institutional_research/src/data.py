from __future__ import annotations

import time
from typing import Dict, Iterable

import pandas as pd
import yfinance as yf


def download_prices(tickers: Iterable[str], period: str = "5y") -> pd.DataFrame:
    tickers = list(dict.fromkeys(t.upper().strip() for t in tickers if str(t).strip()))
    if not tickers:
        raise ValueError("No tickers supplied.")

    raw = yf.download(
        tickers=tickers,
        period=period,
        auto_adjust=True,
        progress=False,
        threads=True,
        group_by="column",
    )

    if raw.empty:
        raise RuntimeError("No price history returned by Yahoo Finance.")

    if isinstance(raw.columns, pd.MultiIndex):
        if "Close" not in raw.columns.get_level_values(0):
            raise RuntimeError("Price download did not contain Close data.")
        close = raw["Close"].copy()
    else:
        close = raw[["Close"]].copy()
        close.columns = tickers[:1]

    if isinstance(close, pd.Series):
        close = close.to_frame(name=tickers[0])

    close = close.dropna(how="all").sort_index()
    close.columns = [str(c).upper() for c in close.columns]
    return close


def fetch_info(tickers: Iterable[str], pause: float = 0.05) -> Dict[str, dict]:
    out: Dict[str, dict] = {}
    for ticker in dict.fromkeys(str(t).upper().strip() for t in tickers):
        if not ticker:
            continue
        try:
            out[ticker] = yf.Ticker(ticker).info or {}
        except Exception as exc:
            out[ticker] = {"_error": str(exc)}
        if pause:
            time.sleep(pause)
    return out
