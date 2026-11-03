from __future__ import annotations

"""Offline-readiness check for this Windows portfolio installation.

Only checks LOCAL files and installed modules. It never sends holdings or market data
to a provider and does not print share counts, cost basis or individual tickers.
"""

import importlib.util
import json
from pathlib import Path

import pandas as pd

from src.portfolio import load_portfolio
from src.data import cached_prices, price_cache_status


BASE = Path(__file__).resolve().parent
PORTFOLIO = BASE / "portfolio.csv"
SUMMARY = BASE / "outputs" / "latest" / "summary.json"
HOLDINGS = BASE / "outputs" / "latest" / "holdings_analysis.csv"


def check(base: Path = BASE) -> tuple[bool, list[str]]:
    notes: list[str] = []
    portfolio_file = base / "portfolio.csv"
    summary_file = base / "outputs" / "latest" / "summary.json"
    holdings_file = base / "outputs" / "latest" / "holdings_analysis.csv"
    if not portfolio_file.is_file():
        return False, ["Missing portfolio.csv. Pull the latest GitHub main branch while online."]
    try:
        holdings = load_portfolio(portfolio_file)
    except Exception as exc:
        return False, [f"Portfolio CSV is invalid: {type(exc).__name__}: {exc}"]
    count = len(holdings)
    notes.append(f"Portfolio loaded: {count} holding(s) with a valid ticker column.")
    if importlib.util.find_spec("streamlit") is None:
        notes.append("Streamlit is not installed for this Python interpreter.")
        return False, notes + ["Run install_desktop_app.bat while connected to the internet."]

    if not summary_file.is_file() or not holdings_file.is_file():
        return False, notes + [
            "No complete cached portfolio analytics exist yet.",
            "Connect once and choose Sync All Data in Research Hub to create offline outputs.",
        ]
    try:
        summary = json.loads(summary_file.read_text(encoding="utf-8"))
        previous = pd.read_csv(holdings_file)
    except (OSError, ValueError, pd.errors.ParserError) as exc:
        return False, notes + [f"Saved analytics are unreadable: {type(exc).__name__}."]
    if not isinstance(summary, dict) or "portfolio" not in summary:
        return False, notes + ["Saved summary.json is incomplete."]
    if "Ticker" not in previous:
        return False, notes + ["Cached holdings_analysis.csv has no Ticker column."]
    actual = set(holdings["Ticker"])
    cached = set(previous["Ticker"].astype(str).str.upper().str.strip())
    if actual != cached:
        return False, notes + [
            f"Local input contains {len(actual)} securities but cached analysis covers {len(cached)}.",
            "Connect and choose Sync All Data to rebuild analytics for the current holdings.",
        ]
    notes.append("Saved portfolio analytics match all current holdings.")

    # Keep the cache directory overridable for isolated tests that don't run inside
    # this repository; production always uses src.data's private .offline_cache.
    if base.resolve() == BASE.resolve():
        prices = cached_prices()
        status = price_cache_status()
    else:
        cache_file = base / ".offline_cache" / "adjusted_closes.csv"
        try:
            prices = pd.read_csv(cache_file, index_col=0) if cache_file.exists() else pd.DataFrame()
        except (OSError, ValueError, pd.errors.ParserError):
            prices = pd.DataFrame()
        status = {}
    missing = [ticker for ticker in actual if ticker not in prices or pd.to_numeric(prices[ticker],errors="coerce").dropna().empty]
    if missing:
        return False, notes + [
            f"Saved price history is missing for {len(missing)} holding(s).",
            "Connect and choose Sync All Data once to cache all stock prices.",
        ]
    notes.append(f"Offline price history covers all {count} holdings.")
    date = status.get("latest_observed_market_date")
    if date:
        notes.append(f"Last saved market date: {date}. Offline data will not update until the next sync.")
    notes.append("READY: You can disconnect and open the Portfolio Dashboard offline.")
    return True, notes


def main() -> int:
    ok, lines = check()
    for line in lines:
        print(line)
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
