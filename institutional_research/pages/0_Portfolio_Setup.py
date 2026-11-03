from __future__ import annotations

# First-run and offline recovery screen for the private portfolio workstation.

from collections import deque
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import sys

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parents[1]
ROOT = BASE.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from src.portfolio import load_portfolio
from src.data import cached_prices, price_cache_status
from offline_portfolio_check import check as offline_readiness_check

LATEST = BASE / "outputs" / "latest"
LOG_PATH = BASE / ".offline_cache" / "portfolio_build.log"

st.set_page_config(page_title="Portfolio Setup & Holdings", layout="wide")
st.title("Portfolio Setup & Holdings")
st.caption(
    "Private local portfolio · safe first-run recovery · saved-data status. "
    "Opening this page never downloads market data."
)

portfolio_path = BASE / "portfolio.csv"
if not portfolio_path.exists():
    st.error(
        "The local portfolio.csv is missing. While connected, pull the main GitHub "
        "repository and reopen the Research Hub."
    )
    st.code(
        f"Set-Location -LiteralPath '{ROOT}'\ngit pull --ff-only origin main",
        language="powershell",
    )
    st.stop()

try:
    portfolio = load_portfolio(portfolio_path)
except Exception as exc:
    st.error(f"Portfolio input cannot be read: {type(exc).__name__}: {exc}")
    st.stop()

ready, notes = offline_readiness_check()
last_saved = price_cache_status()
cache = cached_prices()
has_summary = (LATEST / "summary.json").is_file()
has_holdings = (LATEST / "holdings_analysis.csv").is_file()

p1, p2, p3, p4 = st.columns(4)
p1.metric("Holdings loaded", len(portfolio))
p2.metric("Saved analytics", "Present" if has_summary and has_holdings else "Not generated")
p3.metric(
    "Cached price coverage",
    f"{sum(t in cache.columns and not cache[t].dropna().empty for t in portfolio['Ticker'])}/{len(portfolio)}",
)
p4.metric("Offline ready", "Yes" if ready else "Not yet")

if ready:
    st.success("Offline portfolio is ready. Open Portfolio Dashboard from the sidebar.")
else:
    st.warning(
        "Your holdings file is available, but the computer does not yet have everything "
        "needed for the complete offline dashboard. It is safe to view the holdings below."
    )
    for note in notes:
        st.write("• " + note)

if last_saved.get("latest_observed_market_date"):
    st.caption(
        "Last locally saved market close: "
        + str(last_saved["latest_observed_market_date"])
        + ". Prices stay at this date until the next successful online update."
    )

st.subheader("Your current portfolio")
st.caption("These holdings and cost-basis amounts stay inside this local workstation.")
columns = [c for c in ("Ticker", "Shares", "AverageCost", "ManualWeight", "Notes") if c in portfolio.columns]
st.dataframe(portfolio[columns], use_container_width=True, hide_index=True)

if not cache.empty:
    latest_rows = []
    for ticker in portfolio["Ticker"]:
        series = (
            pd.to_numeric(cache[ticker], errors="coerce").dropna()
            if ticker in cache else pd.Series(dtype=float)
        )
        latest_rows.append({
            "Ticker": ticker,
            "Last saved close": float(series.iloc[-1]) if not series.empty else None,
            "Market date": str(series.index[-1].date()) if not series.empty else None,
        })
    st.subheader("Saved market prices")
    st.dataframe(pd.DataFrame(latest_rows), use_container_width=True, hide_index=True)

st.divider()
st.subheader("Build or repair the full portfolio dashboard")
st.write(
    "Run this only when you have internet access. The full analysis downloads market "
    "history, calculates risk/performance/valuation models, and builds the weekly "
    "Random Forest forecasts. It saves results for later offline use. A failed "
    "rebuild preserves the last good saved analysis."
)

if st.button("Build / refresh portfolio analysis (online)", type="primary", use_container_width=False):
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    environment = os.environ.copy()
    # The dashboard itself may have been opened offline; the *explicit* build
    # action must be allowed to connect to market-data providers.
    environment["ANTZAZ_OFFLINE"] = "0"
    cmd = [sys.executable, "run_research.py"]
    st.info("Running portfolio research. Results and errors are saved to a private local log.")
    try:
        with st.spinner("Building analytics and saving offline market data…"):
            with LOG_PATH.open("w", encoding="utf-8", errors="replace") as log:
                proc = subprocess.run(
                    cmd,
                    cwd=BASE,
                    env=environment,
                    stdout=log,
                    stderr=subprocess.STDOUT,
                    check=False,
                    timeout=1800,
                )
        if proc.returncode != 0:
            st.error(
                f"Portfolio analysis exited with code {proc.returncode}. "
                "The last working analysis, if any, was preserved."
            )
        elif not (LATEST / "summary.json").is_file():
            st.error("The research command ended without producing outputs/latest/summary.json.")
        else:
            st.success("Analysis files generated. Refreshing the dashboard navigation.")
            st.rerun()
    except subprocess.TimeoutExpired:
        st.error("The build exceeded the 30-minute safety limit. Check the local log for its last step.")
    except Exception as exc:
        st.error(f"Could not start the portfolio analysis: {type(exc).__name__}: {exc}")

if LOG_PATH.exists():
    with st.expander("Last analysis log / troubleshooting", expanded=False):
        st.caption(f"Local log: {LOG_PATH}")
        try:
            with LOG_PATH.open("r", encoding="utf-8", errors="replace") as handle:
                last_lines = list(deque(handle, maxlen=55))
            st.code("".join(last_lines) or "No log output.", language="text")
        except OSError as exc:
            st.write(f"Cannot read the last log: {exc}")

st.caption(
    "You can also prepare everything from PowerShell: "
    f"Set-Location -LiteralPath '{ROOT}'; "
    "powershell -NoProfile -ExecutionPolicy Bypass -File .\\prepare_offline_portfolio.ps1"
)
