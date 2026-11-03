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
from recover_cached_portfolio import recover as recover_from_local_prices

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

last_saved = price_cache_status()
cache = cached_prices()
has_summary = (LATEST / "summary.json").is_file()
has_holdings = (LATEST / "holdings_analysis.csv").is_file()

# A complete cached holdings/price set already permits a useful private dashboard.
# Do not require the first-run user to click an online rebuild just to see analytics.
# Only attempt this once per UI session; if the market history is unsuitable, show
# the reason and retain the raw holdings rather than getting stuck in a rerun loop.
cache_covers_holdings = (
    not cache.empty
    and all(t in cache.columns and not cache[t].dropna().empty for t in portfolio["Ticker"])
)
if not has_summary and cache_covers_holdings and not st.session_state.get("offline_recovery_attempted", False):
    st.session_state["offline_recovery_attempted"] = True
    try:
        with st.spinner("Restoring portfolio analytics from already saved prices (no internet)…"):
            result = recover_from_local_prices()
        if result.get("status") == "RECOVERED_OFFLINE":
            st.session_state["offline_recovery_result"] = result
            st.rerun()
    except Exception as exc:
        st.session_state["offline_recovery_error"] = f"{type(exc).__name__}: {exc}"

ready, notes = offline_readiness_check()
has_summary = (LATEST / "summary.json").is_file()
has_holdings = (LATEST / "holdings_analysis.csv").is_file()
analysis_mode = None
if has_summary:
    try:
        analysis_mode = json.loads((LATEST / "summary.json").read_text(encoding="utf-8")).get("analysis_mode")
    except (OSError, ValueError):
        pass

p1, p2, p3, p4 = st.columns(4)
p1.metric("Holdings loaded", len(portfolio))
p2.metric(
    "Saved analytics",
    "Cached recovery" if analysis_mode == "RECOVERY_CACHED_PRICES"
    else "Full" if has_summary and has_holdings else "Not generated",
)
p3.metric(
    "Cached price coverage",
    f"{sum(t in cache.columns and not cache[t].dropna().empty for t in portfolio['Ticker'])}/{len(portfolio)}",
)
p4.metric("Offline ready", "Yes" if ready else "Not yet")

if ready and analysis_mode == "RECOVERY_CACHED_PRICES":
    st.success("Core portfolio dashboard restored from locally cached prices. No network access was needed.")
    st.warning(
        "This is a recovery snapshot, not the full research pipeline. Portfolio weights, "
        "saved prices, historical model returns, volatility and concentration are available. "
        "Advanced optimization, full alpha models and realized returns are not inferred "
        "from missing inputs. Use the full online build button below to complete them."
    )
elif ready:
    st.success("Saved portfolio analytics are ready. Open Portfolio Dashboard from the sidebar.")
else:
    st.warning(
        "The holdings file is available, but the complete offline dashboard is not yet ready. "
        "Any existing local prices can still be viewed."
    )
    for note in notes:
        st.write("• " + note)

if st.session_state.get("offline_recovery_error"):
    st.error(
        "Automatic offline recovery could not finish: "
        + str(st.session_state["offline_recovery_error"])
    )
    st.caption("You can retry manually below or inspect any online build log.")
if not ready and cache_covers_holdings and st.button("Restore dashboard from cached prices (offline)"):
    try:
        recovered = recover_from_local_prices()
        st.session_state.pop("offline_recovery_error", None)
        st.success(f"Local recovery: {recovered.get('status')}.")
        st.rerun()
    except Exception as exc:
        st.error(f"Local recovery failed: {type(exc).__name__}: {exc}")

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
    "Use the complete online research run to update market history, benchmark/factor "
    "analysis, risk models, advanced portfolio optimization and weekly Random Forest "
    "forecasts. The cached-price recovery above requires no internet. The online run "
    "updates the saved dashboard only after exporting its results successfully."
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
        elif not all(
            (LATEST / filename).is_file()
            for filename in ("summary.json", "holdings_analysis.csv")
        ):
            st.error(
                "The research command finished without producing a complete "
                "outputs/latest snapshot. Open the log below to see the last completed step."
            )
        else:
            verified, diagnostics = offline_readiness_check()
            if verified:
                st.success("Analysis and offline price cache are ready. Opening the full dashboard.")
                st.rerun()
            else:
                st.warning(
                    "The analysis files were generated, but some offline prerequisites "
                    "are still missing. Check the diagnostics below; the saved results remain available."
                )
                for diagnostic in diagnostics:
                    st.write("• " + diagnostic)
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
