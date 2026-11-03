from __future__ import annotations
"""Private stock monitor: local portfolio cost basis never enters the public showcase."""
from pathlib import Path
import pandas as pd
import streamlit as st
from src.portfolio import load_portfolio
from src.data import cached_prices, download_prices, offline_mode, price_cache_status

BASE = Path(__file__).resolve().parents[1]
st.title("Stock Prices & Purchase Costs")
st.caption("All holdings · latest available market close · cost basis remains private")
path = BASE / "portfolio.csv"
if not path.exists():
    st.warning("Private portfolio.csv is missing. Restore your local portfolio file first.")
    st.stop()
holdings = load_portfolio(path)
tickers = holdings["Ticker"].tolist()
@st.cache_data(ttl=900, show_spinner=False)
def prices_for(names, use_offline):
    # The Research Hub launches the dashboard with ANTZAZ_OFFLINE=1. In that mode
    # the shared market-data adapter never makes network requests.
    return download_prices(list(names), period="2y")

cache_status = price_cache_status()
saved_date = cache_status.get("latest_observed_market_date")
saved_at = cache_status.get("last_online_refresh_utc")
if offline_mode():
    st.info(
        "Offline mode — showing locally saved adjusted closes. "
        + (f"Latest cached market date: {saved_date}." if saved_date else "No cached quotes yet.")
        + (f" Last online sync: {saved_at}." if saved_at else "")
    )
else:
    st.caption("Online quotes when available; locally saved adjusted closes are the fallback.")

try:
    with st.spinner("Loading locally saved prices…" if offline_mode() else "Loading market prices…"):
        closes = prices_for(tuple(tickers), offline_mode())
except Exception as exc:
    st.warning(f"Market prices unavailable: {exc}. Showing share costs without inventing quotes.")
    closes = cached_prices()

rows = []
for _, h in holdings.iterrows():
    ticker = h["Ticker"]
    series = pd.to_numeric(closes[ticker], errors="coerce").dropna() if ticker in closes.columns else pd.Series(dtype=float)
    current = float(series.iloc[-1]) if not series.empty else None
    cost = float(h["AverageCost"]) if pd.notna(h["AverageCost"]) and h["AverageCost"] > 0 else None
    row = {"Ticker": ticker, "Latest close": current, "Bought at (avg)": cost,
           "Since purchase %": (current/cost-1)*100 if current is not None and cost else None,
           "Price date": str(series.index[-1].date()) if not series.empty else None}
    for label, n in [("1D",1),("1W",5),("1M",21),("3M",63),("6M",126),("1Y",252)]:
        row[label+" %"] = (float(series.iloc[-1]/series.iloc[-1-n]-1)*100) if len(series)>n else None
    rows.append(row)
frame = pd.DataFrame(rows)
st.dataframe(frame, use_container_width=True, hide_index=True)
selected = st.selectbox("Price history", tickers)
if selected in closes.columns:
    chart = closes[selected].dropna()
    if not chart.empty:
        st.line_chart(chart.rename("Adjusted close"))
st.caption("Bought at uses AverageCost from your private portfolio.csv; it is not inferred from the inception date. Offline values reflect the last successful online sync and never claim to be current intraday quotes. Adjusted-price returns include provider corporate-action adjustments, but may not represent cash dividends received or your transaction costs. Currency effects can differ.")
