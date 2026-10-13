from __future__ import annotations
"""Private stock monitor: local portfolio cost basis never enters the public showcase."""
from pathlib import Path
import pandas as pd
import streamlit as st
from src.portfolio import load_portfolio

BASE = Path(__file__).resolve().parents[1]
st.title("Stock Prices & Purchase Costs")
st.caption("All holdings · latest available market close · cost basis remains private")
path = BASE / "portfolio.csv"
if not path.exists():
    st.warning("Private portfolio.csv is missing. Restore your local portfolio file first.")
    st.stop()
holdings = load_portfolio(path)
tickers = holdings["Ticker"].tolist()
try:
    import yfinance as yf
    @st.cache_data(ttl=900, show_spinner=False)
    def prices_for(names):
        raw = yf.download(list(names), period="2y", auto_adjust=True, progress=False, threads=True)
        if isinstance(raw.columns, pd.MultiIndex):
            return raw["Close"]
        return raw[["Close"]].rename(columns={"Close":names[0]})
    with st.spinner("Loading market prices…"):
        closes = prices_for(tuple(tickers))
except Exception as exc:
    st.warning(f"Live quotes unavailable ({exc}). Showing any available saved analytics below.")
    closes = pd.DataFrame()

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
st.caption("Bought at uses AverageCost from your private portfolio.csv; it is not inferred from the inception date. Historical adjusted closes are not intraday quotes. Returns exclude dividends paid as cash and transaction costs; currency effects may differ.")
