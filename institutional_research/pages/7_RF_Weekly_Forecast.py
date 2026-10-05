from __future__ import annotations

from pathlib import Path
import json
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

BASE=Path(__file__).resolve().parents[1]
OUT=BASE/"outputs"
CSV=OUT/"random_forest_next_week.csv"
META=OUT/"random_forest_next_week.json"

st.title("🌲 Next-Week Random Forest")
st.caption("Five-business-day ML price-path research for current portfolio holdings. This is not a target price, valuation output or trading instruction.")

if not CSV.exists():
    st.info("No forecast file yet. Open Portfolio Dashboard from Research Hub to run a fresh portfolio analysis.")
    st.stop()

df=pd.read_csv(CSV)
if df.empty:
    st.warning("The latest Random Forest run produced no usable forecasts.")
    st.stop()

meta={}
if META.exists():
    try: meta=json.loads(META.read_text(encoding="utf-8"))
    except Exception: meta={}

tickers=sorted(df["Ticker"].dropna().astype(str).unique())
selected=st.multiselect("Holdings shown",tickers,default=tickers)
view=df[df["Ticker"].isin(selected)].copy()
view["date"]=pd.to_datetime(view["date"],errors="coerce")

fig=go.Figure()
for ticker,part in view.groupby("Ticker"):
    part=part.sort_values("date")
    last=(meta.get("models",{}).get(ticker,{}) or {}).get("last_price")
    x=list(part["date"]); y=list(part["predicted_price"])
    if last is not None and len(x):
        x=[x[0]-pd.tseries.offsets.BDay(1)]+x
        y=[float(last)]+y
    fig.add_trace(go.Scatter(x=x,y=y,mode="lines+markers",name=ticker))
fig.update_layout(title="Random Forest — next five business days",xaxis_title="Forecast date",yaxis_title="Price (listing currency)",hovermode="x unified")
st.plotly_chart(fig,use_container_width=True)

latest=[]
for ticker,part in view.groupby("Ticker"):
    part=part.sort_values("date")
    m=(meta.get("models",{}).get(ticker,{}) or {})
    last=m.get("last_price")
    end=part.iloc[-1]["predicted_price"] if len(part) else None
    latest.append({"Ticker":ticker,"Last price":last,"5D forecast":end,"Implied 5D change":None if last in (None,0) or end is None else end/last-1,"Training rows":m.get("training_rows"),"Status":m.get("status")})
summary=pd.DataFrame(latest)
st.dataframe(summary.style.format({"Last price":"{:,.2f}","5D forecast":"{:,.2f}","Implied 5D change":"{:+.2%}"}),use_container_width=True,hide_index=True)

with st.expander("Methodology & limitations",expanded=True):
    st.markdown("- One Random Forest model is trained separately for each current holding.\n- Inputs are lagged price returns, realized volatility and moving-average/range features; the model predicts the next daily return recursively for five business days.\n- The path is regenerated whenever portfolio research refresh runs, so it incorporates the newest downloaded prices.\n- Tree dispersion is stored as a model diagnostic, but it is not a calibrated probability interval.\n- One-week equity prices are extremely noisy. Treat this as a model-learning/monitoring experiment, not as intrinsic value or a buy/sell signal.")
