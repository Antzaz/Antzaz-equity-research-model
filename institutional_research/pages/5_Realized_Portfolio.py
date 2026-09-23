from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE=Path(__file__).resolve().parents[1]
OUT=BASE/"outputs"/"latest"

st.set_page_config(page_title="Realized Portfolio",layout="wide")
st.title("Realized Portfolio & Point-in-Time Attribution")
st.caption("Offline/private view built from the transaction ledger. It does not infer trades from current holdings.")


def read(name):
    path=OUT/f"{name}.csv"
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def pct(v):
    try:return f"{float(v):.1%}"
    except:return "—"


if not (OUT/"summary.json").exists():
    st.error("Run python run_research.py first.")
    st.stop()

summary=json.loads((OUT/"summary.json").read_text(encoding="utf-8"))
realized=summary.get("realized_performance",{})
status=realized.get("status","NO_LEDGER")

if status=="NO_LEDGER":
    st.info("No transaction_ledger.csv is present. Copy transaction_ledger_template.csv to transaction_ledger.csv and enter actual transactions/external cash flows.")
    st.stop()

if status!="PASS":
    st.warning(f"Realized-performance status: {status}. Review the warnings and ledger completeness before treating TWR/MWR as final.")

cols=st.columns(6)
cols[0].metric("TWR",pct(realized.get("twr_total")))
cols[1].metric("Annualized TWR",pct(realized.get("twr_annualized")))
cols[2].metric("Money-weighted IRR",pct(realized.get("mwr_xirr")))
cols[3].metric("Benchmark",pct(realized.get("benchmark_total_return")))
cols[4].metric("Active return",pct(realized.get("active_total_return")))
cols[5].metric("Turnover / ending NAV",pct(realized.get("turnover_vs_ending_nav")))

warnings=summary.get("decision_loop",{}).get("realized_warnings",[])
for warning in warnings:
    st.warning(warning)

ts=read("realized_portfolio_timeseries")
if not ts.empty:
    ts["Date"]=pd.to_datetime(ts["Date"],errors="coerce")
    cols=[c for c in ["PortfolioGrowth","BenchmarkGrowth"] if c in ts]
    if cols:
        long=ts.melt(id_vars="Date",value_vars=cols,var_name="Series",value_name="Growth")
        st.plotly_chart(px.line(long,x="Date",y="Growth",color="Series",title="Realized growth path"),use_container_width=True)
    st.subheader("NAV and external cash flows")
    st.dataframe(ts.tail(40),use_container_width=True,hide_index=True)

attr=read("realized_security_attribution")
if not attr.empty:
    st.subheader("Realized security contribution")
    fig=px.bar(attr,x="Ticker",y="TotalContribution",title="Arithmetic contribution from point-in-time weights")
    st.plotly_chart(fig,use_container_width=True)
    st.dataframe(attr,use_container_width=True,hide_index=True)

weights=read("point_in_time_weights")
if not weights.empty:
    st.subheader("Point-in-time portfolio weights")
    weights["Date"]=pd.to_datetime(weights["Date"],errors="coerce")
    latest=weights["Date"].max()
    st.dataframe(weights[weights["Date"]==latest].sort_values("Weight",ascending=False),use_container_width=True,hide_index=True)

brinson=read("brinson_sector_attribution")
st.subheader("Allocation / selection / interaction")
if brinson.empty:
    st.info("Brinson attribution remains gated until benchmark_sector_history.csv contains point-in-time sector weights and returns. Current benchmark sector weights are never backfilled as historical truth.")
else:
    totals=brinson.groupby("Sector",as_index=False)[["Allocation","Selection","Interaction","TotalActiveContribution"]].sum()
    st.dataframe(totals,use_container_width=True,hide_index=True)

st.caption("Accounting note: TWR removes explicit DEPOSIT/WITHDRAW flows. MWR uses dated external flows plus ending NAV. Raw closes and corporate actions are preferred for ledger accounting.")
