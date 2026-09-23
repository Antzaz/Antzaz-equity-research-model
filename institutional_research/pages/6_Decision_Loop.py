from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE=Path(__file__).resolve().parents[1]
OUT=BASE/"outputs"/"latest"

st.set_page_config(page_title="Research Decision Loop",layout="wide")
st.title("Research → Sizing → Outcome → Learning")
st.caption("Offline/private decision surface connecting local equity-research workbooks with portfolio construction and later realized outcomes.")


def read(name):
    path=OUT/f"{name}.csv"
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


if not (OUT/"summary.json").exists():
    st.error("Run python run_research.py first.")
    st.stop()

summary=json.loads((OUT/"summary.json").read_text(encoding="utf-8"))
meta=summary.get("decision_loop",{})

c1,c2,c3=st.columns(3)
c1.metric("Research workbooks linked",meta.get("research_bridge_coverage",0))
c2.metric("Transaction ledger rows",meta.get("transaction_ledger_rows",0))
c3.metric("Decision outcomes tracked",meta.get("decision_journal_rows",0))

bridge=read("research_expected_return_bridge")
st.subheader("Research-model expected-return bridge")
if bridge.empty:
    st.info("Build current local company workbooks first. The portfolio layer will discover the latest workbook for each holding and use confidence-shrunk expected returns.")
else:
    show=[c for c in ["Ticker","CurrentPrice","BearValue","BaseValue","BullValue","RawExpectedReturn","Confidence","ConfidenceAdjustedExpectedReturn","ExpectedAlpha","ModelView","SourceStatus"] if c in bridge]
    st.dataframe(bridge[show],use_container_width=True,hide_index=True)

sizing=read("position_sizing_ranges")
st.subheader("Transparent position-sizing ranges")
if sizing.empty:
    st.info("Sizing ranges require research expected-return coverage.")
else:
    st.dataframe(sizing,use_container_width=True,hide_index=True)
    plot=sizing.melt(id_vars=["Ticker"],value_vars=[c for c in ["CurrentWeight","SuggestedMin","SuggestedMidpoint","SuggestedMax"] if c in sizing],var_name="WeightType",value_name="Weight")
    st.plotly_chart(px.bar(plot,x="Ticker",y="Weight",color="WeightType",barmode="group"),use_container_width=True)

budget=read("thesis_budget")
st.subheader("Capital vs risk vs expected-alpha budget")
if not budget.empty:
    cols=[c for c in ["Ticker","Weight","RiskWeight","ExpectedAlphaBudget"] if c in budget]
    long=budget[cols].melt(id_vars="Ticker",var_name="Budget",value_name="Share")
    st.plotly_chart(px.bar(long,x="Ticker",y="Share",color="Budget",barmode="group"),use_container_width=True)
    st.dataframe(budget,use_container_width=True,hide_index=True)
else:
    st.info("Expected-alpha budget appears once research workbooks provide usable expected returns.")

scenarios=read("fundamental_research_scenarios")
custom=read("custom_thesis_scenarios")
st.subheader("Bottom-up research scenarios")
if scenarios.empty:
    st.info("No company-model bear/base/bull (or P10/P90) valuation coverage is available yet.")
else:
    portfolio_rows=scenarios[scenarios["Ticker"].isna()] if "Ticker" in scenarios else scenarios
    st.dataframe(portfolio_rows,use_container_width=True,hide_index=True)
    st.dataframe(scenarios,use_container_width=True,hide_index=True)

st.subheader("Explicit thesis scenarios")
if custom.empty:
    st.info("Copy fundamental_scenarios_template.csv to fundamental_scenarios.csv to add explicit per-company shocks such as AI capex boom, recession, rates or USD scenarios. Missing shocks are not inferred.")
else:
    st.dataframe(custom,use_container_width=True,hide_index=True)

decomp=read("expected_return_decomposition_summary")
st.subheader("Expected-return decomposition")
if not decomp.empty:
    st.plotly_chart(px.bar(decomp,x="Component",y="PortfolioContribution"),use_container_width=True)
    st.dataframe(decomp,use_container_width=True,hide_index=True)

rebalance=read("rebalance_cost_analysis")
st.subheader("Transaction-cost-aware rebalance gate")
if rebalance.empty:
    st.info("No expected-return-aware optimizer target is available.")
else:
    st.dataframe(rebalance,use_container_width=True,hide_index=True)

learning=read("decision_learning")
outcomes=read("decision_journal_outcomes")
st.subheader("Decision-journal outcome analytics")
if learning.empty:
    st.info("Decision learning appears after journal entries mature through 3M/6M/12M horizons.")
else:
    st.dataframe(learning,use_container_width=True,hide_index=True)
if not outcomes.empty:
    with st.expander("Decision-level outcomes"):
        st.dataframe(outcomes,use_container_width=True,hide_index=True)

st.caption(meta.get("method_note",""))
