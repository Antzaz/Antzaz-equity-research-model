from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "outputs" / "latest"

st.set_page_config(page_title="Alpha Analysis", layout="wide")
st.title("Alpha & Factor-Adjusted Performance")

if not OUT.exists():
    st.error("No analysis outputs found. Run: python run_research.py")
    st.stop()


def read(name: str) -> pd.DataFrame:
    path = OUT / f"{name}.csv"
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def pct(value, digits=1):
    return "—" if value is None or pd.isna(value) else f"{value:.{digits}%}"


def num(value, digits=2):
    return "—" if value is None or pd.isna(value) else f"{value:.{digits}f}"


summary_path = OUT / "summary.json"
summary_json = {}
if summary_path.exists():
    with open(summary_path, "r", encoding="utf-8") as f:
        summary_json = json.load(f)
portfolio = summary_json.get("portfolio", {})
alpha_meta = summary_json.get("alpha_analysis", {})

alpha = read("alpha_summary")
loadings = read("alpha_factor_loadings")
rolling = read("rolling_alpha")
decomp = read("alpha_return_decomposition")
relative = read("benchmark_relative")

st.caption(
    alpha_meta.get(
        "method_note",
        "Alpha is a research diagnostic. If historical portfolio weights are not supplied, the project applies current weights to historical returns.",
    )
)

if alpha.empty:
    st.warning("No alpha regressions were produced. Check price history and the terminal output for factor-download warnings.")
    st.stop()


def model_value(model: str, col: str):
    x = alpha.loc[alpha["Model"] == model, col] if col in alpha.columns else pd.Series(dtype=float)
    return x.iloc[0] if not x.empty else None


c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Raw active return", pct(portfolio.get("active_annualized_return")))
c2.metric("CAPM alpha", pct(model_value("CAPM - Benchmark", "AnnualizedAlpha")))
c3.metric("CAPM alpha t-stat", num(model_value("CAPM - Benchmark", "AlphaTStat")))
c4.metric("Carhart 4 alpha", pct(model_value("Carhart 4", "AnnualizedAlpha")))
c5.metric("FF5 alpha", pct(model_value("Fama-French 5", "AnnualizedAlpha")))

if "AnnualizedAlpha" in alpha.columns:
    fig = px.bar(
        alpha,
        x="Model",
        y="AnnualizedAlpha",
        color="Interpretation" if "Interpretation" in alpha.columns else None,
        title="Residual alpha by risk model",
    )
    fig.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig, use_container_width=True)

show_cols = [
    "Model", "AnnualizedAlpha", "AlphaTStat", "AlphaPValue", "R2",
    "ResidualVolatilityAnnualized", "Observations", "StartDate", "EndDate",
    "Significant5Pct", "Interpretation", "RawActiveAnnualizedReturn",
]
st.subheader("Regression summary")
st.dataframe(alpha[[c for c in show_cols if c in alpha.columns]], use_container_width=True)

left, right = st.columns(2)
with left:
    if not rolling.empty:
        rolling["Date"] = pd.to_datetime(rolling["Date"], errors="coerce")
        choices = rolling["Window"].dropna().unique().tolist()
        selected_window = st.selectbox("Rolling alpha window", choices, index=0)
        view = rolling[rolling["Window"] == selected_window]
        fig = px.line(
            view,
            x="Date",
            y="AnnualizedAlpha",
            color="Model",
            title=f"Rolling {selected_window} alpha",
        )
        fig.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Rolling regressions are sampled weekly to keep the output compact.")

with right:
    if not rolling.empty and "AlphaTStat" in rolling.columns:
        view = rolling[rolling["Window"] == selected_window]
        fig = px.line(
            view,
            x="Date",
            y="AlphaTStat",
            color="Model",
            title=f"Rolling {selected_window} alpha t-statistic",
        )
        fig.add_hline(y=1.96, line_dash="dot")
        fig.add_hline(y=-1.96, line_dash="dot")
        st.plotly_chart(fig, use_container_width=True)

st.subheader("Factor exposures and return decomposition")
model_choices = alpha["Model"].dropna().tolist()
selected_model = st.selectbox("Regression model", model_choices, index=0)

c1, c2 = st.columns(2)
with c1:
    model_loadings = loadings[loadings["Model"] == selected_model] if not loadings.empty else pd.DataFrame()
    if not model_loadings.empty:
        fig = px.bar(
            model_loadings,
            x="Factor",
            y="Beta",
            title=f"{selected_model}: factor loadings",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(model_loadings, use_container_width=True)
with c2:
    model_decomp = decomp[
        (decomp["Model"] == selected_model)
        & (~decomp["Type"].isin(["Factor explained total", "Residual alpha summary"]))
    ] if not decomp.empty else pd.DataFrame()
    if not model_decomp.empty:
        fig = px.bar(
            model_decomp,
            x="Component",
            y="AnnualizedContribution",
            color="Type",
            title=f"{selected_model}: arithmetic annualized return decomposition",
        )
        fig.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig, use_container_width=True)
        st.caption(
            "The decomposition uses annualized arithmetic mean factor contributions, so it is intended to explain the regression, not exactly reconcile geometric CAGR."
        )

warnings = alpha_meta.get("french_factor_warnings") or []
if warnings:
    st.warning("Some Kenneth French factor downloads were unavailable on this run:\n\n" + "\n".join(f"- {x}" for x in warnings))

st.divider()
st.markdown("### How to read this")
st.write(
    "Raw active return tells you whether the portfolio beat the benchmark. CAPM alpha removes broad-market beta. "
    "Fama–French 3 additionally controls for size and value. Carhart 4 adds momentum. Fama–French 5 adds profitability and investment. "
    "The Public Style Proxy model additionally tests whether public ETF exposures to growth, momentum, quality and low volatility explain the return."
)
st.write(
    "A positive alpha with a weak t-statistic is not strong evidence of persistent skill. A positive alpha that remains positive across several models, "
    "is stable through rolling windows, and has a materially positive t-statistic is more convincing—but still does not prove future alpha."
)
st.info(
    "Current limitation: unless you provide point-in-time portfolio weights or transaction history, this project backcasts today's weights through history. "
    "That is useful for studying the current portfolio's exposures, but it is not the same as realized manager performance."
)
