from __future__ import annotations

"""Interactive classical + ML portfolio optimization dashboard."""

import json
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "outputs" / "latest"

st.set_page_config(page_title="Portfolio Optimization", layout="wide")
st.title("Portfolio Optimization — Classical + Machine Learning")
st.caption(
    "Maximum-history portfolio construction · Minimum Variance · Maximum Sharpe · "
    "Mean-Variance · ML expected returns · regime-aware risk"
)


def read(name: str) -> pd.DataFrame:
    path = OUT / f"{name}.csv"
    try:
        return pd.read_csv(path) if path.exists() else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def pct(x, digits=1):
    try:
        return f"{float(x):.{digits}%}" if pd.notna(x) else "—"
    except Exception:
        return "—"


def num(x, digits=2):
    try:
        return f"{float(x):.{digits}f}" if pd.notna(x) else "—"
    except Exception:
        return "—"


if not OUT.exists() or not (OUT / "summary.json").exists():
    st.error("No portfolio outputs found. Run `python run_research.py` from institutional_research first.")
    st.stop()

summary_all = json.loads((OUT / "summary.json").read_text(encoding="utf-8"))
meta = summary_all.get("portfolio_optimization") or {}
summary = read("optimizer_summary")
weights = read("optimizer_weights")
frontier = read("efficient_frontier")
expected = read("optimizer_expected_returns")
corr = read("optimizer_correlation")
coverage = read("optimizer_data_coverage")

if summary.empty:
    st.warning(
        "The maximum-data optimizer has no output yet. Pull the latest code and run "
        "`python run_research.py`, then reload this page."
    )
    st.stop()

regime = (meta.get("regime") or {}).get("regime") or "Unavailable"
regime_conf = (meta.get("regime") or {}).get("confidence") or "—"
cov_meta = meta.get("covariance") or {}
exp_meta = meta.get("expected_returns") or {}

cards = st.columns(6)
cards[0].metric("Assets optimized", int(meta.get("history_assets") or 0))
cards[1].metric("History target", f"{int(meta.get('max_history_years') or 0)}Y")
cards[2].metric("ML coverage", f"{int(exp_meta.get('ml_covered_assets') or 0)}/{int(exp_meta.get('assets') or 0)}")
cards[3].metric("Current regime", regime)
cards[4].metric("Regime confidence", regime_conf)
cards[5].metric("Common overlap", f"{int(cov_meta.get('complete_overlap_rows') or 0):,} days")

st.info(
    "ML does not directly choose weights. It supplies confidence-shrunk 12-month expected-return "
    "inputs and a bounded regime adjustment to the risk matrix; the final portfolio remains a "
    "transparent constrained optimizer."
)

st.subheader("Efficient frontier")
left, right = st.columns([1.65, 1])

with left:
    fig = go.Figure()
    if not frontier.empty:
        for name, group in frontier.groupby("Frontier"):
            group = group.sort_values("ExpectedVolatility")
            fig.add_trace(go.Scatter(
                x=group["ExpectedVolatility"],
                y=group["ExpectedReturn"],
                mode="lines+markers",
                name=name,
                marker={"size": 5},
            ))
    fig.add_trace(go.Scatter(
        x=summary["ExpectedVolatility"],
        y=summary["ExpectedReturn"],
        mode="markers+text",
        text=summary["Portfolio"],
        textposition="top center",
        name="Portfolios",
        marker={"size": 13, "symbol": "diamond"},
        hovertemplate="<b>%{text}</b><br>Return %{y:.1%}<br>Vol %{x:.1%}<extra></extra>",
    ))
    fig.update_layout(
        xaxis_title="Expected annual volatility",
        yaxis_title="Expected annual return",
        xaxis_tickformat=".0%",
        yaxis_tickformat=".0%",
        legend_title=None,
        height=560,
        margin={"l": 20, "r": 20, "t": 20, "b": 20},
    )
    st.plotly_chart(fig, use_container_width=True)

with right:
    display = summary.copy()
    for c in ["ExpectedReturn", "ExpectedVolatility", "Turnover", "MaxWeight"]:
        display[c] = pd.to_numeric(display[c], errors="coerce").map(lambda x: pct(x) if pd.notna(x) else "—")
    display["ExpectedSharpe"] = pd.to_numeric(display["ExpectedSharpe"], errors="coerce").map(
        lambda x: num(x) if pd.notna(x) else "—"
    )
    display["EffectiveHoldings"] = pd.to_numeric(display["EffectiveHoldings"], errors="coerce").map(
        lambda x: num(x, 1) if pd.notna(x) else "—"
    )
    st.dataframe(display, use_container_width=True, hide_index=True, height=500)

choices = [x for x in summary["Portfolio"].tolist() if x not in {"Current Portfolio"}]
selected = st.selectbox(
    "Compare target allocation",
    choices,
    index=choices.index("Regime-Aware ML Maximum Sharpe")
    if "Regime-Aware ML Maximum Sharpe" in choices else 0,
)

selected_summary = summary[summary["Portfolio"] == selected]
if not selected_summary.empty:
    row = selected_summary.iloc[0]
    cols = st.columns(5)
    cols[0].metric("Expected return", pct(row["ExpectedReturn"]))
    cols[1].metric("Expected volatility", pct(row["ExpectedVolatility"]))
    cols[2].metric("Expected Sharpe", num(row["ExpectedSharpe"]))
    cols[3].metric("Turnover", pct(row["Turnover"]))
    cols[4].metric("Effective holdings", num(row["EffectiveHoldings"], 1))

selected_weights = weights[weights["Portfolio"] == selected].copy()
if not selected_weights.empty:
    current = selected_weights[["Ticker", "Sector", "CurrentWeight"]].rename(columns={"CurrentWeight": "Weight"})
    current["Portfolio"] = "Current"
    proposed = selected_weights[["Ticker", "Sector", "TargetWeight"]].rename(columns={"TargetWeight": "Weight"})
    proposed["Portfolio"] = selected
    alloc = pd.concat([current, proposed], ignore_index=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            alloc,
            x="Ticker",
            y="Weight",
            color="Portfolio",
            barmode="group",
            title="Current vs target weights",
        )
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(height=430, legend_title=None)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        delta = selected_weights.sort_values("WeightChange")
        fig = px.bar(
            delta,
            x="WeightChange",
            y="Ticker",
            orientation="h",
            title="Required weight changes",
            hover_data=["CurrentWeight", "TargetWeight"],
        )
        fig.update_xaxes(tickformat="+.0%")
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)

    sector_alloc = (
        alloc.groupby(["Portfolio", "Sector"], as_index=False)["Weight"].sum()
        .sort_values(["Portfolio", "Weight"], ascending=[True, False])
    )
    fig = px.bar(
        sector_alloc,
        x="Sector",
        y="Weight",
        color="Portfolio",
        barmode="group",
        title="Sector exposure — current vs selected optimizer",
    )
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

st.subheader("Expected-return inputs")
if not expected.empty:
    cols = [
        "Ticker", "HistoricalPrior", "ManualExpectedReturn",
        "ClassicalExpectedReturn", "MLTotalReturn", "BlendedMLExpectedReturn",
    ]
    melted = expected[[c for c in cols if c in expected.columns]].melt(
        "Ticker", var_name="Input", value_name="ExpectedReturn"
    )
    melted["Input"] = melted["Input"].replace({
        "HistoricalPrior": "Historical prior",
        "ManualExpectedReturn": "Manual",
        "ClassicalExpectedReturn": "Classical",
        "MLTotalReturn": "Raw ML total",
        "BlendedMLExpectedReturn": "Confidence-shrunk ML",
    })
    fig = px.bar(
        melted,
        x="Ticker",
        y="ExpectedReturn",
        color="Input",
        barmode="group",
        title="Classical, manual and ML expected-return inputs",
    )
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    ml_table = expected[[
        c for c in [
            "Ticker", "MLExcessReturn12M", "MLConfidence", "MLConfidenceScalar",
            "MLBlendWeight", "BlendedMLExpectedReturn", "MLSource"
        ] if c in expected.columns
    ]].copy()
    st.dataframe(
        ml_table.style.format({
            "MLExcessReturn12M": "{:.1%}",
            "MLConfidenceScalar": "{:.0%}",
            "MLBlendWeight": "{:.0%}",
            "BlendedMLExpectedReturn": "{:.1%}",
        }, na_rep="—"),
        use_container_width=True,
        hide_index=True,
    )

st.subheader("Diversification and data depth")
left, right = st.columns([1.4, 1])

with left:
    if not corr.empty and "Ticker" in corr:
        matrix = corr.set_index("Ticker")
        matrix = matrix.apply(pd.to_numeric, errors="coerce")
        fig = px.imshow(
            matrix,
            text_auto=".2f",
            zmin=-1,
            zmax=1,
            aspect="auto",
            title="Maximum-history return correlation",
        )
        fig.update_layout(height=560)
        st.plotly_chart(fig, use_container_width=True)

with right:
    if not coverage.empty:
        cov_display = coverage.copy()
        if "ApproxYears" in cov_display:
            cov_display["ApproxYears"] = pd.to_numeric(
                cov_display["ApproxYears"], errors="coerce"
            ).map(lambda x: f"{x:.1f}" if pd.notna(x) else "—")
        st.caption("History actually used by security")
        st.dataframe(cov_display, use_container_width=True, hide_index=True, height=500)

with st.expander("Methodology and controls"):
    st.markdown(
        """
**Risk estimation**
- Uses up to 20 years per holding from the local ML SQLite history store when available.
- Falls back to maximum downloaded public price history.
- Builds a maximum-history pairwise covariance matrix, shrinks it toward the diagonal, forces
  positive semi-definiteness, and blends in Ledoit–Wolf covariance from the common overlap.

**Expected returns**
- Historical return prior uses multiple horizons rather than one full-sample mean.
- Manual expected returns remain optional.
- The existing walk-forward **Expected 12M Excess Return** ML model is read from the local
  prediction journal or trained from the point-in-time feature database when possible.
- ML is converted from benchmark-relative excess return to total return and then shrunk by
  confidence and freshness. Its maximum blend weight is capped.

**Regime-aware risk**
- The latest stored Market Regime Classifier output selects historically similar market states.
- Conditional covariance is blended into the long-run covariance according to model confidence.
- The long-run risk matrix always remains the majority weight.

**Optimization**
- Long-only, fully invested.
- Position, sector and turnover constraints come from `config.json`.
- Outputs Minimum Variance, Maximum Sharpe, Mean-Variance, ML Maximum Sharpe,
  ML Mean-Variance and Regime-Aware ML Maximum Sharpe.
- These are research allocations; no trades are executed.
        """
    )
    st.json(meta)
