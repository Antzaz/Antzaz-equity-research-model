from __future__ import annotations

"""Resume-safe Streamlit showcase.

This app intentionally contains only synthetic/sanitized demonstration data. It does not
load the private portfolio bundle, GitHub Actions artifacts, portfolio.csv, cost basis,
position values, or downloadable private workbooks.
"""

import math
from datetime import date

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Investment Research & Portfolio Analytics", layout="wide")

st.title("Investment Research & Portfolio Analytics")
st.caption(
    "Interactive demonstration of an automated Python research framework. "
    "All portfolio holdings and position-level figures shown here are anonymized or synthetic."
)

view = st.radio(
    "Explore",
    ["Equity Research", "Portfolio Analytics", "Methodology"],
    horizontal=True,
)


def pct(x, d=1):
    return "—" if x is None or pd.isna(x) else f"{x:.{d}%}"


def num(x, d=2):
    return "—" if x is None or pd.isna(x) else f"{x:.{d}f}"


if view == "Equity Research":
    st.subheader("Equity Research Dashboard")
    st.caption("Demonstration company. Figures are illustrative and are not an investment recommendation.")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Current price", "$182.40")
    c2.metric("Base fair value", "$211.00")
    c3.metric("Modeled upside", "+15.7%")
    c4.metric("Quant score", "71 / 100")
    c5.metric("Model view", "Potentially attractive")

    hist = pd.DataFrame(
        {
            "Year": [2020, 2021, 2022, 2023, 2024, 2025],
            "Revenue": [82.0, 97.5, 111.4, 126.7, 145.2, 164.8],
            "Operating Income": [17.0, 22.8, 23.9, 29.4, 35.8, 41.2],
            "Free Cash Flow": [14.6, 19.1, 18.4, 23.5, 28.9, 33.7],
        }
    )
    hist_long = hist.melt("Year", var_name="Metric", value_name="$bn")
    fig = px.line(hist_long, x="Year", y="$bn", color="Metric", markers=True, title="Historical financial progression")
    st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)
    with left:
        segments = pd.DataFrame(
            {
                "Segment": ["Core Platform", "Cloud & Data", "Subscriptions", "Other"],
                "Revenue": [76.0, 46.0, 31.0, 11.8],
                "Segment Margin": [0.31, 0.26, 0.22, 0.08],
            }
        )
        fig = px.bar(segments, x="Segment", y="Revenue", title="Latest segment revenue ($bn)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(segments, use_container_width=True, hide_index=True)

    with right:
        scenarios = pd.DataFrame(
            {
                "Scenario": ["Severe Bear", "Bear", "Base", "Bull"],
                "Value / Share": [118.0, 157.0, 211.0, 268.0],
                "Probability": [0.10, 0.20, 0.50, 0.20],
            }
        )
        fig = px.bar(scenarios, x="Scenario", y="Value / Share", title="Scenario valuation")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(scenarios, use_container_width=True, hide_index=True)

    st.markdown("#### Research workflow demonstrated")
    st.write(
        "Historical statements → issuer/regulatory source checks → segment analysis → dynamic peers → "
        "DCF/scenario valuation → stress testing → model-quality controls → investment-summary synthesis."
    )
    st.info(
        "The production project additionally supports company Investor Relations pages, annual/results reports, "
        "SEC 10-K/20-F/40-F/6-K filings, dynamic peer selection, ownership analysis, news-impact analysis and downloadable Excel models."
    )

elif view == "Portfolio Analytics":
    st.subheader("Portfolio Analytics Dashboard")
    st.caption("Anonymous demonstration portfolio. No real holdings, portfolio value or cost basis are exposed.")

    metrics = {
        "Ann. return": 0.161,
        "Ann. volatility": 0.214,
        "Sharpe": 0.66,
        "Tracking error": 0.132,
        "Information ratio": 0.43,
        "Max drawdown": -0.286,
    }
    cols = st.columns(6)
    cols[0].metric("Ann. return", pct(metrics["Ann. return"]))
    cols[1].metric("Ann. volatility", pct(metrics["Ann. volatility"]))
    cols[2].metric("Sharpe", num(metrics["Sharpe"]))
    cols[3].metric("Tracking error", pct(metrics["Tracking error"]))
    cols[4].metric("Info ratio", num(metrics["Information ratio"]))
    cols[5].metric("Max drawdown", pct(metrics["Max drawdown"]))

    weights = pd.DataFrame(
        {
            "Holding": [f"Holding {c}" for c in "ABCDEFGH"],
            "Weight": [0.19, 0.16, 0.15, 0.13, 0.12, 0.10, 0.08, 0.07],
            "Risk Contribution": [0.22, 0.17, 0.16, 0.12, 0.12, 0.09, 0.07, 0.05],
        }
    )
    left, right = st.columns(2)
    with left:
        fig = px.pie(weights, names="Holding", values="Weight", title="Anonymous portfolio weights")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        risk_long = weights.melt("Holding", value_vars=["Weight", "Risk Contribution"], var_name="Series", value_name="Value")
        fig = px.bar(risk_long, x="Holding", y="Value", color="Series", barmode="group", title="Capital weight vs risk contribution")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Alpha & factor-adjusted performance")
    alpha = pd.DataFrame(
        {
            "Model": ["CAPM / Jensen", "Fama-French 3", "Carhart 4", "Fama-French 5", "Style Proxy"],
            "Annualized Alpha": [0.041, 0.033, 0.026, 0.030, 0.023],
            "t-stat": [1.82, 1.55, 1.27, 1.44, 1.10],
            "R²": [0.74, 0.80, 0.84, 0.83, 0.87],
        }
    )
    fig = px.bar(alpha, x="Model", y="Annualized Alpha", title="Residual alpha across risk models")
    fig.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(alpha, use_container_width=True, hide_index=True)

    factors = pd.DataFrame(
        {
            "Factor": ["Market", "Size", "Value", "Momentum", "Quality", "Low Volatility"],
            "Exposure": [1.08, -0.12, -0.28, 0.34, 0.29, -0.08],
        }
    )
    stress = pd.DataFrame(
        {
            "Scenario": ["Broad market -20%", "Rates +150 bps", "Growth selloff", "Credit shock"],
            "Estimated Portfolio Return": [-0.176, -0.082, -0.142, -0.094],
        }
    )
    left, right = st.columns(2)
    with left:
        fig = px.bar(factors, x="Factor", y="Exposure", title="Factor exposures")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = px.bar(stress, x="Scenario", y="Estimated Portfolio Return", title="Stress scenarios")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Portfolio methods demonstrated")
    st.write(
        "Benchmark-relative performance, Sharpe/Sortino, tracking error, information ratio, Jensen alpha, "
        "Fama-French/Carhart regressions, rolling alpha, factor exposures, risk budgeting, liquidity checks, "
        "stress testing, Monte Carlo analysis and constrained portfolio alternatives."
    )

else:
    st.subheader("Methodology & Architecture")
    st.write(
        "The production system is an automated equity-research and portfolio-analytics framework built in Python. "
        "A scheduled workflow refreshes market and issuer data, rebuilds research outputs, and serves dashboards through Streamlit."
    )

    architecture = pd.DataFrame(
        {
            "Layer": ["Sources", "Research engine", "Portfolio engine", "Automation", "Presentation"],
            "Examples": [
                "Issuer IR, annual/results reports, SEC filings, public market data",
                "Historical financials, segment analysis, peers, DCF, scenarios, news, quality controls",
                "Risk, alpha, factors, attribution, liquidity, stress, Monte Carlo, optimization",
                "Scheduled GitHub Actions refresh and encrypted private research bundle",
                "Private research portal + sanitized public Streamlit showcase",
            ],
        }
    )
    st.dataframe(architecture, use_container_width=True, hide_index=True)

    st.markdown("#### Privacy design")
    st.write(
        "The public showcase intentionally contains no real holdings, position values, average costs, transaction history, "
        "private workbook downloads, API credentials or GitHub Actions secrets. The production portfolio portal is separate and access-controlled."
    )

    st.markdown("#### Technologies")
    st.write("Python · pandas · NumPy · SciPy · Streamlit · Plotly · openpyxl · GitHub Actions · public issuer/regulatory data")

st.divider()
st.caption(
    f"Demonstration build · {date.today().isoformat()} · For portfolio/project demonstration only; not investment advice."
)
