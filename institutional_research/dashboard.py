from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE = Path(__file__).resolve().parent
OUT = BASE / "outputs" / "latest"

st.set_page_config(page_title="Institutional Research Dashboard", layout="wide")
st.title("Institutional Portfolio Research Dashboard")

if not OUT.exists():
    st.error("No analysis outputs found. Run: python run_research.py")
    st.stop()

with open(OUT / "summary.json", "r", encoding="utf-8") as f:
    summary = json.load(f)

portfolio = summary["portfolio"]
mc = summary["monte_carlo"]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ann. return", f"{portfolio['annualized_return']:.1%}")
c2.metric("Ann. volatility", f"{portfolio['annualized_volatility']:.1%}")
c3.metric("Sharpe", "—" if portfolio["sharpe"] is None else f"{portfolio['sharpe']:.2f}")
c4.metric("Beta", "—" if portfolio["beta"] is None else f"{portfolio['beta']:.2f}")
c5.metric("Max drawdown", f"{portfolio['max_drawdown']:.1%}")

tabs = st.tabs([
    "Portfolio",
    "Risk",
    "Factors",
    "Monte Carlo",
    "Reverse DCF",
    "Forecasts",
    "Power BI",
])

with tabs[0]:
    holdings = pd.read_csv(OUT / "holdings_analysis.csv")
    sector = pd.read_csv(OUT / "sector_exposure.csv")
    ts = pd.read_csv(OUT / "portfolio_timeseries.csv", parse_dates=["Date"])

    left, right = st.columns(2)
    with left:
        fig = px.pie(holdings, names="Ticker", values="Weight", title="Portfolio weights")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = px.bar(sector, x="Sector", y="Weight", title="Sector exposure")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    growth = ts[["Date", "PortfolioGrowth", "BenchmarkGrowth"]].melt(
        "Date", var_name="Series", value_name="Growth"
    )
    fig = px.line(growth, x="Date", y="Growth", color="Series", title="Portfolio vs benchmark growth")
    st.plotly_chart(fig, use_container_width=True)

    cols = [
        "Ticker", "Company", "Weight", "CurrentPrice", "MarketValue", "UnrealizedPnLPct",
        "AnnualizedVolatility", "Beta", "MaxDrawdown", "RiskContributionPct",
    ]
    st.dataframe(holdings[[c for c in cols if c in holdings.columns]], use_container_width=True)

with tabs[1]:
    risk = pd.read_csv(OUT / "risk_contribution.csv")
    corr = pd.read_csv(OUT / "correlation_matrix.csv").set_index("Ticker")

    c1, c2, c3 = st.columns(3)
    c1.metric("Daily VaR 95%", f"{portfolio['daily_var_95']:.2%}")
    c2.metric("Daily Expected Shortfall 95%", f"{portfolio['daily_expected_shortfall_95']:.2%}")
    c3.metric("Annualized alpha", "—" if portfolio["annualized_alpha"] is None else f"{portfolio['annualized_alpha']:.1%}")

    fig = px.bar(risk, x="Ticker", y="RiskContributionPct", title="Contribution to portfolio volatility")
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    fig = px.imshow(corr, text_auto=".2f", aspect="auto", title="Correlation matrix", zmin=-1, zmax=1)
    st.plotly_chart(fig, use_container_width=True)

    stress = pd.read_csv(OUT / "stress_tests.csv")
    portfolio_stress = stress[stress["Ticker"] == "PORTFOLIO"].copy()
    fig = px.bar(
        portfolio_stress,
        x="Scenario",
        y="EstimatedHoldingReturn",
        title="Configured portfolio stress scenarios",
    )
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(stress, use_container_width=True)

with tabs[2]:
    scores = pd.read_csv(OUT / "factor_scores.csv")
    exposure = pd.read_csv(OUT / "factor_exposure.csv")

    fig = px.bar(exposure, x="Factor", y="PortfolioExposure", title="Portfolio factor exposure (portfolio-relative z-scores)")
    st.plotly_chart(fig, use_container_width=True)

    factor_cols = [
        "Ticker", "ValueScore", "QualityScore", "GrowthScore",
        "MomentumScore", "LowVolatilityScore", "CompositeScore", "PortfolioWeight",
    ]
    st.dataframe(scores[factor_cols], use_container_width=True)

with tabs[3]:
    dist = pd.read_csv(OUT / "monte_carlo_distribution.csv")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Median 1Y return", f"{mc['median_return']:.1%}")
    c2.metric("5th percentile", f"{mc['p05_return']:.1%}")
    c3.metric("P(loss)", f"{mc['probability_loss']:.1%}")
    c4.metric("P(loss >20%)", f"{mc['probability_loss_20pct']:.1%}")

    fig = px.histogram(
        dist,
        x="TerminalReturn",
        nbins=80,
        title="Bootstrap Monte Carlo terminal return distribution",
    )
    fig.update_xaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "This is a historical bootstrap simulation: it preserves the empirical daily "
        "portfolio-return distribution better than a simple normal model, but it is not a forecast."
    )

with tabs[4]:
    dcf = pd.read_csv(OUT / "reverse_dcf.csv")
    st.subheader("Market-implied FCF growth")
    st.caption(
        "Solves the constant annual FCF growth rate that makes a simplified 10-year DCF "
        "equal the current market capitalization under the configured WACC and terminal growth."
    )
    st.dataframe(dcf, use_container_width=True)

with tabs[5]:
    forecasts_path = OUT / "forecast_accuracy.csv"
    forecasts = pd.read_csv(forecasts_path) if forecasts_path.exists() else pd.DataFrame()
    if forecasts.empty:
        st.info(
            "No forecasts logged yet. Add rows to forecasts.csv before results are known, "
            "then later fill Actual and rerun the analysis."
        )
    else:
        completed = forecasts[forecasts["Actual"].notna()].copy()
        if not completed.empty:
            c1, c2 = st.columns(2)
            c1.metric("Mean abs. forecast error", f"{completed['AbsoluteErrorPct'].mean():.1%}")
            c2.metric("Median abs. forecast error", f"{completed['AbsoluteErrorPct'].median():.1%}")
        st.dataframe(forecasts, use_container_width=True)

with tabs[6]:
    st.subheader("Power BI-ready output")
    st.write(
        "Every analysis table is exported as a flat CSV under outputs/latest. "
        "Power BI can connect directly to this folder or to individual CSV files."
    )
    st.code(str(OUT))
    st.write(
        "Recommended Power BI tables: holdings_analysis, portfolio_timeseries, "
        "risk_contribution, factor_scores, factor_exposure, stress_tests, "
        "monte_carlo_distribution, reverse_dcf and forecast_accuracy."
    )
