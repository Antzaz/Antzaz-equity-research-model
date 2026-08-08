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


def read(name: str) -> pd.DataFrame:
    path = OUT / f"{name}.csv"
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def pct(value, digits=1):
    return "—" if value is None or pd.isna(value) else f"{value:.{digits}%}"


def num(value, digits=2):
    return "—" if value is None or pd.isna(value) else f"{value:.{digits}f}"


c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Ann. return", pct(portfolio.get("annualized_return")))
c2.metric("Ann. volatility", pct(portfolio.get("annualized_volatility")))
c3.metric("Sharpe", num(portfolio.get("sharpe")))
c4.metric("Tracking error", pct(portfolio.get("tracking_error")))
c5.metric("Info ratio", num(portfolio.get("information_ratio")))
c6.metric("Max drawdown", pct(portfolio.get("max_drawdown")))

tabs = st.tabs([
    "Portfolio",
    "Risk",
    "Construction",
    "Factors",
    "Attribution",
    "Liquidity",
    "Monte Carlo",
    "Reverse DCF",
    "Forecasts",
    "Power BI",
])

with tabs[0]:
    holdings = read("holdings_analysis")
    sector = read("sector_exposure")
    ts = read("portfolio_timeseries")
    if not ts.empty and "Date" in ts:
        ts["Date"] = pd.to_datetime(ts["Date"])

    left, right = st.columns(2)
    with left:
        if not holdings.empty:
            fig = px.pie(holdings, names="Ticker", values="Weight", title="Portfolio weights")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        if not sector.empty:
            fig = px.bar(sector, x="Sector", y="Weight", title="Sector exposure")
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    if not ts.empty:
        growth = ts[["Date", "PortfolioGrowth", "BenchmarkGrowth"]].melt(
            "Date", var_name="Series", value_name="Growth"
        )
        fig = px.line(growth, x="Date", y="Growth", color="Series", title="Portfolio vs benchmark growth")
        st.plotly_chart(fig, use_container_width=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Top 1 weight", pct(portfolio.get("top_1_weight")))
    c2.metric("Top 5 weight", pct(portfolio.get("top_5_weight")))
    c3.metric("Effective holdings", num(portfolio.get("effective_number_of_holdings"), 1))
    c4.metric("Active Share", pct(portfolio.get("active_share")))

    cols = [
        "Ticker", "Company", "Weight", "CurrentPrice", "MarketValue", "UnrealizedPnLPct",
        "AnnualizedVolatility", "Beta", "MaxDrawdown", "RiskContributionPct",
    ]
    if not holdings.empty:
        st.dataframe(holdings[[c for c in cols if c in holdings.columns]], use_container_width=True)

with tabs[1]:
    risk = read("risk_contribution")
    risk_budget = read("risk_budget")
    corr = read("correlation_matrix")
    if not corr.empty and "Ticker" in corr:
        corr = corr.set_index("Ticker")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Daily VaR 95%", pct(portfolio.get("daily_var_95"), 2))
    c2.metric("Daily Expected Shortfall 95%", pct(portfolio.get("daily_expected_shortfall_95"), 2))
    c3.metric("Beta", num(portfolio.get("beta")))
    c4.metric("Largest risk contribution", pct(portfolio.get("largest_risk_contribution")))

    if not risk.empty:
        fig = px.bar(risk, x="Ticker", y="RiskContributionPct", title="Contribution to portfolio volatility")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)
    if not risk_budget.empty:
        fig = px.bar(
            risk_budget,
            x="Ticker",
            y=["Weight", "RiskContributionPct"],
            barmode="group",
            title="Capital weight vs risk budget",
        )
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    if not corr.empty:
        fig = px.imshow(corr, text_auto=".2f", aspect="auto", title="Correlation matrix", zmin=-1, zmax=1)
        st.plotly_chart(fig, use_container_width=True)

    stress = read("stress_tests")
    if not stress.empty:
        portfolio_stress = stress[stress["Ticker"] == "PORTFOLIO"].copy()
        if not portfolio_stress.empty:
            fig = px.bar(
                portfolio_stress,
                x="Scenario",
                y="EstimatedHoldingReturn",
                title="Configured forward stress scenarios",
            )
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    historical = read("historical_stress_windows")
    if not historical.empty:
        st.subheader("Historical stress windows")
        st.dataframe(historical, use_container_width=True)

    rolling = read("rolling_risk")
    if not rolling.empty:
        st.subheader("Rolling risk")
        st.dataframe(rolling, use_container_width=True)

with tabs[2]:
    constraints = read("constraint_report")
    optimization = read("portfolio_optimizations")
    expected = read("expected_returns_inputs")
    active = read("active_share_detail")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Active return", pct(portfolio.get("active_annualized_return")))
    c2.metric("Tracking error", pct(portfolio.get("tracking_error")))
    c3.metric("Information ratio", num(portfolio.get("information_ratio")))
    c4.metric("Daily active hit rate", pct(portfolio.get("daily_active_hit_rate")))

    if not constraints.empty:
        st.subheader("Portfolio policy / constraint monitor")
        st.dataframe(constraints, use_container_width=True)

    if not optimization.empty:
        st.subheader("Constrained portfolio alternatives")
        choices = optimization["Portfolio"].dropna().unique().tolist()
        selected = st.selectbox("Portfolio construction method", choices)
        view = optimization[optimization["Portfolio"] == selected].copy()
        fig = px.bar(
            view,
            x="Ticker",
            y=["CurrentWeight", "TargetWeight"],
            barmode="group",
            title=f"{selected}: current vs target weights",
        )
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(view, use_container_width=True)
    else:
        st.info(
            "Minimum-variance and equal-risk portfolios need at least two holdings and enough history. "
            "Add expected_returns.csv to enable expected-return / max-Sharpe sizing."
        )

    if not expected.empty:
        st.subheader("Expected-return / conviction inputs")
        st.dataframe(expected, use_container_width=True)

    if not active.empty:
        st.subheader("Active Share detail")
        st.caption("Requires a user-supplied benchmark_weights.csv with constituent weights.")
        st.dataframe(active, use_container_width=True)

with tabs[3]:
    scores = read("factor_scores")
    exposure = read("factor_exposure")
    proxy = read("factor_proxy_sensitivity")

    if not exposure.empty:
        fig = px.bar(exposure, x="Factor", y="PortfolioExposure", title="Holding-relative factor scores")
        st.plotly_chart(fig, use_container_width=True)
    if not proxy.empty:
        st.subheader("Market-factor proxy sensitivities")
        st.caption(
            "Public ETF proxy betas are diagnostics, not a commercial Barra/Axioma-style factor risk model."
        )
        fig = px.bar(proxy, x="Proxy", y="BetaToProxy", title="Portfolio beta to factor proxies")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(proxy, use_container_width=True)
    if not scores.empty:
        factor_cols = [
            "Ticker", "ValueScore", "QualityScore", "GrowthScore",
            "MomentumScore", "LowVolatilityScore", "CompositeScore", "PortfolioWeight",
        ]
        st.dataframe(scores[[c for c in factor_cols if c in scores]], use_container_width=True)

with tabs[4]:
    attribution = read("return_attribution")
    rel = read("benchmark_relative")
    if not attribution.empty:
        fig = px.bar(
            attribution,
            x="Ticker",
            y="StaticArithmeticContribution",
            title="Static-weight arithmetic return contribution",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(attribution, use_container_width=True)
        st.caption(
            "This uses today's weights applied to historical returns. It is a research diagnostic, "
            "not transaction-level realized performance attribution."
        )
    if not rel.empty:
        st.subheader("Benchmark-relative metrics")
        st.dataframe(rel, use_container_width=True)

with tabs[5]:
    liquidity = read("liquidity_analysis")
    if not liquidity.empty:
        fig = px.bar(
            liquidity,
            x="Ticker",
            y="EstimatedDaysToLiquidate",
            title="Estimated days to liquidate at configured ADV participation",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(liquidity, use_container_width=True)
        st.caption(
            "Liquidity uses public average-volume snapshots and is only a first-pass capacity check."
        )

with tabs[6]:
    dist = read("monte_carlo_distribution")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Median 1Y return", pct(mc.get("median_return")))
    c2.metric("5th percentile", pct(mc.get("p05_return")))
    c3.metric("P(loss)", pct(mc.get("probability_loss")))
    c4.metric("P(loss >20%)", pct(mc.get("probability_loss_20pct")))

    if not dist.empty:
        fig = px.histogram(
            dist,
            x="TerminalReturn",
            nbins=80,
            title="Bootstrap Monte Carlo terminal return distribution",
        )
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Historical bootstrap preserves the empirical daily distribution better than a simple normal model, "
        "but it remains a historical scenario engine, not a forecast."
    )

with tabs[7]:
    dcf = read("reverse_dcf")
    st.subheader("Market-implied FCF growth")
    st.caption(
        "Solves the constant annual FCF growth rate that makes a simplified 10-year DCF "
        "equal the current market capitalization under the configured WACC and terminal growth."
    )
    if not dcf.empty:
        st.dataframe(dcf, use_container_width=True)

with tabs[8]:
    forecasts = read("forecast_accuracy")
    if forecasts.empty:
        st.info(
            "No forecasts logged yet. Add rows to forecasts.csv before results are known, "
            "then later fill Actual and rerun the analysis."
        )
    else:
        completed = forecasts[forecasts["Actual"].notna()].copy()
        if not completed.empty:
            c1, c2 = st.columns(2)
            c1.metric("Mean abs. forecast error", pct(completed["AbsoluteErrorPct"].mean()))
            c2.metric("Median abs. forecast error", pct(completed["AbsoluteErrorPct"].median()))
        st.dataframe(forecasts, use_container_width=True)

with tabs[9]:
    st.subheader("Power BI-ready output")
    st.write(
        "Every analysis table is exported as a flat CSV under outputs/latest. "
        "Power BI can connect directly to this folder or to individual CSV files."
    )
    st.code(str(OUT))
    st.write(
        "New professional tables include risk_budget, benchmark_relative, concentration_summary, "
        "liquidity_analysis, factor_proxy_sensitivity, rolling_risk, historical_stress_windows, "
        "return_attribution, portfolio_optimizations, constraint_report and active_share_detail."
    )
