from __future__ import annotations

"""Resume-safe Streamlit showcase.

Equity-research examples are illustrative. Portfolio analytics load the newest sanitized
snapshot from the public showcase GitHub repository whenever a user opens the app. A bundled
snapshot is retained as a fallback. The snapshot excludes tickers, company names, shares,
cost basis, market value, transactions, private workbooks and credentials.
"""

import json
import time
from datetime import date
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd
import plotly.express as px
import streamlit as st

BASE = Path(__file__).resolve().parent
SNAPSHOT_PATH = BASE / "data" / "portfolio_snapshot.json"
LIVE_SNAPSHOT_URL = (
    "https://raw.githubusercontent.com/Antzaz/"
    "Antzaz-investment-research-showcase/main/data/portfolio_snapshot.json"
)

st.set_page_config(page_title="Investment Research & Portfolio Analytics", layout="wide")
st.title("Investment Research & Portfolio Analytics")
st.caption(
    "Interactive demonstration of an automated Python research framework. "
    "Portfolio analytics can reflect the real production portfolio while holdings remain anonymized."
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


def _valid_snapshot(data):
    return isinstance(data, dict) and data.get("snapshot_type") == "sanitized_real_portfolio_analytics"


def load_snapshot():
    """Fetch the newest public GitHub snapshot for each Streamlit script session.

    A cache-busting query parameter is used so a new browser session/reload asks GitHub for
    the current file. If GitHub is temporarily unavailable, the bundled validated snapshot
    is used as a fallback.
    """
    try:
        request = Request(
            f"{LIVE_SNAPSHOT_URL}?v={int(time.time())}",
            headers={
                "User-Agent": "Antzaz-investment-research-showcase",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
        )
        with urlopen(request, timeout=5) as response:
            remote = json.loads(response.read().decode("utf-8"))
        if _valid_snapshot(remote):
            return remote, "Live GitHub snapshot"
    except Exception:
        pass

    if SNAPSHOT_PATH.exists():
        try:
            local = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
            if _valid_snapshot(local):
                return local, "Bundled fallback snapshot"
        except Exception:
            pass

    return None, "Unavailable"


snapshot, snapshot_source = load_snapshot()


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
    fig = px.line(
        hist.melt("Year", var_name="Metric", value_name="$bn"),
        x="Year",
        y="$bn",
        color="Metric",
        markers=True,
        title="Historical financial progression",
    )
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

    if not snapshot:
        st.error(
            "The validated real portfolio snapshot is temporarily unavailable. "
            "No demonstration portfolio is shown in its place."
        )
        st.info(
            "The public dashboard only displays validated, sanitized production analytics. "
            "Reload after the next GitHub portfolio refresh."
        )
        st.stop()

    st.caption(
        "Real aggregate analytics from the production portfolio. Holdings are anonymized and sensitive position economics are excluded."
    )
    st.caption(f"Data source: {snapshot_source}. The app checks GitHub whenever a new session opens.")
    generated = snapshot.get("generated_utc")
    if generated:
        st.caption(f"Sanitized analytics generated: {generated}")

    metrics = snapshot.get("metrics", {})
    holdings = pd.DataFrame(snapshot.get("holdings", []))
    alpha = pd.DataFrame(snapshot.get("alpha", []))
    factors = pd.DataFrame(snapshot.get("factors", []))
    stress = pd.DataFrame(snapshot.get("stress", []))
    timeseries = pd.DataFrame(snapshot.get("timeseries", []))

    cols = st.columns(6)
    cols[0].metric("Ann. return", pct(metrics.get("annualized_return")))
    cols[1].metric("Ann. volatility", pct(metrics.get("annualized_volatility")))
    cols[2].metric("Sharpe", num(metrics.get("sharpe")))
    cols[3].metric("Tracking error", pct(metrics.get("tracking_error")))
    cols[4].metric("Info ratio", num(metrics.get("information_ratio")))
    cols[5].metric("Max drawdown", pct(metrics.get("max_drawdown")))

    extra1, extra2, extra3, extra4 = st.columns(4)
    extra1.metric("Beta", num(metrics.get("beta")))
    extra2.metric("Sortino", num(metrics.get("sortino")))
    extra3.metric("Active return", pct(metrics.get("active_annualized_return")))
    extra4.metric("Daily ES 95%", pct(metrics.get("daily_expected_shortfall_95"), 2))

    if not timeseries.empty and {"date", "portfolio_growth"}.issubset(timeseries.columns):
        timeseries["date"] = pd.to_datetime(timeseries["date"], errors="coerce")
        cols_to_show = ["date", "portfolio_growth"]
        if "benchmark_growth" in timeseries.columns:
            cols_to_show.append("benchmark_growth")
        growth = timeseries[cols_to_show].melt("date", var_name="Series", value_name="Growth")
        fig = px.line(growth, x="date", y="Growth", color="Series", title="Real portfolio vs benchmark growth path")
        st.plotly_chart(fig, use_container_width=True)

    if not holdings.empty:
        left, right = st.columns(2)
        with left:
            fig = px.pie(holdings, names="holding", values="weight", title="Anonymous real portfolio weights")
            st.plotly_chart(fig, use_container_width=True)
        with right:
            value_vars = [c for c in ["weight", "risk_contribution"] if c in holdings.columns]
            risk_long = holdings.melt("holding", value_vars=value_vars, var_name="Series", value_name="Value")
            fig = px.bar(
                risk_long,
                x="holding",
                y="Value",
                color="Series",
                barmode="group",
                title="Capital weight vs risk contribution",
            )
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Alpha & factor-adjusted performance")
    if not alpha.empty:
        fig = px.bar(alpha, x="model", y="annualized_alpha", title="Residual alpha across risk models")
        fig.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig, use_container_width=True)
        display_cols = [c for c in ["model", "annualized_alpha", "t_stat", "p_value", "r2", "significant_5pct", "interpretation"] if c in alpha.columns]
        st.dataframe(alpha[display_cols], use_container_width=True, hide_index=True)
    else:
        st.info("No alpha regression results were available in the sanitized snapshot.")

    left, right = st.columns(2)
    with left:
        if not factors.empty:
            fig = px.bar(factors, x="factor", y="exposure", title="Factor/style exposures")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        if not stress.empty:
            fig = px.bar(stress, x="scenario", y="estimated_return", title="Portfolio stress scenarios")
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Privacy boundary")
    st.write(
        "The displayed portfolio analytics are real, but the public snapshot contains no ticker symbols, company names, "
        "share counts, average costs, portfolio value, unrealized P&L, transaction history, private Excel models or credentials."
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
                "Public Streamlit app fetches the newest sanitized GitHub snapshot when opened",
            ],
        }
    )
    st.dataframe(architecture, use_container_width=True, hide_index=True)

    st.markdown("#### Public showcase data policy")
    st.write(
        "The public portfolio dashboard uses genuine aggregate model outputs while all position identifiers and sensitive "
        "economics are removed before publication. Equity-research demonstration data remains illustrative unless explicitly published as a public case study."
    )

    st.markdown("#### Technologies")
    st.write("Python · pandas · NumPy · SciPy · Streamlit · Plotly · openpyxl · GitHub Actions · public issuer/regulatory data")

st.divider()
st.caption(
    f"Showcase build · {date.today().isoformat()} · For portfolio/project demonstration only; not investment advice."
)
