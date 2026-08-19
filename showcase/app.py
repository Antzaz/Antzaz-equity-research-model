from __future__ import annotations

"""Recruiter-facing Streamlit showcase.

Equity-research examples are illustrative. Portfolio analytics load the newest validated,
sanitized snapshot from the public showcase GitHub repository whenever a user opens the app.
Company names and recruiter-safe investment theses are intentionally public; share counts,
cost basis, market value, transactions, tickers, private notes and credentials remain excluded.
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
    "Interactive demonstration of an automated Python research framework, real portfolio "
    "analytics and documented investment theses."
)

view = st.radio(
    "Explore",
    ["Portfolio Analytics", "Investment Thesis", "Equity Research", "Methodology"],
    horizontal=True,
)


def pct(x, d=1):
    return "—" if x is None or pd.isna(x) else f"{x:.{d}%}"


def num(x, d=2):
    return "—" if x is None or pd.isna(x) else f"{x:.{d}f}"


def _valid_snapshot(data):
    return (
        isinstance(data, dict)
        and data.get("snapshot_type") == "sanitized_real_portfolio_analytics"
    )


def load_snapshot():
    """Fetch the newest public GitHub snapshot for each Streamlit script session."""
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


def require_snapshot():
    if not snapshot:
        st.error(
            "The validated real portfolio snapshot is temporarily unavailable. "
            "No demonstration portfolio is shown in its place."
        )
        st.info("Reload after the next successful GitHub portfolio refresh.")
        st.stop()


def show_freshness():
    st.caption(f"Data source: {snapshot_source}. The app checks GitHub when a new session opens.")
    generated = snapshot.get("generated_utc") if snapshot else None
    if generated:
        st.caption(f"Portfolio analytics generated: {generated}")


def show_text_section(title, value):
    if value:
        st.markdown(f"#### {title}")
        st.write(value)


snapshot, snapshot_source = load_snapshot()


if view == "Portfolio Analytics":
    require_snapshot()
    st.subheader("Portfolio Analytics Dashboard")
    st.caption(
        "Real production analytics with company names and portfolio weights. Sensitive position "
        "economics such as shares, average cost and total portfolio value remain private."
    )
    show_freshness()

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
        growth = timeseries[cols_to_show].melt(
            "date", var_name="Series", value_name="Growth"
        )
        fig = px.line(
            growth,
            x="date",
            y="Growth",
            color="Series",
            title="Portfolio vs benchmark growth path",
        )
        st.plotly_chart(fig, use_container_width=True)

    if not holdings.empty:
        left, right = st.columns(2)
        with left:
            fig = px.pie(
                holdings,
                names="company",
                values="weight",
                title="Current portfolio weights",
            )
            st.plotly_chart(fig, use_container_width=True)
        with right:
            value_vars = [
                c for c in ["weight", "risk_contribution"] if c in holdings.columns
            ]
            if value_vars:
                risk_long = holdings.melt(
                    "company",
                    value_vars=value_vars,
                    var_name="Series",
                    value_name="Value",
                )
                fig = px.bar(
                    risk_long,
                    x="company",
                    y="Value",
                    color="Series",
                    barmode="group",
                    title="Capital weight vs risk contribution",
                )
                fig.update_yaxes(tickformat=".0%")
                st.plotly_chart(fig, use_container_width=True)

        display = holdings.copy()
        display = display.rename(
            columns={
                "company": "Company",
                "sector": "Sector",
                "weight": "Portfolio Weight",
                "risk_contribution": "Risk Contribution",
            }
        )
        for col in ["Portfolio Weight", "Risk Contribution"]:
            if col in display.columns:
                display[col] = display[col].map(lambda x: pct(x) if pd.notna(x) else "—")
        st.markdown("#### Holdings")
        st.dataframe(display, use_container_width=True, hide_index=True)

        if "sector" in holdings.columns and holdings["sector"].notna().any():
            sector = (
                holdings.dropna(subset=["sector"])
                .groupby("sector", as_index=False)["weight"]
                .sum()
                .sort_values("weight", ascending=False)
            )
            fig = px.bar(
                sector,
                x="sector",
                y="weight",
                title="Sector allocation",
            )
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Alpha & factor-adjusted performance")
    if not alpha.empty:
        fig = px.bar(
            alpha, x="model", y="annualized_alpha", title="Residual alpha across risk models"
        )
        fig.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig, use_container_width=True)
        display_cols = [
            c
            for c in [
                "model",
                "annualized_alpha",
                "t_stat",
                "p_value",
                "r2",
                "significant_5pct",
                "interpretation",
            ]
            if c in alpha.columns
        ]
        st.dataframe(alpha[display_cols], use_container_width=True, hide_index=True)
    else:
        st.info("No alpha regression results were available in the latest snapshot.")

    left, right = st.columns(2)
    with left:
        if not factors.empty:
            fig = px.bar(
                factors, x="factor", y="exposure", title="Factor/style exposures"
            )
            st.plotly_chart(fig, use_container_width=True)
    with right:
        if not stress.empty:
            fig = px.bar(
                stress,
                x="scenario",
                y="estimated_return",
                title="Portfolio stress scenarios",
            )
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Public / private boundary")
    st.write(
        "Public: company names, portfolio weights, aggregate performance/risk analytics and "
        "recruiter-safe investment theses. Private: tickers in the data pipeline, share counts, "
        "average cost, exact portfolio value, unrealized P&L, transaction history, private notes, "
        "credentials and private Excel research models."
    )

elif view == "Investment Thesis":
    require_snapshot()
    st.subheader("Investment Philosophy & Company Theses")
    st.caption(
        "This section documents the reasoning behind portfolio construction and individual "
        "security selection. Thesis text is user-authored; company names and portfolio analytics "
        "come from the latest production portfolio run."
    )
    show_freshness()

    philosophy = snapshot.get("portfolio_philosophy") or {}
    theses = snapshot.get("theses") or []

    if philosophy:
        st.markdown("### Portfolio Philosophy")
        headline_cols = st.columns(3)
        headline_cols[0].metric("Benchmark", philosophy.get("benchmark", "—"))
        headline_cols[1].metric("Time horizon", philosophy.get("time_horizon", "—"))
        headline_cols[2].metric("Return objective", philosophy.get("return_objective", "—"))

        philosophy_sections = [
            ("Investment Philosophy", "investment_philosophy"),
            ("Portfolio Objective", "portfolio_objective"),
            ("Research Process", "research_process"),
            ("Selection Criteria", "selection_criteria"),
            ("Position Sizing", "position_sizing"),
            ("Diversification", "diversification"),
            ("Risk Management", "risk_management"),
            ("Sell Discipline", "sell_discipline"),
            ("Monitoring & Review", "monitoring_and_review"),
            ("Portfolio Edge", "portfolio_edge"),
            ("What I Avoid", "what_i_avoid"),
            ("Closing Summary", "closing_summary"),
        ]
        for title, key in philosophy_sections:
            show_text_section(title, philosophy.get(key))

    if not theses:
        st.info(
            "No company thesis rows have been published yet. Fill the Portfolio Investment "
            "Thesis Excel workbook and run the normal portfolio sync command."
        )
    else:
        st.markdown("### Company Thesis Overview")
        overview_rows = []
        for item in theses:
            overview_rows.append(
                {
                    "Company": item.get("company"),
                    "Weight": item.get("weight"),
                    "Status": item.get("status"),
                    "Conviction": item.get("conviction"),
                    "Composite Score": item.get("composite_score"),
                    "Expected Annual Return": item.get("expected_annual_return"),
                }
            )
        overview = pd.DataFrame(overview_rows)
        if "Weight" in overview.columns:
            overview["Weight"] = overview["Weight"].map(pct)
        if "Expected Annual Return" in overview.columns:
            overview["Expected Annual Return"] = overview["Expected Annual Return"].map(pct)
        st.dataframe(overview, use_container_width=True, hide_index=True)

        companies = [item.get("company") for item in theses if item.get("company")]
        selected = st.selectbox("Select a portfolio company", companies)
        item = next(x for x in theses if x.get("company") == selected)

        st.markdown(f"## {selected}")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Portfolio weight", pct(item.get("weight")))
        m2.metric("Conviction", num(item.get("conviction"), 1) + " / 5" if item.get("conviction") is not None else "—")
        m3.metric("Research score", num(item.get("composite_score"), 2) + " / 5" if item.get("composite_score") is not None else "—")
        m4.metric("Expected annual return", pct(item.get("expected_annual_return")))
        m5.metric("Time horizon", item.get("time_horizon", "—"))

        scores = item.get("scores") or {}
        if scores:
            label_map = {
                "business_quality": "Business Quality",
                "moat_score": "Moat",
                "management_capital_allocation": "Management / Capital Allocation",
                "balance_sheet": "Balance Sheet",
                "growth": "Growth",
                "valuation": "Valuation",
                "risk_resilience": "Risk / Resilience",
            }
            score_df = pd.DataFrame(
                [
                    {"Criterion": label_map.get(k, k), "Score": v}
                    for k, v in scores.items()
                ]
            )
            fig = px.bar(
                score_df,
                x="Criterion",
                y="Score",
                range_y=[0, 5],
                title="Investment criteria scorecard",
            )
            st.plotly_chart(fig, use_container_width=True)

        show_text_section("Investment Thesis", item.get("investment_thesis"))
        show_text_section("Why I Own It", item.get("why_owned"))

        left, right = st.columns(2)
        with left:
            show_text_section(
                "Competitive Advantage / Moat", item.get("competitive_advantage")
            )
            show_text_section("Growth Drivers", item.get("growth_drivers"))
            show_text_section("Valuation Rationale", item.get("valuation_rationale"))
            show_text_section("Catalysts", item.get("catalysts"))
        with right:
            show_text_section("Key Risks", item.get("key_risks"))
            show_text_section(
                "Falsification / Sell Condition", item.get("sell_condition")
            )
            show_text_section("Monitoring KPI", item.get("monitoring_kpi"))
            show_text_section("Public Notes", item.get("public_notes"))

        details = []
        if item.get("status"):
            details.append(f"Status: {item['status']}")
        if item.get("review_date"):
            details.append(f"Last/next review: {item['review_date']}")
        if item.get("sector"):
            details.append(f"Sector: {item['sector']}")
        if details:
            st.caption(" · ".join(details))

elif view == "Equity Research":
    st.subheader("Equity Research Dashboard")
    st.caption(
        "Demonstration company. Figures are illustrative and are not an investment recommendation."
    )

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
        fig = px.bar(
            segments, x="Segment", y="Revenue", title="Latest segment revenue ($bn)"
        )
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
        fig = px.bar(
            scenarios, x="Scenario", y="Value / Share", title="Scenario valuation"
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(scenarios, use_container_width=True, hide_index=True)

    st.markdown("#### Research workflow demonstrated")
    st.write(
        "Historical statements → issuer/regulatory source checks → segment analysis → dynamic "
        "peers → DCF/scenario valuation → stress testing → model-quality controls → "
        "investment-summary synthesis."
    )
    st.info(
        "The production project additionally supports company Investor Relations pages, "
        "annual/results reports, SEC 10-K/20-F/40-F/6-K filings, dynamic peer selection, "
        "ownership analysis, news-impact analysis and downloadable Excel models."
    )

else:
    st.subheader("Methodology & Architecture")
    st.write(
        "The production system is an automated equity-research and portfolio-analytics framework "
        "built in Python. A scheduled workflow refreshes market and issuer data, rebuilds research "
        "outputs, validates the recruiter-safe snapshot and serves the latest data through Streamlit."
    )

    architecture = pd.DataFrame(
        {
            "Layer": [
                "Sources",
                "Research engine",
                "Portfolio engine",
                "Investment thesis",
                "Automation",
                "Presentation",
            ],
            "Examples": [
                "Issuer IR, annual/results reports, SEC filings, public market data",
                "Historical financials, segment analysis, peers, DCF, scenarios, news, quality controls",
                "Risk, alpha, factors, attribution, liquidity, stress, Monte Carlo, optimization",
                "Private Excel thesis workbook with public-field controls and scorecards",
                "Scheduled GitHub Actions refresh, sanitization and validation",
                "Public Streamlit app fetches the newest sanitized GitHub snapshot when opened",
            ],
        }
    )
    st.dataframe(architecture, use_container_width=True, hide_index=True)

    st.markdown("#### Data integrity")
    st.write(
        "Company names, holdings weights and portfolio analytics are produced from the latest "
        "portfolio research run. Thesis text and scorecards are authored in the private Excel "
        "workbook. The public snapshot is validated before publication."
    )

    st.markdown("#### Public showcase data policy")
    st.write(
        "Company names, weights, aggregate analytics and explicitly recruiter-safe thesis content "
        "are public. Share counts, average costs, exact portfolio value, transactions, private "
        "notes, tickers in the data pipeline, credentials and private workbooks remain excluded."
    )

    st.markdown("#### Technologies")
    st.write(
        "Python · pandas · NumPy · SciPy · Streamlit · Plotly · openpyxl · GitHub Actions · "
        "public issuer/regulatory data"
    )

st.divider()
st.caption(
    f"Showcase build · {date.today().isoformat()} · For portfolio/project demonstration only; not investment advice."
)
