from __future__ import annotations

"""Recruiter-facing investment research portfolio.

This app intentionally presents investment work, not source code. All recruiter-facing
company theses and public values are controlled from data/recruiter_portfolio.json.
Optional aggregate portfolio analytics can fall back to the sanitized production snapshot.
"""

import json
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE = Path(__file__).resolve().parent
PUBLIC_DATA_PATH = BASE / "data" / "recruiter_portfolio.json"
SNAPSHOT_PATH = BASE / "data" / "portfolio_snapshot.json"

st.set_page_config(
    page_title="Anton Hiltunen | Investment Research Portfolio",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 3rem;}
    div[data-testid="stMetric"] {
        border: 1px solid rgba(128,128,128,.22);
        border-radius: 12px;
        padding: .8rem 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def pct(value, digits=1):
    if value is None or pd.isna(value):
        return "—"
    try:
        return f"{float(value):.{digits}%}"
    except (TypeError, ValueError):
        return "—"


def num(value, digits=2):
    if value is None or pd.isna(value):
        return "—"
    try:
        return f"{float(value):.{digits}f}"
    except (TypeError, ValueError):
        return "—"


def money(value, currency="USD"):
    if value is None or pd.isna(value):
        return "—"
    symbols = {"USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥"}
    symbol = symbols.get(str(currency).upper(), f"{currency} ")
    try:
        return f"{symbol}{float(value):,.2f}"
    except (TypeError, ValueError):
        return "—"


def calc_upside(company):
    explicit = company.get("upside")
    if explicit is not None:
        return explicit
    current = company.get("current_price")
    fair = company.get("fair_value")
    try:
        if current and fair:
            return float(fair) / float(current) - 1
    except (TypeError, ValueError, ZeroDivisionError):
        pass
    return None


def safe_list(value):
    return value if isinstance(value, list) else []


def render_bullets(items):
    items = [str(x).strip() for x in safe_list(items) if str(x).strip()]
    if not items:
        st.write("—")
        return
    for item in items:
        st.markdown(f"- {item}")


def merge_metrics(public_metrics, snapshot):
    metrics = {}
    if snapshot:
        metrics.update(snapshot.get("metrics", {}) or {})
    metrics.update({k: v for k, v in (public_metrics or {}).items() if v is not None})
    return metrics


def published_companies(data):
    return [c for c in data.get("companies", []) if c.get("published", False)]


data = load_json(
    PUBLIC_DATA_PATH,
    {
        "profile": {},
        "portfolio": {},
        "portfolio_metrics": {},
        "performance": [],
        "companies": [],
    },
)
snapshot = load_json(SNAPSHOT_PATH, None)
profile = data.get("profile", {})
portfolio = data.get("portfolio", {})
companies = published_companies(data)
metrics = merge_metrics(data.get("portfolio_metrics", {}), snapshot)

name = profile.get("name", "Anton Hiltunen")
headline = profile.get("headline", "Investment Research Portfolio")
summary = profile.get(
    "summary",
    "Fundamental equity research, valuation, portfolio construction and risk analytics.",
)

st.title(name)
st.subheader(headline)
st.write(summary)

meta = []
if profile.get("location"):
    meta.append(profile["location"])
if profile.get("email"):
    meta.append(profile["email"])
if profile.get("linkedin_url"):
    meta.append(f"[LinkedIn]({profile['linkedin_url']})")
if meta:
    st.markdown(" · ".join(meta))

st.divider()

view = st.radio(
    "Explore",
    ["Portfolio", "Company Research", "Process & Methodology"],
    horizontal=True,
    label_visibility="collapsed",
)


if view == "Portfolio":
    st.header(portfolio.get("name", "Model Portfolio"))
    if portfolio.get("strategy"):
        st.write(portfolio["strategy"])

    detail_bits = []
    if portfolio.get("benchmark"):
        detail_bits.append(f"Benchmark: {portfolio['benchmark']}")
    if portfolio.get("inception_date"):
        detail_bits.append(f"Inception: {portfolio['inception_date']}")
    if portfolio.get("last_updated"):
        detail_bits.append(f"Last updated: {portfolio['last_updated']}")
    if detail_bits:
        st.caption(" · ".join(detail_bits))

    metric_cols = st.columns(6)
    metric_cols[0].metric("Ann. return", pct(metrics.get("annualized_return")))
    metric_cols[1].metric("Benchmark return", pct(metrics.get("benchmark_annualized_return")))
    alpha_value = metrics.get("active_annualized_return")
    if alpha_value is None:
        alpha_value = metrics.get("annualized_alpha")
    metric_cols[2].metric("Active return / alpha", pct(alpha_value))
    metric_cols[3].metric("Volatility", pct(metrics.get("annualized_volatility")))
    metric_cols[4].metric("Sharpe", num(metrics.get("sharpe")))
    metric_cols[5].metric("Max drawdown", pct(metrics.get("max_drawdown")))

    performance = pd.DataFrame(data.get("performance", []))
    if performance.empty and snapshot:
        performance = pd.DataFrame(snapshot.get("timeseries", []))

    if not performance.empty and "date" in performance.columns:
        performance["date"] = pd.to_datetime(performance["date"], errors="coerce")
        rename_map = {
            "portfolio_growth": "Portfolio",
            "benchmark_growth": "Benchmark",
            "portfolio": "Portfolio",
            "benchmark": "Benchmark",
        }
        performance = performance.rename(columns=rename_map)
        series_cols = [c for c in ["Portfolio", "Benchmark"] if c in performance.columns]
        if series_cols:
            growth = performance[["date"] + series_cols].melt("date", var_name="Series", value_name="Growth")
            fig = px.line(growth, x="date", y="Growth", color="Series", title="Portfolio vs benchmark growth")
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("Portfolio holdings")
    if not companies:
        st.info(
            "No companies are published yet. Add companies in "
            "`showcase/data/recruiter_portfolio.json` and set `published` to true."
        )
    else:
        rows = []
        for c in companies:
            rows.append(
                {
                    "Company": c.get("company", "—"),
                    "Ticker": c.get("ticker", "—"),
                    "Sector": c.get("sector", "—"),
                    "Weight": c.get("weight"),
                    "View": c.get("rating", "—"),
                    "Fair value": money(c.get("fair_value"), c.get("currency", "USD")),
                    "Upside": calc_upside(c),
                    "Conviction": c.get("conviction", "—"),
                }
            )
        holdings = pd.DataFrame(rows)
        st.dataframe(
            holdings,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Weight": st.column_config.NumberColumn(format="%.1%%"),
                "Upside": st.column_config.NumberColumn(format="%.1%%"),
            },
        )

        weighted = pd.DataFrame(
            [
                {"Company": c.get("company", c.get("ticker", "Holding")), "Weight": c.get("weight")}
                for c in companies
                if c.get("weight") is not None
            ]
        )
        if not weighted.empty:
            fig = px.pie(weighted, names="Company", values="Weight", title="Model portfolio allocation", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Research highlights")
        for c in companies:
            with st.container(border=True):
                left, right = st.columns([3, 1])
                with left:
                    st.markdown(f"### {c.get('company', 'Company')} ({c.get('ticker', '—')})")
                    one_liner = c.get("one_line_thesis")
                    if one_liner:
                        st.write(one_liner)
                    thesis = safe_list(c.get("thesis"))
                    if thesis:
                        for point in thesis[:3]:
                            st.markdown(f"- {point}")
                with right:
                    st.metric("View", c.get("rating", "—"))
                    st.metric("Upside", pct(calc_upside(c)))
                    st.caption(f"Conviction: {c.get('conviction', '—')}")

    disclosure = portfolio.get("disclosure")
    if disclosure:
        st.caption(disclosure)


elif view == "Company Research":
    st.header("Company Research")
    st.caption(
        "Condensed investment cases designed for fast review. "
        "Select a company to see the thesis, valuation, risks and falsification conditions."
    )

    if not companies:
        st.info(
            "No public company research has been added yet. "
            "Use `showcase/data/recruiter_portfolio.json` to publish a research card."
        )
    else:
        labels = {f"{c.get('company', 'Company')} ({c.get('ticker', '—')})": c for c in companies}
        selected_label = st.selectbox("Company", list(labels.keys()))
        c = labels[selected_label]
        currency = c.get("currency", "USD")

        st.subheader(selected_label)
        if c.get("one_line_thesis"):
            st.write(c["one_line_thesis"])

        cols = st.columns(6)
        cols[0].metric("View", c.get("rating", "—"))
        cols[1].metric("Current price", money(c.get("current_price"), currency))
        cols[2].metric("Fair value", money(c.get("fair_value"), currency))
        cols[3].metric("Upside", pct(calc_upside(c)))
        cols[4].metric("Conviction", c.get("conviction", "—"))
        cols[5].metric("Horizon", c.get("investment_horizon", "—"))

        if c.get("last_reviewed"):
            st.caption(f"Last reviewed: {c['last_reviewed']}")

        left, right = st.columns(2)
        with left:
            st.markdown("### Investment thesis")
            render_bullets(c.get("thesis"))
        with right:
            st.markdown("### Key catalysts")
            render_bullets(c.get("catalysts"))

        scorecard = c.get("scorecard", {})
        if scorecard:
            st.markdown("### Research scorecard")
            score_rows = [
                {"Dimension": k.replace("_", " ").title(), "Score": v}
                for k, v in scorecard.items()
                if v is not None
            ]
            if score_rows:
                score_df = pd.DataFrame(score_rows)
                fig = px.bar(score_df, x="Dimension", y="Score", range_y=[0, 10], title="Research scorecard (0-10)")
                st.plotly_chart(fig, use_container_width=True)

        valuation = pd.DataFrame(c.get("valuation_scenarios", []))
        if not valuation.empty:
            st.markdown("### Valuation")
            display_valuation = valuation.copy()
            if "value" in display_valuation.columns:
                display_valuation["Value"] = display_valuation["value"].apply(lambda x: money(x, currency))
                display_valuation = display_valuation.drop(columns=["value"])
            st.dataframe(display_valuation, use_container_width=True, hide_index=True)

        left, right = st.columns(2)
        with left:
            st.markdown("### Key risks")
            render_bullets(c.get("risks"))
        with right:
            st.markdown("### What would prove the thesis wrong?")
            render_bullets(c.get("falsification_conditions"))

        kpis = pd.DataFrame(c.get("key_kpis", []))
        if not kpis.empty:
            st.markdown("### KPIs I monitor")
            st.dataframe(kpis, use_container_width=True, hide_index=True)

        institutional = safe_list(c.get("institutional_comparison"))
        if institutional:
            st.markdown("### Institutional / market comparison")
            render_bullets(institutional)

        notes = c.get("notes")
        if notes:
            st.markdown("### Research notes")
            st.write(notes)

        sources = safe_list(c.get("sources"))
        if sources:
            st.markdown("### Selected sources")
            for source in sources:
                if isinstance(source, dict):
                    label = source.get("label", source.get("url", "Source"))
                    url = source.get("url")
                    if url:
                        st.markdown(f"- [{label}]({url})")
                    else:
                        st.markdown(f"- {label}")
                else:
                    st.markdown(f"- {source}")

        report_url = c.get("full_report_url")
        if report_url:
            st.link_button("Open full research report", report_url)


else:
    st.header("Process & Methodology")
    st.write(
        "I use a repeatable research process designed to separate evidence, assumptions "
        "and judgement. The public portfolio is a recruiter-facing summary of that work."
    )

    process = [
        ("1. Screen", "Identify companies worth deeper work."),
        ("2. Fundamentals", "Normalize financials, segments, KPIs and balance-sheet quality."),
        ("3. Forecast", "Build explicit operating assumptions and scenario ranges."),
        ("4. Value", "DCF, reverse DCF, peer checks and scenario / Monte Carlo analysis."),
        ("5. Challenge", "Compare with market expectations and define what could prove the thesis wrong."),
        ("6. Construct", "Size positions with portfolio risk, concentration and factor exposure in mind."),
        ("7. Monitor", "Track earnings, KPIs, thesis changes and decision quality over time."),
    ]
    for title, body in process:
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.write(body)

    st.subheader("Public presentation policy")
    st.write(
        "The recruiter view contains only information intentionally marked for publication. "
        "Private holdings, transaction history, cost basis, portfolio value, credentials, "
        "private Excel models and research-engine internals are not required for this showcase."
    )

    st.subheader("Tools used")
    st.write(
        "Excel · Python · pandas · NumPy · Streamlit · Plotly · Git/GitHub · "
        "issuer filings and public market data"
    )


st.divider()
footer = profile.get("footer", "For project demonstration only; not investment advice.")
st.caption(f"{footer} · {date.today().isoformat()}")
