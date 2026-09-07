from __future__ import annotations

from pathlib import Path
import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from .ai_optionality import (
    DEFAULT_SCENARIOS,
    build_overlay_table,
    normalize_scenarios,
    portfolio_ai_scenarios,
    sensitivity_grid,
)
from .ai_research_inputs import suggested_ai_config


BASE = Path(__file__).resolve().parents[1]
ROOT = Path(__file__).resolve().parents[2]
OUT = BASE / "outputs" / "latest"
LOCAL_CONFIG = BASE / "ai_optionality_inputs.csv"


def _read(name: str) -> pd.DataFrame:
    path = OUT / f"{name}.csv"
    try:
        return pd.read_csv(path) if path.exists() else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def _pct(value, digits=1) -> str:
    try:
        x = float(value)
        return f"{x:.{digits}%}" if np.isfinite(x) else "—"
    except Exception:
        return "—"


def _load_expected_returns() -> tuple[pd.DataFrame, str]:
    expected = _read("optimizer_expected_returns")
    if not expected.empty and "Ticker" in expected:
        return expected, "Maximum-data optimizer export"

    fallback = _read("expected_returns_inputs")
    if not fallback.empty and {"Ticker", "ExpectedReturn"}.issubset(fallback.columns):
        out = fallback[["Ticker", "ExpectedReturn"]].copy()
        out["HistoricalPrior"] = pd.to_numeric(out["ExpectedReturn"], errors="coerce")
        out["ClassicalExpectedReturn"] = out["HistoricalPrior"]
        out["BlendedMLExpectedReturn"] = out["HistoricalPrior"]
        return out.drop(columns=["ExpectedReturn"]), "Manual expected-return fallback"

    return pd.DataFrame(), "Unavailable"


def _load_current_weights() -> tuple[pd.DataFrame, str]:
    weights_raw = _read("optimizer_weights")
    if not weights_raw.empty and {"Ticker", "CurrentWeight"}.issubset(weights_raw.columns):
        w = weights_raw[["Ticker", "CurrentWeight"]].drop_duplicates("Ticker").rename(columns={"CurrentWeight": "Weight"})
        w["Weight"] = pd.to_numeric(w["Weight"], errors="coerce").fillna(0.0)
        return w, "Optimizer current weights"

    holdings = _read("holdings_analysis")
    if not holdings.empty and {"Ticker", "Weight"}.issubset(holdings.columns):
        w = holdings[["Ticker", "Weight"]].copy()
        w["Weight"] = pd.to_numeric(w["Weight"], errors="coerce").fillna(0.0)
        return w, "Holdings-analysis weights"

    return pd.DataFrame(), "Unavailable"


def _merge_saved(base_cfg: pd.DataFrame, saved: pd.DataFrame) -> pd.DataFrame:
    if saved is None or saved.empty or "Ticker" not in saved:
        return base_cfg
    saved = saved.copy()
    saved["Ticker"] = saved["Ticker"].astype(str).str.upper().str.strip()
    editable = ["AIExposure", "EvidenceConfidence", "MaturityConfidence", "AISkillConfidence"]
    cols = ["Ticker"] + [c for c in editable if c in saved]
    merged = base_cfg.merge(saved[cols], on="Ticker", how="left", suffixes=("", "_saved"))
    for col in editable:
        saved_col = f"{col}_saved"
        if saved_col in merged:
            merged[col] = pd.to_numeric(merged[saved_col], errors="coerce").combine_first(pd.to_numeric(merged[col], errors="coerce"))
            merged = merged.drop(columns=saved_col)
    return merged


def render() -> None:
    st.set_page_config(page_title="AI Optionality", layout="wide")
    st.title("AI Optionality & Uncertainty — Standalone Overlay")
    st.caption(
        "Research-derived AI sensitivity inputs · scenario-weighted optionality · uncertainty penalty · "
        "portfolio AI-factor exposure · no changes to existing ML models."
    )
    st.info(
        "This page is a separate decision-support overlay. It reads existing expected-return and portfolio outputs but does "
        "not retrain, overwrite or feed back into the existing machine-learning models or optimizer."
    )

    expected, expected_source = _load_expected_returns()
    weights, weights_source = _load_current_weights()
    if expected.empty or "Ticker" not in expected:
        st.error(
            "No usable security-level expected-return output exists yet. Run `python run_research.py` from "
            "`institutional_research`, then reload this page."
        )
        st.stop()

    expected = expected.copy()
    expected["Ticker"] = expected["Ticker"].astype(str).str.upper().str.strip()
    tickers = expected["Ticker"].dropna().drop_duplicates().tolist()

    core_candidates = {
        "Existing confidence-shrunk ML forecast": "BlendedMLExpectedReturn",
        "Existing classical forecast": "ClassicalExpectedReturn",
        "Historical prior": "HistoricalPrior",
    }
    core_candidates = {label: col for label, col in core_candidates.items() if col in expected.columns}
    if not core_candidates:
        st.error("Expected-return output does not contain a supported existing forecast column.")
        st.stop()

    suggestions = suggested_ai_config(ROOT, tickers)
    if suggestions.empty:
        suggestions = pd.DataFrame({
            "Ticker": tickers,
            "AIExposure": 0.0,
            "EvidenceConfidence": 0.20,
            "MaturityConfidence": 0.20,
            "AISkillConfidence": 0.25,
            "ResearchSource": "No research-derived suggestion",
            "ResearchEvidenceCount": 0,
        })

    base_cfg = suggestions.copy()
    if LOCAL_CONFIG.exists():
        try:
            base_cfg = _merge_saved(base_cfg, pd.read_csv(LOCAL_CONFIG))
        except Exception:
            pass

    health = st.columns(4)
    health[0].metric("Expected-return source", expected_source)
    health[1].metric("Weight source", weights_source)
    health[2].metric("Securities", len(tickers))
    evidence_names = int((pd.to_numeric(base_cfg.get("ResearchEvidenceCount"), errors="coerce").fillna(0) > 0).sum())
    health[3].metric("Names with dated AI evidence", evidence_names)

    left, middle, right = st.columns(3)
    with left:
        core_label = st.selectbox("Existing forecast used as anchor", list(core_candidates), index=0)
    with middle:
        uncertainty_aversion = st.slider(
            "Uncertainty aversion (λ)", 0.0, 1.5, 0.40, 0.05,
            help="Higher values penalize a wide AI scenario distribution more heavily.",
        )
    with right:
        st.metric("Existing ML changed?", "No")

    st.subheader("1. AI scenario distribution")
    st.caption(
        "Incremental return is the 12-month return contribution for a security with AI Exposure = 1.0. "
        "These are editable analyst scenarios, not ML predictions. Probabilities are normalized automatically."
    )
    scenario_display = DEFAULT_SCENARIOS.copy()
    scenario_display["ProbabilityPct"] = scenario_display["Probability"] * 100.0
    scenario_display["IncrementalReturnPct"] = scenario_display["IncrementalReturn"] * 100.0
    scenario_display = scenario_display[["Scenario", "ProbabilityPct", "IncrementalReturnPct"]]

    scenario_edit = st.data_editor(
        scenario_display,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Scenario": st.column_config.TextColumn("Scenario"),
            "ProbabilityPct": st.column_config.NumberColumn("Probability %", min_value=0.0, step=1.0, format="%.1f"),
            "IncrementalReturnPct": st.column_config.NumberColumn("Incremental return %", step=1.0, format="%.1f"),
        },
        key="ai_scenarios_v2",
    )
    scenario_engine = scenario_edit.rename(columns={"ProbabilityPct": "Probability", "IncrementalReturnPct": "IncrementalReturn"}).copy()
    scenario_engine["Probability"] = pd.to_numeric(scenario_engine["Probability"], errors="coerce") / 100.0
    scenario_engine["IncrementalReturn"] = pd.to_numeric(scenario_engine["IncrementalReturn"], errors="coerce") / 100.0
    try:
        scenario_engine = normalize_scenarios(scenario_engine)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    scenario_mean = float((scenario_engine["Probability"] * scenario_engine["IncrementalReturn"]).sum())
    scenario_sigma = float(np.sqrt((scenario_engine["Probability"] * (scenario_engine["IncrementalReturn"] - scenario_mean) ** 2).sum()))

    st.subheader("2. Security-level AI exposure and credibility")
    st.caption(
        "AI Exposure is an analyst sensitivity proxy, not a measured beta. Where dated AI KPI/source evidence exists, "
        "the page now starts from a conservative research-derived suggestion. You can edit every value."
    )

    uploaded = st.file_uploader("Optional: load a saved AI-overlay CSV", type=["csv"], key="ai_overlay_upload_v2")
    if uploaded is not None:
        try:
            base_cfg = _merge_saved(base_cfg, pd.read_csv(uploaded))
        except Exception as exc:
            st.warning(f"Could not load that configuration: {exc}")

    cfg = st.data_editor(
        base_cfg,
        use_container_width=True,
        hide_index=True,
        disabled=["Ticker", "ResearchSource", "ResearchEvidenceCount"],
        column_config={
            "Ticker": st.column_config.TextColumn("Ticker"),
            "AIExposure": st.column_config.NumberColumn(
                "AI exposure", min_value=0.0, max_value=2.0, step=0.05, format="%.2f",
                help="0 = no overlay; 1 = full scenario exposure. Research-derived values are only starting-point sensitivity proxies.",
            ),
            "EvidenceConfidence": st.column_config.NumberColumn("Evidence", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
            "MaturityConfidence": st.column_config.NumberColumn("AI maturity", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
            "AISkillConfidence": st.column_config.NumberColumn(
                "AI-overlay skill", min_value=0.0, max_value=1.0, step=0.05, format="%.2f",
                help="Keep low until this standalone AI overlay has enough realized out-of-sample evidence.",
            ),
            "ResearchEvidenceCount": st.column_config.NumberColumn("AI evidence rows", format="%d"),
            "ResearchSource": st.column_config.TextColumn("Input basis"),
        },
        key="ai_exposure_config_v2",
    )

    if pd.to_numeric(cfg["AIExposure"], errors="coerce").fillna(0).abs().sum() == 0:
        st.warning(
            "All AI exposures are currently zero, so the overlay will correctly equal the existing forecast. "
            "No dated AI-specific evidence was found for these names, or the saved configuration explicitly sets them to zero."
        )

    download_cols = ["Ticker", "AIExposure", "EvidenceConfidence", "MaturityConfidence", "AISkillConfidence"]
    st.download_button(
        "Download AI-overlay configuration CSV",
        data=cfg[download_cols].to_csv(index=False).encode("utf-8"),
        file_name="ai_optionality_inputs.csv",
        mime="text/csv",
    )

    core_column = core_candidates[core_label]
    overlay = build_overlay_table(
        expected,
        cfg,
        core_column=core_column,
        scenarios=scenario_engine,
        uncertainty_aversion=uncertainty_aversion,
    )
    if overlay.empty:
        st.warning("No valid expected-return rows were available for the AI overlay.")
        st.stop()

    if not weights.empty:
        weights = weights.copy()
        weights["Ticker"] = weights["Ticker"].astype(str).str.upper().str.strip()
        ai_scenarios = portfolio_ai_scenarios(weights, cfg, scenario_engine)
        factor_exposure = float(ai_scenarios["PortfolioAIFactorExposure"].iloc[0])
        weight_map = weights.set_index("Ticker")["Weight"]
        overlay["CurrentWeight"] = overlay["Ticker"].map(weight_map).fillna(0.0)
        weighted_ai_contribution = float((overlay["CurrentWeight"] * overlay["CertaintyEquivalentAI"]).sum())
    else:
        ai_scenarios = pd.DataFrame()
        factor_exposure = np.nan
        weighted_ai_contribution = np.nan

    st.subheader("3. AI-overlay result")
    cards = st.columns(5)
    cards[0].metric("Scenario mean", _pct(scenario_mean))
    cards[1].metric("Scenario uncertainty", _pct(scenario_sigma))
    cards[2].metric("Portfolio AI-factor exposure", f"{factor_exposure:.2f}" if np.isfinite(factor_exposure) else "—")
    cards[3].metric("Portfolio certainty-equivalent AI", _pct(weighted_ai_contribution) if np.isfinite(weighted_ai_contribution) else "—")
    cards[4].metric("Optimizer weights changed", "No")

    chart = overlay[["Ticker", "CoreExpectedReturn", "AIOverlayExpectedReturn"]].melt("Ticker", var_name="Forecast", value_name="ExpectedReturn")
    chart["Forecast"] = chart["Forecast"].replace({
        "CoreExpectedReturn": "Existing forecast",
        "AIOverlayExpectedReturn": "Existing + AI overlay",
    })
    fig = px.bar(chart, x="Ticker", y="ExpectedReturn", color="Forecast", barmode="group", title="Existing expected return vs standalone AI sensitivity")
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(legend_title=None)
    st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)
    with left:
        contrib = overlay.sort_values("CertaintyEquivalentAI")
        fig = px.bar(
            contrib, x="CertaintyEquivalentAI", y="Ticker", orientation="h",
            hover_data=["RawAIContribution", "AIUncertaintyPenalty", "AICredibility"],
            title="Certainty-equivalent AI contribution by security",
        )
        fig.update_xaxes(tickformat="+.1%")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        exposure_plot = cfg[["Ticker", "AIExposure"]].copy()
        if not weights.empty:
            exposure_plot = exposure_plot.merge(weights, on="Ticker", how="left")
            exposure_plot["Weight"] = exposure_plot["Weight"].fillna(0.0)
            exposure_plot["WeightedAIExposure"] = exposure_plot["Weight"] * exposure_plot["AIExposure"]
            fig = px.bar(exposure_plot, x="Ticker", y="WeightedAIExposure", title="Contribution to portfolio AI-factor exposure")
            fig.update_yaxes(tickformat=".1%")
        else:
            fig = px.bar(exposure_plot, x="Ticker", y="AIExposure", title="Analyst-defined AI exposure")
        st.plotly_chart(fig, use_container_width=True)

    if not ai_scenarios.empty:
        fig = px.bar(
            ai_scenarios, x="Scenario", y="PortfolioAIShock",
            title="Portfolio return sensitivity to each AI scenario — exposure diagnostic only",
            hover_data=["Probability", "IncrementalReturn", "PortfolioAIFactorExposure"],
        )
        fig.update_yaxes(tickformat="+.1%")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("4. Security drill-down")
    selected = st.selectbox("Security", overlay["Ticker"].tolist(), key="ai_security_v2")
    row = overlay[overlay["Ticker"] == selected].iloc[0]
    cols = st.columns(6)
    cols[0].metric("Existing forecast", _pct(row["CoreExpectedReturn"]))
    cols[1].metric("Raw AI mean", _pct(row["RawAIContribution"]))
    cols[2].metric("Uncertainty penalty", _pct(row["AIUncertaintyPenalty"]))
    cols[3].metric("AI credibility", _pct(row["AICredibility"], 0))
    cols[4].metric("Certainty-equivalent AI", _pct(row["CertaintyEquivalentAI"]))
    cols[5].metric("Overlay forecast", _pct(row["AIOverlayExpectedReturn"]))

    selected_exposure = float(row["AIExposure"])
    ticker_scen = scenario_engine.copy()
    ticker_scen["TickerImpact"] = ticker_scen["IncrementalReturn"] * selected_exposure
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=ticker_scen["Scenario"], y=ticker_scen["TickerImpact"],
        customdata=np.c_[ticker_scen["Probability"]],
        hovertemplate="%{x}<br>Return impact %{y:+.1%}<br>Probability %{customdata[0]:.0%}<extra></extra>",
    ))
    fig.update_layout(title=f"{selected}: AI scenario return distribution", yaxis_tickformat="+.1%", height=430)
    st.plotly_chart(fig, use_container_width=True)

    sens = sensitivity_grid(
        float(row["CoreExpectedReturn"]),
        scenario_engine,
        evidence_confidence=float(row["EvidenceConfidence"]),
        maturity_confidence=float(row["MaturityConfidence"]),
        ai_skill_confidence=float(row["AISkillConfidence"]),
    )
    pivot = sens.pivot(index="AIExposure", columns="UncertaintyAversion", values="AdjustedReturn")
    fig = px.imshow(
        pivot, text_auto=".1%", aspect="auto",
        labels={"x": "Uncertainty aversion", "y": "AI exposure", "color": "Overlay return"},
        title=f"{selected}: overlay-return sensitivity",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Audit table")
    audit = overlay.merge(
        cfg[["Ticker", "ResearchSource", "ResearchEvidenceCount"]],
        on="Ticker", how="left",
    )
    st.dataframe(
        audit,
        use_container_width=True,
        hide_index=True,
        column_config={
            "CoreExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
            "AIScenarioMean": st.column_config.NumberColumn(format="%.1%"),
            "AIScenarioSigma": st.column_config.NumberColumn(format="%.1%"),
            "RawAIContribution": st.column_config.NumberColumn(format="%.1%"),
            "AIUncertaintyPenalty": st.column_config.NumberColumn(format="%.1%"),
            "CertaintyEquivalentAI": st.column_config.NumberColumn(format="%.1%"),
            "AIOverlayExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
        },
    )
    st.caption(
        "Research-derived exposure is only a starting-point sensitivity proxy. It is intentionally not presented as a measured AI beta, "
        "and no value on this page is written back into the existing ML prediction journal or optimizer."
    )
