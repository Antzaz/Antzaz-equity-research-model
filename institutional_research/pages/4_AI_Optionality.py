from __future__ import annotations

"""Standalone AI optionality and uncertainty dashboard.

This page never retrains or mutates the existing machine-learning models.  It is an
analyst-controlled scenario overlay on top of already-produced expected returns.
"""

import io
import json
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src.ai_optionality import (
    DEFAULT_SCENARIOS,
    build_overlay_table,
    normalize_scenarios,
    portfolio_ai_scenarios,
    sensitivity_grid,
)


BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "outputs" / "latest"
LOCAL_CONFIG = BASE / "ai_optionality_inputs.csv"

st.set_page_config(page_title="AI Optionality", layout="wide")
st.title("AI Optionality & Uncertainty — Standalone Overlay")
st.caption(
    "Scenario-weighted AI optionality · uncertainty penalty · evidence/maturity/skill haircuts · "
    "portfolio AI-factor exposure"
)
st.info(
    "This is a separate research overlay. It reads the existing expected-return outputs but does "
    "not change, retrain, overwrite or feed back into any existing machine-learning model or "
    "portfolio optimizer."
)


def read(name: str) -> pd.DataFrame:
    path = OUT / f"{name}.csv"
    try:
        return pd.read_csv(path) if path.exists() else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def pct(value, digits=1) -> str:
    try:
        x = float(value)
        return f"{x:.{digits}%}" if np.isfinite(x) else "—"
    except Exception:
        return "—"


expected = read("optimizer_expected_returns")
weights_raw = read("optimizer_weights")

if expected.empty or "Ticker" not in expected:
    st.error(
        "No expected-return inputs were found. Run `python run_research.py` from "
        "`institutional_research` first, then reload this page."
    )
    st.stop()

core_candidates = {
    "Existing confidence-shrunk ML forecast": "BlendedMLExpectedReturn",
    "Existing classical forecast": "ClassicalExpectedReturn",
    "Historical prior": "HistoricalPrior",
}
core_candidates = {label: col for label, col in core_candidates.items() if col in expected.columns}
if not core_candidates:
    st.error("Expected-return output does not contain a supported existing forecast column.")
    st.stop()

left, middle, right = st.columns(3)
with left:
    core_label = st.selectbox("Existing forecast used as the anchor", list(core_candidates), index=0)
with middle:
    uncertainty_aversion = st.slider(
        "Uncertainty aversion (λ)", 0.0, 1.5, 0.40, 0.05,
        help="Higher values penalize a wide AI scenario distribution more heavily.",
    )
with right:
    st.metric("Existing ML changed?", "No")

st.subheader("1. AI scenario distribution")
st.caption(
    "Incremental return means the 12-month return contribution for a security with AI Exposure = 1.0. "
    "These are analyst assumptions, not ML predictions. Probabilities are normalized automatically."
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
    key="ai_scenarios",
)

scenario_engine = scenario_edit.rename(
    columns={"ProbabilityPct": "Probability", "IncrementalReturnPct": "IncrementalReturn"}
).copy()
scenario_engine["Probability"] = pd.to_numeric(scenario_engine["Probability"], errors="coerce") / 100.0
scenario_engine["IncrementalReturn"] = pd.to_numeric(scenario_engine["IncrementalReturn"], errors="coerce") / 100.0
try:
    scenario_engine = normalize_scenarios(scenario_engine)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

scenario_mean = float((scenario_engine["Probability"] * scenario_engine["IncrementalReturn"]).sum())
scenario_sigma = float(np.sqrt(
    (scenario_engine["Probability"] * (scenario_engine["IncrementalReturn"] - scenario_mean) ** 2).sum()
))

st.subheader("2. Security-level AI exposure and credibility")
st.caption(
    "AI Exposure is an analyst sensitivity multiplier, not a factual beta. 0 means no AI overlay; "
    "1 means full scenario exposure. Evidence, maturity and AI-specific skill are deliberately "
    "separate from the existing ML model's confidence."
)

base_cfg = pd.DataFrame({"Ticker": expected["Ticker"].astype(str).str.upper().str.strip().drop_duplicates()})
base_cfg["AIExposure"] = 0.0
base_cfg["EvidenceConfidence"] = 0.50
base_cfg["MaturityConfidence"] = 0.50
base_cfg["AISkillConfidence"] = 0.25

if LOCAL_CONFIG.exists():
    try:
        saved = pd.read_csv(LOCAL_CONFIG)
        if "Ticker" in saved:
            saved["Ticker"] = saved["Ticker"].astype(str).str.upper().str.strip()
            keep = [c for c in base_cfg.columns if c in saved.columns]
            base_cfg = base_cfg.drop(columns=[c for c in keep if c != "Ticker"]).merge(
                saved[keep], on="Ticker", how="left"
            )
            for col, default in {
                "AIExposure": 0.0,
                "EvidenceConfidence": 0.50,
                "MaturityConfidence": 0.50,
                "AISkillConfidence": 0.25,
            }.items():
                if col not in base_cfg:
                    base_cfg[col] = default
                base_cfg[col] = pd.to_numeric(base_cfg[col], errors="coerce").fillna(default)
    except Exception:
        pass

uploaded = st.file_uploader("Optional: load a saved AI-overlay CSV", type=["csv"])
if uploaded is not None:
    try:
        incoming = pd.read_csv(uploaded)
        if "Ticker" in incoming:
            incoming["Ticker"] = incoming["Ticker"].astype(str).str.upper().str.strip()
            cols = [c for c in base_cfg.columns if c in incoming.columns]
            base_cfg = base_cfg[["Ticker"]].merge(incoming[cols], on="Ticker", how="left")
            for col, default in {
                "AIExposure": 0.0,
                "EvidenceConfidence": 0.50,
                "MaturityConfidence": 0.50,
                "AISkillConfidence": 0.25,
            }.items():
                if col not in base_cfg:
                    base_cfg[col] = default
                base_cfg[col] = pd.to_numeric(base_cfg[col], errors="coerce").fillna(default)
    except Exception as exc:
        st.warning(f"Could not load that configuration: {exc}")

cfg = st.data_editor(
    base_cfg,
    use_container_width=True,
    hide_index=True,
    disabled=["Ticker"],
    column_config={
        "Ticker": st.column_config.TextColumn("Ticker"),
        "AIExposure": st.column_config.NumberColumn(
            "AI exposure", min_value=0.0, max_value=2.0, step=0.05, format="%.2f",
            help="0 = no overlay; 1 = full scenario exposure; >1 = deliberately high sensitivity.",
        ),
        "EvidenceConfidence": st.column_config.NumberColumn(
            "Evidence", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"
        ),
        "MaturityConfidence": st.column_config.NumberColumn(
            "AI maturity", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"
        ),
        "AISkillConfidence": st.column_config.NumberColumn(
            "AI-overlay skill", min_value=0.0, max_value=1.0, step=0.05, format="%.2f",
            help="Keep low until the separate AI overlay has enough realized out-of-sample evidence.",
        ),
    },
    key="ai_exposure_config",
)

st.download_button(
    "Download AI-overlay configuration CSV",
    data=cfg.to_csv(index=False).encode("utf-8"),
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
    st.warning("No valid existing expected-return rows were available for the AI overlay.")
    st.stop()

weights = pd.DataFrame()
if not weights_raw.empty and {"Ticker", "CurrentWeight"}.issubset(weights_raw.columns):
    weights = weights_raw[["Ticker", "CurrentWeight"]].drop_duplicates("Ticker").rename(
        columns={"CurrentWeight": "Weight"}
    )
    weights["Weight"] = pd.to_numeric(weights["Weight"], errors="coerce").fillna(0.0)

if not weights.empty:
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
cards[0].metric("Scenario mean", pct(scenario_mean))
cards[1].metric("Scenario uncertainty", pct(scenario_sigma))
cards[2].metric("Portfolio AI-factor exposure", f"{factor_exposure:.2f}" if np.isfinite(factor_exposure) else "—")
cards[3].metric("Portfolio certainty-equivalent AI", pct(weighted_ai_contribution) if np.isfinite(weighted_ai_contribution) else "—")
cards[4].metric("Optimizer weights changed", "No")

chart = overlay[["Ticker", "CoreExpectedReturn", "AIOverlayExpectedReturn"]].melt(
    "Ticker", var_name="Forecast", value_name="ExpectedReturn"
)
chart["Forecast"] = chart["Forecast"].replace({
    "CoreExpectedReturn": "Existing forecast",
    "AIOverlayExpectedReturn": "Existing + standalone AI overlay",
})
fig = px.bar(
    chart,
    x="Ticker",
    y="ExpectedReturn",
    color="Forecast",
    barmode="group",
    title="Existing expected return vs standalone AI-overlay sensitivity",
)
fig.update_yaxes(tickformat=".0%")
fig.update_layout(legend_title=None)
st.plotly_chart(fig, use_container_width=True)

left, right = st.columns(2)
with left:
    contrib = overlay.sort_values("CertaintyEquivalentAI")
    fig = px.bar(
        contrib,
        x="CertaintyEquivalentAI",
        y="Ticker",
        orientation="h",
        hover_data=["RawAIContribution", "AIUncertaintyPenalty", "AICredibility"],
        title="Certainty-equivalent AI contribution by security",
    )
    fig.update_xaxes(tickformat="+.1%")
    st.plotly_chart(fig, use_container_width=True)

with right:
    exposure_plot = cfg.copy()
    if not weights.empty:
        exposure_plot = exposure_plot.merge(weights, on="Ticker", how="left")
        exposure_plot["Weight"] = exposure_plot["Weight"].fillna(0.0)
        exposure_plot["WeightedAIExposure"] = exposure_plot["Weight"] * exposure_plot["AIExposure"]
        y_col = "WeightedAIExposure"
        title = "Contribution to portfolio AI-factor exposure"
    else:
        y_col = "AIExposure"
        title = "Analyst-defined AI exposure"
    fig = px.bar(exposure_plot, x="Ticker", y=y_col, title=title)
    if y_col == "WeightedAIExposure":
        fig.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig, use_container_width=True)

if not ai_scenarios.empty:
    fig = px.bar(
        ai_scenarios,
        x="Scenario",
        y="PortfolioAIShock",
        title="Portfolio return sensitivity to each AI scenario — exposure diagnostic only",
        hover_data=["Probability", "IncrementalReturn", "PortfolioAIFactorExposure"],
    )
    fig.update_yaxes(tickformat="+.1%")
    st.plotly_chart(fig, use_container_width=True)

st.subheader("4. Security drill-down")
selected = st.selectbox("Security", overlay["Ticker"].tolist())
row = overlay[overlay["Ticker"] == selected].iloc[0]

cols = st.columns(6)
cols[0].metric("Existing forecast", pct(row["CoreExpectedReturn"]))
cols[1].metric("Raw AI mean", pct(row["RawAIContribution"]))
cols[2].metric("Uncertainty penalty", pct(row["AIUncertaintyPenalty"]))
cols[3].metric("AI credibility", pct(row["AICredibility"], 0))
cols[4].metric("Certainty-equivalent AI", pct(row["CertaintyEquivalentAI"]))
cols[5].metric("Overlay forecast", pct(row["AIOverlayExpectedReturn"]))

selected_exposure = float(row["AIExposure"])
ticker_scen = scenario_engine.copy()
ticker_scen["TickerImpact"] = ticker_scen["IncrementalReturn"] * selected_exposure
fig = go.Figure()
fig.add_trace(go.Bar(
    x=ticker_scen["Scenario"],
    y=ticker_scen["TickerImpact"],
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
    pivot,
    text_auto=".1%",
    aspect="auto",
    labels={"x": "Uncertainty aversion", "y": "AI exposure", "color": "Overlay return"},
    title=f"{selected}: sensitivity of overlay return to exposure and uncertainty aversion",
)
st.plotly_chart(fig, use_container_width=True)

st.subheader("Audit table")
display_cols = [
    "Ticker", "CoreExpectedReturn", "AIExposure", "AIScenarioMean", "AIScenarioSigma",
    "EvidenceConfidence", "MaturityConfidence", "AISkillConfidence", "AICredibility",
    "RawAIContribution", "AIUncertaintyPenalty", "CertaintyEquivalentAI", "AIOverlayExpectedReturn",
]
st.dataframe(
    overlay[[c for c in display_cols if c in overlay]].style.format({
        "CoreExpectedReturn": "{:.1%}",
        "AIScenarioMean": "{:.1%}",
        "AIScenarioSigma": "{:.1%}",
        "EvidenceConfidence": "{:.0%}",
        "MaturityConfidence": "{:.0%}",
        "AISkillConfidence": "{:.0%}",
        "AICredibility": "{:.1%}",
        "RawAIContribution": "{:+.1%}",
        "AIUncertaintyPenalty": "{:.1%}",
        "CertaintyEquivalentAI": "{:+.1%}",
        "AIOverlayExpectedReturn": "{:.1%}",
    }, na_rep="—"),
    use_container_width=True,
    hide_index=True,
)

with st.expander("Methodology and separation from existing ML"):
    st.markdown(
        r"""
The standalone overlay uses:

$$
C_{AI}=C_{evidence}\times C_{maturity}\times C_{AI\ skill}
$$

$$
AI_{raw}=Exposure_{AI}\times E[AI\ scenario\ return]
$$

$$
Penalty_{AI}=\lambda\times |Exposure_{AI}|\times \sigma_{AI}
$$

$$
AI_{CE}=C_{AI}\times(AI_{raw}-Penalty_{AI})
$$

$$
E[R]_{overlay}=E[R]_{existing}+AI_{CE}
$$

**Governance**
- The existing forecast is read-only input to this page.
- Existing ML model code, model parameters, training data, prediction journal and optimizer are untouched.
- AI exposure and scenario probabilities are analyst-controlled assumptions.
- AI-specific skill starts deliberately low until this separate overlay builds genuine realized out-of-sample evidence.
- Portfolio AI-factor exposure is a diagnostic sensitivity, not an optimizer constraint.
- No trades are executed and no target weights are changed.
        """
    )
    st.json({
        "anchor_forecast": core_label,
        "anchor_column": core_column,
        "uncertainty_aversion": uncertainty_aversion,
        "scenario_probability_sum": float(scenario_engine["Probability"].sum()),
        "scenario_mean": scenario_mean,
        "scenario_sigma": scenario_sigma,
        "portfolio_ai_factor_exposure": None if not np.isfinite(factor_exposure) else factor_exposure,
    })
