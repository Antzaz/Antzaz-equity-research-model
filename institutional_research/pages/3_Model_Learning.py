from __future__ import annotations

from pathlib import Path
import json
import sqlite3

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE = Path(__file__).resolve().parents[1]
ROOT = Path(__file__).resolve().parents[2]
DB = ROOT / "ml_data" / "ml_history.sqlite"
OUT = BASE / "outputs" / "latest"

st.set_page_config(page_title="Model Learning", layout="wide")
st.title("Model Learning & Governance")
st.caption(
    "Continual-learning control room · realized forecasts, baseline skill, drift, calibration, "
    "training evidence and the ML influence currently allowed into portfolio construction."
)


def pct(v, digits=1):
    try:
        if v is None or pd.isna(v): return "—"
        return f"{float(v):.{digits}%}"
    except Exception:
        return "—"


def num(v, digits=2):
    try:
        if v is None or pd.isna(v): return "—"
        return f"{float(v):.{digits}f}"
    except Exception:
        return "—"


def table_exists(con, name):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def parse_prediction(value):
    if value in (None, ""):
        return None
    try:
        obj = json.loads(value) if isinstance(value, str) else value
    except Exception:
        obj = value
    if isinstance(obj, dict):
        for key in ("prediction", "value", "expected_return", "forecast"):
            if key in obj:
                obj = obj[key]; break
    try:
        x = float(obj)
        return x if np.isfinite(x) else None
    except Exception:
        return None


def read_sql(query, params=()):
    try:
        with sqlite3.connect(DB) as con:
            return pd.read_sql_query(query, con, params=params)
    except Exception:
        return pd.DataFrame()


if not DB.exists():
    st.warning(
        "The local ML history database does not exist yet. Run a normal equity-research ML search "
        "first; the continual-learning schema and model registry are created automatically."
    )
    st.code("python research.py GOOGL", language="powershell")
    st.stop()

with sqlite3.connect(DB) as con:
    has_registry = table_exists(con, "model_registry")
    has_perf = table_exists(con, "model_performance")
    has_training = table_exists(con, "training_runs")
    has_predictions = table_exists(con, "predictions")

if not has_registry:
    st.info(
        "This database predates the continual-learning upgrade. Run any normal ML research search "
        "or execute `python -m machine_learning.continual_learning` once to migrate it."
    )
    st.stop()

registry = read_sql("SELECT * FROM model_registry ORDER BY model")
performance = read_sql("SELECT * FROM model_performance ORDER BY evaluated_at,model") if has_perf else pd.DataFrame()
training = read_sql("SELECT * FROM training_runs ORDER BY trained_at,model") if has_training else pd.DataFrame()
pred = read_sql(
    """SELECT id,symbol,model,as_of,horizon_days,prediction,confidence,model_version,
              realized_at,realized_value,error,baseline_value,baseline_name,target_type,created_at
       FROM predictions ORDER BY datetime(created_at),id"""
) if has_predictions else pd.DataFrame()

if not pred.empty:
    pred["Prediction"] = pred["prediction"].map(parse_prediction)
    for c in ("realized_value", "error", "baseline_value"):
        pred[c] = pd.to_numeric(pred[c], errors="coerce")
    pred["created_at"] = pd.to_datetime(pred["created_at"], errors="coerce", utc=True)
    pred["realized_at"] = pd.to_datetime(pred["realized_at"], errors="coerce", utc=True)
    pred["as_of_day"] = pd.to_datetime(pred["as_of"], errors="coerce", utc=True).dt.date
    # Same-day reruns are one forecast decision, not independent statistical observations.
    pred = pred.sort_values(["created_at", "id"]).drop_duplicates(
        ["model", "symbol", "as_of_day", "model_version"], keep="last"
    )

matured = pred[pred["realized_value"].notna() & pred["Prediction"].notna()].copy() if not pred.empty else pd.DataFrame()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Matured forecasts", int(len(matured)))
c2.metric("Champions", int((registry.get("status") == "CHAMPION").sum()) if not registry.empty else 0)
c3.metric("Challengers", int((registry.get("status") == "CHALLENGER").sum()) if not registry.empty else 0)
c4.metric("Demoted", int((registry.get("status") == "DEMOTED").sum()) if not registry.empty else 0)
exp_row = registry[registry["model"] == "Expected 12M Excess Return"] if not registry.empty else pd.DataFrame()
c5.metric("Expected-return ML influence", pct(exp_row.iloc[0].get("influence_multiplier")) if not exp_row.empty else "—")

st.info(
    "Governance rule: realized out-of-sample evidence controls portfolio influence. "
    "UNPROVEN models receive only a small fraction of their confidence-based weight; CHAMPION models can earn the full cap; "
    "DEMOTED models receive zero expected-return influence. Context-only models such as anomaly/regime classification continue to retrain but are not scored as numeric return forecasts."
)

st.subheader("Model registry")
if registry.empty:
    st.info("No registry rows yet.")
else:
    show = registry.copy()
    for c in ["mae", "baseline_mae", "skill_vs_baseline", "directional_accuracy", "information_coefficient", "calibration_score", "drift_ratio", "influence_multiplier"]:
        if c in show: show[c] = pd.to_numeric(show[c], errors="coerce")
    st.dataframe(
        show[[c for c in [
            "model", "status", "champion_version", "matured_predictions", "mae", "baseline_mae",
            "skill_vs_baseline", "directional_accuracy", "information_coefficient",
            "calibration_score", "drift_ratio", "influence_multiplier", "last_evaluated"
        ] if c in show]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "mae": st.column_config.NumberColumn("MAE", format="%.2%"),
            "baseline_mae": st.column_config.NumberColumn("Baseline MAE", format="%.2%"),
            "skill_vs_baseline": st.column_config.NumberColumn("Skill vs baseline", format="%.1%"),
            "directional_accuracy": st.column_config.NumberColumn("Direction", format="%.1%"),
            "information_coefficient": st.column_config.NumberColumn("IC", format="%.2f"),
            "calibration_score": st.column_config.NumberColumn("Confidence calibration", format="%.2f"),
            "drift_ratio": st.column_config.NumberColumn("Drift ratio", format="%.2f"),
            "influence_multiplier": st.column_config.ProgressColumn("Allowed influence", min_value=0.0, max_value=1.0, format="%.0%%"),
        },
    )

supervised_models = sorted(matured["model"].dropna().unique().tolist()) if not matured.empty else []
if supervised_models:
    selected = st.selectbox("Forecast model", supervised_models, index=0)
    view = matured[matured["model"] == selected].copy().sort_values("realized_at")

    st.subheader("Predicted vs realized")
    left, right = st.columns([1.25, 1])
    with left:
        fig = px.scatter(
            view, x="Prediction", y="realized_value", color="confidence", hover_name="symbol",
            hover_data=["as_of", "model_version", "baseline_value"],
            labels={"Prediction": "Predicted", "realized_value": "Realized"},
            title=f"{selected}: forecast calibration",
        )
        vals = pd.concat([view["Prediction"], view["realized_value"]]).dropna()
        if not vals.empty:
            lo, hi = float(vals.min()), float(vals.max())
            fig.add_trace(go.Scatter(x=[lo, hi], y=[lo, hi], mode="lines", name="Perfect forecast", line=dict(dash="dash")))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        abs_error = (view["realized_value"] - view["Prediction"]).abs()
        baseline_error = (view["realized_value"] - view["baseline_value"]).abs()
        m1, m2 = st.columns(2)
        m1.metric("Model MAE", pct(abs_error.mean()))
        m2.metric("Baseline MAE", pct(baseline_error.mean()) if baseline_error.notna().any() else "—")
        m3, m4 = st.columns(2)
        skill = 1 - abs_error.mean() / baseline_error.mean() if baseline_error.notna().any() and baseline_error.mean() > 0 else None
        direction = (np.sign(view["Prediction"]) == np.sign(view["realized_value"])).mean()
        m3.metric("Skill vs baseline", pct(skill))
        m4.metric("Directional accuracy", pct(direction))
        st.caption("Positive skill means the model's realized MAE is lower than the contemporaneous simple baseline.")

    st.subheader("Rolling forecast error and cumulative baseline edge")
    err = view[["realized_at", "Prediction", "realized_value", "baseline_value", "symbol"]].copy()
    err["ModelAbsError"] = (err["realized_value"] - err["Prediction"]).abs()
    err["BaselineAbsError"] = (err["realized_value"] - err["baseline_value"]).abs()
    window = min(10, max(3, len(err) // 3))
    err["ModelRollingMAE"] = err["ModelAbsError"].rolling(window, min_periods=2).mean()
    err["BaselineRollingMAE"] = err["BaselineAbsError"].rolling(window, min_periods=2).mean()
    err["CumulativeEdge"] = (err["BaselineAbsError"] - err["ModelAbsError"]).fillna(0).cumsum()
    l, r = st.columns(2)
    with l:
        rolling = err.melt("realized_at", value_vars=["ModelRollingMAE", "BaselineRollingMAE"], var_name="Series", value_name="MAE")
        rolling["Series"] = rolling["Series"].replace({"ModelRollingMAE": "Model", "BaselineRollingMAE": "Baseline"})
        st.plotly_chart(px.line(rolling, x="realized_at", y="MAE", color="Series", title=f"Rolling {window}-forecast MAE"), use_container_width=True)
    with r:
        st.plotly_chart(px.line(err, x="realized_at", y="CumulativeEdge", title="Cumulative error advantage vs baseline"), use_container_width=True)
        st.caption("Above zero means the model has cumulatively made less absolute forecast error than the baseline.")

    st.subheader("Confidence calibration")
    cal = view.assign(AbsError=(view["realized_value"] - view["Prediction"]).abs()).groupby("confidence", as_index=False).agg(
        MeanAbsoluteError=("AbsError", "mean"), Predictions=("AbsError", "size")
    )
    if not cal.empty:
        st.plotly_chart(px.bar(cal, x="confidence", y="MeanAbsoluteError", text="Predictions", title="Does higher confidence actually mean lower error?"), use_container_width=True)

    with st.expander("Matured forecast history"):
        st.dataframe(view[[c for c in ["symbol", "as_of", "Prediction", "realized_value", "error", "baseline_value", "confidence", "model_version", "realized_at"] if c in view]], use_container_width=True, hide_index=True)
else:
    st.warning(
        "No supervised forecasts have matured yet. This is normal for a newly installed 12-month feedback loop. "
        "The earnings-surprise model can mature sooner after the next reported earnings event."
    )

st.subheader("Training / walk-forward evidence")
if training.empty:
    st.info("Training diagnostics will appear after the next normal ML research run.")
else:
    training["trained_at"] = pd.to_datetime(training["trained_at"], errors="coerce", utc=True)
    latest_training = training.sort_values("trained_at").groupby("model", as_index=False).tail(1)
    st.dataframe(
        latest_training[[c for c in ["model", "model_version", "trained_at", "status", "confidence", "training_rows"] if c in latest_training]],
        use_container_width=True, hide_index=True,
    )

    selected_training_model = st.selectbox("Training diagnostics model", sorted(training["model"].dropna().unique()), key="training_model")
    tr = training[training["model"] == selected_training_model].sort_values("trained_at")
    if not tr.empty:
        counts = tr[["trained_at", "training_rows"]].copy(); counts["training_rows"] = pd.to_numeric(counts["training_rows"], errors="coerce")
        st.plotly_chart(px.line(counts, x="trained_at", y="training_rows", markers=True, title="Training sample growth"), use_container_width=True)
        latest = tr.iloc[-1]
        try:
            drivers = json.loads(latest.get("drivers_json") or "[]")
        except Exception:
            drivers = []
        if drivers:
            d = pd.DataFrame(drivers)
            numeric_cols = [c for c in d.columns if c != "feature" and pd.api.types.is_numeric_dtype(d[c])]
            if "feature" in d and numeric_cols:
                metric = numeric_cols[0]
                d["Magnitude"] = pd.to_numeric(d[metric], errors="coerce").abs()
                st.plotly_chart(px.bar(d.sort_values("Magnitude"), x="Magnitude", y="feature", orientation="h", title="Latest model driver attribution"), use_container_width=True)
                st.caption("Driver importance/coefficients describe model attribution, not economic causality.")

st.subheader("Portfolio learning link")
expected_path = OUT / "optimizer_expected_returns.csv"
if expected_path.exists():
    try:
        expected = pd.read_csv(expected_path)
    except Exception:
        expected = pd.DataFrame()
    if not expected.empty and "Ticker" in expected:
        cols = [c for c in [
            "Ticker", "ClassicalExpectedReturn", "MLTotalReturn", "PreLearningMLBlendWeight",
            "LearningInfluenceMultiplier", "MLBlendWeight", "BlendedMLExpectedReturn",
            "LearningStatus", "LearningMaturedPredictions", "LearningSkillVsBaseline",
        ] if c in expected]
        st.dataframe(expected[cols], use_container_width=True, hide_index=True,
                     column_config={
                         "ClassicalExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
                         "MLTotalReturn": st.column_config.NumberColumn(format="%.1%"),
                         "PreLearningMLBlendWeight": st.column_config.NumberColumn(format="%.1%"),
                         "LearningInfluenceMultiplier": st.column_config.NumberColumn(format="%.1%"),
                         "MLBlendWeight": st.column_config.NumberColumn(format="%.1%"),
                         "BlendedMLExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
                         "LearningSkillVsBaseline": st.column_config.NumberColumn(format="%.1%"),
                     })
        st.caption("This is the exact learned influence currently flowing into the portfolio optimizer. The optimizer itself remains deterministic and constrained.")
else:
    st.caption("Run `python institutional_research/run_research.py` after an ML research run to export the learned portfolio inputs here.")

st.markdown("---")
st.caption(
    "Continual learning is intentionally governed batch learning, not uncontrolled online trading. "
    "Models retrain on expanding point-in-time data; promotion/demotion depends on realized out-of-sample evidence."
)
