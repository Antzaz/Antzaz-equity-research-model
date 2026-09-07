from __future__ import annotations

from pathlib import Path
import json
import sqlite3
import sys

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE = Path(__file__).resolve().parents[1]
ROOT = Path(__file__).resolve().parents[2]
DB = ROOT / "ml_data" / "ml_history.sqlite"
OUT = BASE / "outputs" / "latest"

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

st.set_page_config(page_title="Model Learning", layout="wide")
st.title("Model Learning & Governance")
st.caption(
    "Live 1M / 3M / 6M / 12M forecast feedback · baseline skill · drift · calibration · "
    "walk-forward evidence · governed portfolio influence."
)


def pct(v, digits=1):
    try:
        if v is None or pd.isna(v):
            return "—"
        return f"{float(v):.{digits}%}"
    except Exception:
        return "—"


def num(v, digits=2):
    try:
        if v is None or pd.isna(v):
            return "—"
        return f"{float(v):.{digits}f}"
    except Exception:
        return "—"


def parse_prediction(value):
    if value in (None, ""):
        return None
    try:
        obj = json.loads(value) if isinstance(value, str) else value
    except Exception:
        obj = value
    if isinstance(obj, dict):
        for key in ("prediction", "value", "expected_return", "forecast", "median_prediction"):
            if key in obj:
                obj = obj[key]
                break
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


HORIZON_ORDER = {
    "Expected 1M Excess Return": 1,
    "Expected 3M Excess Return": 2,
    "Expected 6M Excess Return": 3,
    "Expected 12M Excess Return": 4,
    "Consensus / Earnings Surprise": 5,
}
HORIZON_LABEL = {
    "Expected 1M Excess Return": "1M",
    "Expected 3M Excess Return": "3M",
    "Expected 6M Excess Return": "6M",
    "Expected 12M Excess Return": "12M",
    "Consensus / Earnings Surprise": "Next earnings",
}
SHORT_MODELS = {
    "Expected 1M Excess Return",
    "Expected 3M Excess Return",
    "Expected 6M Excess Return",
}


if not DB.exists():
    st.warning(
        "No local ML history database exists yet. Build/restore the ML history first; the page cannot "
        "invent forecast outcomes without that point-in-time database."
    )
    st.code(
        "cd C:\\Users\\Antza\\Documents\\Antzaz-equity-research-model\n"
        "python ml_history.py daily-refresh --universe sp500 --limit 500 --years 1 --deep-years 20 --deep-batch 25\n"
        "python -m machine_learning.learning_runner",
        language="powershell",
    )
    st.stop()

# Importing the package installs the governed learning extensions. Existing databases are
# migrated automatically instead of making the dashboard fail on an older schema.
try:
    import machine_learning  # noqa: F401
    from machine_learning.history_store import HistoryStore
    from machine_learning import continual_learning as cl

    store = HistoryStore(DB)
    cl.ensure_learning_schema(store)
except Exception as exc:
    st.error(f"Could not initialize the continual-learning database: {type(exc).__name__}: {exc}")
    st.stop()

left, right = st.columns([1, 4])
with left:
    if st.button("Refresh learning now", type="primary", use_container_width=True):
        try:
            with st.spinner("Updating short-horizon forecasts, maturities and governance..."):
                result = cl.run_continual_learning_cycle(store, "SPY")
            st.success(
                f"Learning refreshed. New short-horizon forecasts: {int(result.get('short_horizon_journaled', 0) or 0)} · "
                f"matured now: {int(result.get('matured', 0) or 0)}"
            )
            st.rerun()
        except Exception as exc:
            st.error(f"Learning refresh failed safely: {type(exc).__name__}: {exc}")
with right:
    st.info(
        "1M/3M/6M are separate excess-return models trained on their own forward targets. They provide faster feedback, "
        "but they are **research-only**. The portfolio optimizer continues to use only the governed 12M expected-return model."
    )

registry = read_sql("SELECT * FROM model_registry ORDER BY model")
performance = read_sql("SELECT * FROM model_performance ORDER BY evaluated_at,model")
training = read_sql("SELECT * FROM training_runs ORDER BY trained_at,model")
pred = read_sql(
    """SELECT id,symbol,model,as_of,horizon_days,prediction,confidence,model_version,
              realized_at,realized_value,error,baseline_value,baseline_name,target_type,created_at
       FROM predictions ORDER BY datetime(created_at),id"""
)

if not pred.empty:
    pred["Prediction"] = pred["prediction"].map(parse_prediction)
    for c in ("realized_value", "error", "baseline_value", "horizon_days"):
        pred[c] = pd.to_numeric(pred[c], errors="coerce")
    pred["created_at"] = pd.to_datetime(pred["created_at"], errors="coerce", utc=True)
    pred["realized_at"] = pd.to_datetime(pred["realized_at"], errors="coerce", utc=True)
    pred["as_of_ts"] = pd.to_datetime(pred["as_of"], errors="coerce", utc=True)
    pred["as_of_day"] = pred["as_of_ts"].dt.date
    pred = pred.sort_values(["created_at", "id"]).drop_duplicates(
        ["model", "symbol", "as_of_day", "model_version"], keep="last"
    )

matured = pred[pred["realized_value"].notna() & pred["Prediction"].notna()].copy() if not pred.empty else pd.DataFrame()
pending = pred[pred["realized_value"].isna() & pred["Prediction"].notna()].copy() if not pred.empty else pd.DataFrame()

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Live matured", int(len(matured)))
c2.metric("1M/3M/6M matured", int(matured["model"].isin(SHORT_MODELS).sum()) if not matured.empty else 0)
c3.metric("Pending forecasts", int(len(pending)))
c4.metric("Champions", int((registry.get("status") == "CHAMPION").sum()) if not registry.empty else 0)
c5.metric("Demoted", int((registry.get("status") == "DEMOTED").sum()) if not registry.empty else 0)
exp_row = registry[registry["model"] == "Expected 12M Excess Return"] if not registry.empty else pd.DataFrame()
c6.metric("12M portfolio influence", pct(exp_row.iloc[0].get("influence_multiplier")) if not exp_row.empty else "—")

st.subheader("Learning horizon coverage")
coverage_rows = []
now = pd.Timestamp.now(tz="UTC")
models_for_coverage = [m for m in HORIZON_ORDER if (not pred.empty and m in set(pred["model"].dropna())) or (not registry.empty and m in set(registry["model"].dropna()))]
for model in sorted(models_for_coverage, key=lambda x: HORIZON_ORDER.get(x, 99)):
    p = pred[pred["model"] == model].copy() if not pred.empty else pd.DataFrame()
    live = p[p["realized_value"].notna()] if not p.empty else pd.DataFrame()
    waiting = p[p["realized_value"].isna()] if not p.empty else pd.DataFrame()
    next_due = None
    if not waiting.empty and waiting["horizon_days"].notna().any():
        due = waiting["as_of_ts"] + pd.to_timedelta(waiting["horizon_days"].fillna(0), unit="D")
        future_due = due[due >= now]
        if not future_due.empty:
            next_due = future_due.min().date().isoformat()
    reg = registry[registry["model"] == model] if not registry.empty else pd.DataFrame()
    coverage_rows.append({
        "Target": HORIZON_LABEL.get(model, model),
        "Model": model,
        "Matured": int(len(live)),
        "Pending": int(len(waiting)),
        "Next expected maturity": next_due,
        "Status": reg.iloc[0].get("status") if not reg.empty else "UNPROVEN",
        "Influence": reg.iloc[0].get("influence_multiplier") if not reg.empty else None,
        "Portfolio use": "12M governed input" if model == "Expected 12M Excess Return" else "Research only",
    })
coverage = pd.DataFrame(coverage_rows)
if coverage.empty:
    st.info("No horizon forecasts have been journaled yet. Click **Refresh learning now** after price/features history exists.")
else:
    st.dataframe(
        coverage,
        use_container_width=True,
        hide_index=True,
        column_config={"Influence": st.column_config.NumberColumn(format="%.0%")},
    )

st.caption(
    "A newly installed 1M model still needs roughly one month before its first *live* forecast can mature. "
    "Until then, its purged walk-forward evidence below gives an immediate historical out-of-sample diagnostic."
)

st.subheader("Model registry")
if registry.empty:
    st.info("No registry rows yet.")
else:
    show = registry.copy()
    show["_order"] = show["model"].map(HORIZON_ORDER).fillna(99)
    show = show.sort_values(["_order", "model"]).drop(columns="_order")
    for c in ["mae", "baseline_mae", "skill_vs_baseline", "directional_accuracy", "information_coefficient", "calibration_score", "drift_ratio", "influence_multiplier"]:
        if c in show:
            show[c] = pd.to_numeric(show[c], errors="coerce")
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
            "influence_multiplier": st.column_config.ProgressColumn("Governance score", min_value=0.0, max_value=1.0, format="%.0%%"),
        },
    )

all_models = set()
if not pred.empty:
    all_models.update(pred["model"].dropna().astype(str))
if not training.empty:
    all_models.update(training["model"].dropna().astype(str))
forecast_models = [m for m in all_models if m in HORIZON_ORDER]
forecast_models = sorted(forecast_models, key=lambda x: HORIZON_ORDER.get(x, 99))

if forecast_models:
    selected = st.selectbox(
        "Forecast target",
        forecast_models,
        format_func=lambda m: f"{HORIZON_LABEL.get(m, m)} — {m}",
    )
    view = matured[matured["model"] == selected].copy().sort_values("realized_at") if not matured.empty else pd.DataFrame()

    st.subheader("Live predicted vs realized")
    if view.empty:
        selected_pending = pending[pending["model"] == selected] if not pending.empty else pd.DataFrame()
        st.info(
            f"No live {HORIZON_LABEL.get(selected, selected)} forecasts have matured yet. "
            f"Pending live forecasts: {len(selected_pending)}. Use the walk-forward section below for evidence available now."
        )
    else:
        left, right = st.columns([1.25, 1])
        with left:
            fig = px.scatter(
                view, x="Prediction", y="realized_value", color="confidence", hover_name="symbol",
                hover_data=["as_of", "model_version", "baseline_value"],
                labels={"Prediction": "Predicted excess return", "realized_value": "Realized excess return"},
                title=f"{HORIZON_LABEL.get(selected, selected)} live forecast calibration",
            )
            vals = pd.concat([view["Prediction"], view["realized_value"]]).dropna()
            if not vals.empty:
                lo, hi = float(vals.min()), float(vals.max())
                fig.add_trace(go.Scatter(x=[lo, hi], y=[lo, hi], mode="lines", name="Perfect forecast", line=dict(dash="dash")))
            fig.update_xaxes(tickformat="+.1%")
            fig.update_yaxes(tickformat="+.1%")
            st.plotly_chart(fig, use_container_width=True)
        with right:
            abs_error = (view["realized_value"] - view["Prediction"]).abs()
            baseline_error = (view["realized_value"] - view["baseline_value"]).abs()
            m1, m2 = st.columns(2)
            m1.metric("Model MAE", pct(abs_error.mean()))
            m2.metric("Baseline MAE", pct(baseline_error.mean()) if baseline_error.notna().any() else "—")
            skill = 1 - abs_error.mean() / baseline_error.mean() if baseline_error.notna().any() and baseline_error.mean() > 0 else None
            direction = (np.sign(view["Prediction"]) == np.sign(view["realized_value"])).mean()
            m3, m4 = st.columns(2)
            m3.metric("Skill vs baseline", pct(skill))
            m4.metric("Directional accuracy", pct(direction))
            st.caption("Positive skill means lower realized absolute error than the contemporaneous trailing-return baseline.")

        if len(view) >= 3:
            st.subheader("Live error path")
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

        with st.expander("Live forecast history"):
            st.dataframe(
                view[[c for c in ["symbol", "as_of", "Prediction", "realized_value", "error", "baseline_value", "confidence", "model_version", "realized_at"] if c in view]],
                use_container_width=True,
                hide_index=True,
            )

st.subheader("Training / purged walk-forward evidence")
if training.empty:
    st.info("Training diagnostics will appear after the next normal ML/learning run.")
else:
    training["trained_at"] = pd.to_datetime(training["trained_at"], errors="coerce", utc=True)
    latest_training = training.sort_values("trained_at").groupby("model", as_index=False).tail(1).copy()
    latest_training["_order"] = latest_training["model"].map(HORIZON_ORDER).fillna(99)
    latest_training = latest_training.sort_values(["_order", "model"]).drop(columns="_order")

    evidence_rows = []
    for _, row in latest_training.iterrows():
        try:
            metrics = json.loads(row.get("validation_json") or "{}")
        except Exception:
            metrics = {}
        hgb = metrics.get("hgb_walk_forward") or {}
        if isinstance(hgb, dict) and "metrics" in hgb:
            hgb = hgb.get("metrics") or {}
        elastic = metrics.get("elastic_walk_forward") or {}
        if isinstance(elastic, dict) and "metrics" in elastic:
            elastic = elastic.get("metrics") or {}
        primary = hgb if hgb else metrics.get("walk_forward") or {}
        evidence_rows.append({
            "Model": row.get("model"),
            "Target": HORIZON_LABEL.get(row.get("model"), row.get("model")),
            "Training rows": pd.to_numeric(row.get("training_rows"), errors="coerce"),
            "Status": row.get("status"),
            "Confidence": row.get("confidence"),
            "WF observations": primary.get("n") if isinstance(primary, dict) else None,
            "WF MAE": primary.get("mae") if isinstance(primary, dict) else None,
            "WF baseline MAE": primary.get("baseline_mae") if isinstance(primary, dict) else None,
            "WF skill vs baseline": primary.get("mae_improvement_vs_baseline") if isinstance(primary, dict) else None,
            "WF direction": primary.get("directional_accuracy") if isinstance(primary, dict) else None,
            "Trained at": row.get("trained_at"),
        })
    evidence = pd.DataFrame(evidence_rows)
    st.dataframe(
        evidence,
        use_container_width=True,
        hide_index=True,
        column_config={
            "WF MAE": st.column_config.NumberColumn(format="%.2%"),
            "WF baseline MAE": st.column_config.NumberColumn(format="%.2%"),
            "WF skill vs baseline": st.column_config.NumberColumn(format="%.1%"),
            "WF direction": st.column_config.NumberColumn(format="%.1%"),
        },
    )
    st.caption(
        "Walk-forward tests are purged by target horizon: a historical row can enter training only after its future return target was actually observable. "
        "This is historical OOS evidence, not the same thing as a live track record."
    )

    models_with_training = sorted(training["model"].dropna().unique(), key=lambda x: HORIZON_ORDER.get(x, 99))
    selected_training_model = st.selectbox("Training sample history", models_with_training, key="training_model")
    tr = training[training["model"] == selected_training_model].sort_values("trained_at")
    if not tr.empty:
        counts = tr[["trained_at", "training_rows"]].copy()
        counts["training_rows"] = pd.to_numeric(counts["training_rows"], errors="coerce")
        st.plotly_chart(px.line(counts, x="trained_at", y="training_rows", markers=True, title="Training sample growth"), use_container_width=True)

st.subheader("Portfolio learning link — 12M only")
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
        st.dataframe(
            expected[cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "ClassicalExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
                "MLTotalReturn": st.column_config.NumberColumn(format="%.1%"),
                "PreLearningMLBlendWeight": st.column_config.NumberColumn(format="%.1%"),
                "LearningInfluenceMultiplier": st.column_config.NumberColumn(format="%.1%"),
                "MLBlendWeight": st.column_config.NumberColumn(format="%.1%"),
                "BlendedMLExpectedReturn": st.column_config.NumberColumn(format="%.1%"),
                "LearningSkillVsBaseline": st.column_config.NumberColumn(format="%.1%"),
            },
        )
        st.caption(
            "This table is intentionally 12M-only. The 1M/3M/6M models accelerate research feedback but cannot silently change optimizer weights."
        )
else:
    st.caption("Run `python run_research.py` from institutional_research to export the governed 12M portfolio inputs here.")

st.markdown("---")
st.caption(
    "Continual learning is governed batch learning, not uncontrolled online trading. Shorter targets improve feedback speed; "
    "they do not make overlapping market observations independent, so live evidence still needs time and breadth before being trusted."
)
