from __future__ import annotations

from pathlib import Path
import json
import sys

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE = Path(__file__).resolve().parents[1]
ROOT = Path(__file__).resolve().parents[2]
DB = ROOT / "ml_data" / "ml_history.sqlite"

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from institutional_research.src.ml_research_dashboard import (  # noqa: E402
    RETURN_MODEL_LABELS,
    RETURN_MODEL_ORDER,
    governance_explanation,
    latest_data_dates,
    latest_return_forecasts,
    latest_state,
    load_dashboard_tables,
    forecast_matrix,
    prepare_experiments,
    prepare_predictions,
    prepare_registry,
    return_forecast_explanation,
    risk_explanation,
    risk_forecast_frame,
    risk_matrix,
    table_counts,
    training_validation_table,
)


st.set_page_config(page_title="ML Research Lab", page_icon="🔬", layout="wide")
st.title("ML Research Lab")
st.caption(
    "Current 1D / 1W / 1M / 3M / 6M / 12M ML forecasts, forward-risk estimates, "
    "live model skill, and autonomous challenger experiments — explained in one place."
)


RETURN_ORDER = {model: i for i, model in enumerate(RETURN_MODEL_ORDER)}
HORIZON_ORDER = {RETURN_MODEL_LABELS[m]: i for i, m in enumerate(RETURN_MODEL_ORDER)}


def pct(value, digits=1, signed=False):
    try:
        if value is None or pd.isna(value):
            return "—"
        f = float(value)
        return f"{f:+.{digits}%}" if signed else f"{f:.{digits}%}"
    except Exception:
        return "—"


def num(value, digits=2):
    try:
        if value is None or pd.isna(value):
            return "—"
        return f"{float(value):.{digits}f}"
    except Exception:
        return "—"


def date_text(value):
    try:
        if value is None or pd.isna(value):
            return "—"
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    except Exception:
        return str(value or "—")


@st.cache_data(ttl=30, show_spinner=False)
def load_data(db_path: str):
    tables = load_dashboard_tables(Path(db_path))
    pred = prepare_predictions(tables.get("predictions", pd.DataFrame()))
    latest = latest_return_forecasts(pred)
    registry = prepare_registry(tables.get("registry", pd.DataFrame()))
    experiments = prepare_experiments(tables.get("experiments", pd.DataFrame()))
    risk = risk_forecast_frame(tables.get("provider_state", pd.DataFrame()))
    return tables, pred, latest, registry, experiments, risk


def clear_and_reload():
    st.cache_data.clear()
    st.rerun()


if not DB.exists():
    st.warning(
        "The ML database does not exist on this machine yet. The dashboard will never fabricate forecasts; "
        "initialize or restore the point-in-time history first."
    )
    st.code(
        "cd C:\\Users\\Antza\\Documents\\Antzaz-equity-research-model\n"
        "python ml_history.py daily-refresh --universe sp500 --limit 500 --years 1 --deep-years 20 --deep-batch 25\n"
        "python -m machine_learning.learning_runner\n"
        "python -m streamlit run institutional_research/app.py",
        language="powershell",
    )
    st.stop()

try:
    import machine_learning  # noqa: F401,E402 - installs governed learning hooks
    from machine_learning.history_store import HistoryStore  # noqa: E402
    from machine_learning import continual_learning as cl  # noqa: E402
    from machine_learning.research_loop import run_research_loop  # noqa: E402

    store = HistoryStore(DB)
    cl.ensure_learning_schema(store)
except Exception as exc:
    st.error(f"Could not initialize the ML database safely: {type(exc).__name__}: {exc}")
    st.stop()


tables, pred, latest, registry, experiments, risk = load_data(str(DB))
provider_state = tables.get("provider_state", pd.DataFrame())
training = tables.get("training", pd.DataFrame())

# ---------------------------------------------------------------------------
# Controls
# ---------------------------------------------------------------------------
with st.sidebar:
    st.subheader("Research controls")
    st.caption("These actions update research evidence only. They do not execute trades or change DCF assumptions.")

    if st.button("Refresh forecasts & learning", type="primary", use_container_width=True):
        try:
            with st.spinner("Generating forecasts, maturing old predictions, updating governance and risk models..."):
                result = cl.run_continual_learning_cycle(store, "SPY")
            st.success(
                "Learning cycle completed. "
                f"New short-horizon forecasts: {int(result.get('short_horizon_journaled', 0) or 0)} · "
                f"matured now: {int(result.get('matured', 0) or 0)}"
            )
            clear_and_reload()
        except Exception as exc:
            st.error(f"Learning cycle failed safely: {type(exc).__name__}: {exc}")

    if st.button("Run challenger research now", use_container_width=True):
        try:
            with st.spinner("Running bounded walk-forward challenger experiments..."):
                result = run_research_loop(store, "SPY", force=True, max_experiments=12)
            st.success(
                f"Research loop completed: {int(result.get('experiments', 0) or 0)} experiment(s), "
                f"{len(result.get('promotion_candidates', []) or [])} promotion candidate(s)."
            )
            clear_and_reload()
        except Exception as exc:
            st.error(f"Research loop failed safely: {type(exc).__name__}: {exc}")

    st.divider()
    st.caption(f"Database: `{DB}`")
    st.caption("Benchmark for excess-return forecasts: **SPY**")


all_symbols = sorted(set(latest.get("symbol", pd.Series(dtype=str)).dropna().astype(str))) if not latest.empty else []
if not risk.empty:
    all_symbols = sorted(set(all_symbols) | set(risk["symbol"].dropna().astype(str)))

selected_symbol = st.selectbox(
    "Company / symbol",
    all_symbols if all_symbols else ["No forecasts yet"],
    index=0,
    help="Select a company to see its latest multi-horizon return and risk forecasts.",
)
valid_symbol = selected_symbol if selected_symbol != "No forecasts yet" else None


# ---------------------------------------------------------------------------
# Header metrics
# ---------------------------------------------------------------------------
matured = pred[pred.get("realized_value", pd.Series(index=pred.index, dtype=float)).notna()] if not pred.empty else pd.DataFrame()
pending = pred[pred.get("realized_value", pd.Series(index=pred.index, dtype=float)).isna()] if not pred.empty else pd.DataFrame()
champions = int((registry.get("status", pd.Series(dtype=str)) == "CHAMPION").sum()) if not registry.empty else 0
promotion_candidates = int(experiments.get("promotion_candidate", pd.Series(dtype=bool)).fillna(False).sum()) if not experiments.empty else 0
latest_prediction_date = latest["as_of"].max() if not latest.empty else None

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Latest forecast date", date_text(latest_prediction_date))
m2.metric("Symbols forecast", int(latest["symbol"].nunique()) if not latest.empty else 0)
m3.metric("Pending live forecasts", int(len(pending)))
m4.metric("Matured live forecasts", int(len(matured)))
m5.metric("Champion models", champions)
m6.metric("Promotion candidates", promotion_candidates)


# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------
tab_overview, tab_forecasts, tab_risk, tab_skill, tab_agents, tab_data = st.tabs(
    ["Overview", "Return Forecasts", "Forward Risk", "Model Skill", "Agent Experiments", "Data Health"]
)

with tab_overview:
    st.subheader("Current research snapshot")
    if valid_symbol is None:
        st.info("No current return or risk forecasts have been journaled yet. Use **Refresh forecasts & learning** once history exists.")
    else:
        sym_latest = latest[latest["symbol"] == valid_symbol].copy() if not latest.empty else pd.DataFrame()
        sym_risk = risk[risk["symbol"] == valid_symbol].copy() if not risk.empty else pd.DataFrame()

        if not sym_latest.empty:
            cards = st.columns(max(1, len(RETURN_MODEL_ORDER)))
            for i, model in enumerate(RETURN_MODEL_ORDER):
                row = sym_latest[sym_latest["model"] == model]
                value = row.iloc[-1]["prediction_num"] if not row.empty else None
                confidence = str(row.iloc[-1].get("confidence") or "—") if not row.empty else "—"
                cards[i].metric(RETURN_MODEL_LABELS[model], pct(value, 2, signed=True), help=f"Confidence: {confidence}")
            st.caption("All return figures above are expected **excess returns versus SPY**, not expected total stock returns.")

            chart = sym_latest.sort_values("horizon_order")
            fig = px.bar(
                chart,
                x="horizon",
                y="prediction_num",
                hover_data=["confidence", "as_of", "model_version"],
                labels={"horizon": "Forecast horizon", "prediction_num": "Expected excess return"},
                title=f"{valid_symbol}: current expected excess return by horizon",
            )
            fig.update_yaxes(tickformat="+.1%", zeroline=True)
            st.plotly_chart(fig, use_container_width=True)

            short_row = chart.iloc[0]
            st.info(return_forecast_explanation(str(short_row["horizon"]), float(short_row["prediction_num"])))
        else:
            st.info(f"No return forecast is currently stored for {valid_symbol}.")

        vol = None
        dd = None
        if not sym_risk.empty:
            v = sym_risk[sym_risk["model"] == "Expected 1W Realized Volatility"]
            d = sym_risk[sym_risk["model"] == "Expected 1M Forward Drawdown"]
            vol = float(v.iloc[-1]["value"]) if not v.empty else None
            dd = float(d.iloc[-1]["value"]) if not d.empty else None
            r1, r2 = st.columns(2)
            r1.metric("Expected 1W realized volatility", pct(vol))
            r2.metric("Expected 1M forward drawdown", pct(dd))
            st.warning(risk_explanation(vol, dd))

        st.subheader("How much should I trust these models?")
        status_rows = []
        for model in RETURN_MODEL_ORDER:
            row = registry[registry["model"] == model] if not registry.empty else pd.DataFrame()
            if row.empty:
                continue
            r = row.iloc[-1]
            status_rows.append({
                "Horizon": RETURN_MODEL_LABELS[model],
                "Status": r.get("status"),
                "Matured": r.get("matured_predictions"),
                "Skill vs baseline": r.get("skill_vs_baseline"),
                "Direction": r.get("directional_accuracy"),
                "IC": r.get("information_coefficient"),
                "Drift": r.get("drift_ratio"),
            })
        if status_rows:
            st.dataframe(
                pd.DataFrame(status_rows),
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Skill vs baseline": st.column_config.NumberColumn(format="%.1%"),
                    "Direction": st.column_config.NumberColumn(format="%.1%"),
                    "IC": st.column_config.NumberColumn(format="%.2f"),
                    "Drift": st.column_config.NumberColumn(format="%.2f"),
                },
            )
        st.caption(
            "A good current forecast can still be statistically unproven. The governance status is based on matured live forecasts, "
            "not only historical backtests."
        )


# ---------------------------------------------------------------------------
# Return forecasts
# ---------------------------------------------------------------------------
with tab_forecasts:
    st.subheader("Latest multi-horizon return forecasts")
    matrix = forecast_matrix(latest)
    if matrix.empty:
        st.info("No current return forecasts are stored yet.")
    else:
        forecast_columns = {c: st.column_config.NumberColumn(c, format="%+.2f%%") for c in matrix.columns if c != "symbol"}
        # Streamlit's percent format expects the stored decimal, so use a styled display frame for clarity.
        display = matrix.copy()
        for col in [c for c in display.columns if c != "symbol"]:
            display[col] = display[col].map(lambda x: f"{x:+.2%}" if pd.notna(x) else "—")
        st.dataframe(display, use_container_width=True, hide_index=True)
        st.caption("Positive = expected to outperform SPY over the stated horizon; negative = expected to underperform SPY.")

    if valid_symbol and not latest.empty:
        sym = latest[latest["symbol"] == valid_symbol].copy().sort_values("horizon_order")
        if not sym.empty:
            st.subheader(f"{valid_symbol} forecast detail")
            horizon = st.selectbox(
                "Forecast horizon",
                list(sym["horizon"]),
                key="forecast_horizon_selector",
            )
            selected = sym[sym["horizon"] == horizon].iloc[-1]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Expected excess return", pct(selected["prediction_num"], 2, signed=True))
            c2.metric("Confidence", str(selected.get("confidence") or "—"))
            c3.metric("As of", date_text(selected.get("as_of")))
            c4.metric("Model version", str(selected.get("model_version") or "—"))
            st.info(return_forecast_explanation(horizon, float(selected["prediction_num"])))

            model_name = str(selected["model"])
            history = pred[(pred["symbol"] == valid_symbol) & (pred["model"] == model_name) & pred["prediction_num"].notna()].copy()
            history = history.sort_values("as_of")
            if not history.empty:
                chart_data = history[["as_of", "prediction_num", "realized_value", "baseline_value"]].copy()
                chart_data = chart_data.rename(columns={
                    "prediction_num": "Model forecast",
                    "realized_value": "Realized excess return",
                    "baseline_value": "Naive baseline",
                })
                long = chart_data.melt("as_of", var_name="Series", value_name="Return").dropna(subset=["Return"])
                if not long.empty:
                    fig = px.line(
                        long,
                        x="as_of",
                        y="Return",
                        color="Series",
                        markers=True,
                        title=f"{valid_symbol} {horizon} forecast history",
                    )
                    fig.update_yaxes(tickformat="+.1%")
                    st.plotly_chart(fig, use_container_width=True)

            st.subheader("Historical walk-forward evidence for this model")
            validation = training_validation_table(training, model_name)
            if validation.empty:
                st.info("No training-run validation summary is stored for this model yet.")
            else:
                st.dataframe(
                    validation,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "mae": st.column_config.NumberColumn("MAE", format="%.2%"),
                        "baseline_mae": st.column_config.NumberColumn("Baseline MAE", format="%.2%"),
                        "mae_improvement_vs_baseline": st.column_config.NumberColumn("MAE edge", format="%.1%"),
                        "directional_accuracy": st.column_config.NumberColumn("Direction", format="%.1%"),
                        "directional_accuracy_edge_vs_baseline": st.column_config.NumberColumn("Direction edge", format="%+.1%"),
                        "r2": st.column_config.NumberColumn("OOS R²", format="%.2f"),
                    },
                )
                st.caption(
                    "Walk-forward evidence is historical out-of-sample simulation with purged forward targets. Live matured forecasts in the Model Skill tab are the stronger real-world check."
                )


# ---------------------------------------------------------------------------
# Risk forecasts
# ---------------------------------------------------------------------------
with tab_risk:
    st.subheader("Forward risk estimates")
    rmatrix = risk_matrix(risk)
    if rmatrix.empty:
        st.info("No forward-risk state is stored yet. Run **Refresh forecasts & learning** after sufficient price history exists.")
    else:
        display = rmatrix.copy()
        if "1W Volatility" in display:
            display["1W Volatility"] = display["1W Volatility"].map(lambda x: f"{x:.1%}" if pd.notna(x) else "—")
        if "1M Drawdown" in display:
            display["1M Drawdown"] = display["1M Drawdown"].map(lambda x: f"{x:.1%}" if pd.notna(x) else "—")
        st.dataframe(display, use_container_width=True, hide_index=True)

    if valid_symbol and not risk.empty:
        sym = risk[risk["symbol"] == valid_symbol].copy()
        vol_row = sym[sym["model"] == "Expected 1W Realized Volatility"]
        dd_row = sym[sym["model"] == "Expected 1M Forward Drawdown"]
        vol = float(vol_row.iloc[-1]["value"]) if not vol_row.empty else None
        dd = float(dd_row.iloc[-1]["value"]) if not dd_row.empty else None
        c1, c2, c3 = st.columns(3)
        c1.metric("1W annualized realized volatility", pct(vol))
        c2.metric("1M forward drawdown", pct(dd))
        training_rows = None
        if not vol_row.empty:
            training_rows = vol_row.iloc[-1].get("training_rows")
        elif not dd_row.empty:
            training_rows = dd_row.iloc[-1].get("training_rows")
        c3.metric("Risk-model training rows", int(training_rows) if pd.notna(training_rows) else "—")
        st.info(risk_explanation(vol, dd))

        if vol is not None or dd is not None:
            risk_chart = pd.DataFrame([
                {"Metric": "1W annualized volatility", "Value": vol},
                {"Metric": "1M forward drawdown", "Value": dd},
            ]).dropna()
            fig = px.bar(risk_chart, x="Metric", y="Value", title=f"{valid_symbol}: forward-risk context")
            fig.update_yaxes(tickformat=".1%", zeroline=True)
            st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "The volatility model forecasts realized variability, while the drawdown model forecasts the worst decline from today's price over the forward window. "
        "Neither model automatically changes the portfolio optimizer."
    )


# ---------------------------------------------------------------------------
# Model skill and governance
# ---------------------------------------------------------------------------
with tab_skill:
    st.subheader("Live champion / challenger governance")
    if registry.empty:
        st.info("No model registry rows are available yet.")
    else:
        columns = [c for c in [
            "model", "status", "champion_version", "matured_predictions", "mae", "baseline_mae",
            "skill_vs_baseline", "directional_accuracy", "information_coefficient",
            "calibration_score", "drift_ratio", "influence_multiplier", "last_evaluated",
        ] if c in registry]
        st.dataframe(
            registry[columns],
            use_container_width=True,
            hide_index=True,
            column_config={
                "mae": st.column_config.NumberColumn("MAE", format="%.2%"),
                "baseline_mae": st.column_config.NumberColumn("Baseline MAE", format="%.2%"),
                "skill_vs_baseline": st.column_config.NumberColumn("Skill vs baseline", format="%+.1%"),
                "directional_accuracy": st.column_config.NumberColumn("Direction", format="%.1%"),
                "information_coefficient": st.column_config.NumberColumn("IC", format="%.2f"),
                "calibration_score": st.column_config.NumberColumn("Calibration", format="%.2f"),
                "drift_ratio": st.column_config.NumberColumn("Drift", format="%.2f"),
                "influence_multiplier": st.column_config.ProgressColumn("Influence", min_value=0.0, max_value=1.0, format="%.0%%"),
            },
        )

        returns_reg = registry[registry["model"].isin(RETURN_MODEL_ORDER)].copy()
        if not returns_reg.empty:
            returns_reg["Horizon"] = returns_reg["model"].map(RETURN_MODEL_LABELS)
            returns_reg["order"] = returns_reg["model"].map(RETURN_ORDER)
            returns_reg = returns_reg.sort_values("order")
            left, right = st.columns(2)
            with left:
                fig = px.bar(
                    returns_reg,
                    x="Horizon",
                    y="skill_vs_baseline",
                    hover_data=["status", "matured_predictions"],
                    title="Live MAE skill vs baseline",
                )
                fig.update_yaxes(tickformat="+.1%", zeroline=True)
                st.plotly_chart(fig, use_container_width=True)
            with right:
                fig = px.bar(
                    returns_reg,
                    x="Horizon",
                    y="directional_accuracy",
                    hover_data=["status", "matured_predictions"],
                    title="Live directional accuracy",
                )
                fig.update_yaxes(tickformat=".1%", range=[0, 1])
                fig.add_hline(y=0.5, line_dash="dash", annotation_text="50%")
                st.plotly_chart(fig, use_container_width=True)

            model_pick = st.selectbox(
                "Explain a model's governance state",
                list(returns_reg["model"]),
                format_func=lambda m: f"{RETURN_MODEL_LABELS.get(m, m)} — {m}",
                key="governance_model_selector",
            )
            reg_row = returns_reg[returns_reg["model"] == model_pick].iloc[-1]
            st.info(governance_explanation(reg_row))

        with st.expander("What do the governance metrics mean?"):
            st.markdown(
                """
- **MAE**: average absolute forecast error. Lower is better.
- **Skill vs baseline**: improvement in MAE versus the contemporaneous naive baseline. Positive is better.
- **Directional accuracy**: how often the forecast got outperform/underperform direction correct.
- **IC**: rank correlation between prediction and realized outcome; useful for cross-sectional ranking.
- **Calibration**: whether higher stated confidence has historically corresponded to lower errors.
- **Drift ratio**: recent MAE divided by the preceding MAE window. Above 1 means errors have recently worsened.
- **Influence**: governance multiplier. Only the governed 12M model is eligible for production portfolio blending; short horizons remain research-only.
                """
            )


# ---------------------------------------------------------------------------
# Agent experiments
# ---------------------------------------------------------------------------
with tab_agents:
    st.subheader("Autonomous challenger research")
    st.write(
        "The research loop proposes bounded hypotheses, tests approved model families with purged expanding walk-forward validation, "
        "and lets the skeptic gate reject weak results. It records candidates but never edits production model code automatically."
    )

    last_research = latest_state(provider_state, "ml_research_loop", "last_run", {}) or {}
    r1, r2, r3 = st.columns(3)
    r1.metric("Last loop", date_text(last_research.get("started_at")) if isinstance(last_research, dict) else "—")
    r2.metric("Experiments in last loop", int(last_research.get("experiments", 0) or 0) if isinstance(last_research, dict) else 0)
    r3.metric(
        "Promotion candidates in last loop",
        len(last_research.get("promotion_candidates", []) or []) if isinstance(last_research, dict) else 0,
    )

    if experiments.empty:
        st.info("No autonomous experiments are stored yet. Use **Run challenger research now** to create the first bounded experiment batch.")
    else:
        model_options = ["All"] + sorted(experiments["model"].dropna().astype(str).unique().tolist())
        status_options = ["All"] + sorted(experiments["status"].dropna().astype(str).unique().tolist())
        f1, f2 = st.columns(2)
        model_filter = f1.selectbox("Model", model_options, key="experiment_model_filter")
        status_filter = f2.selectbox("Experiment status", status_options, key="experiment_status_filter")
        view = experiments.copy()
        if model_filter != "All":
            view = view[view["model"] == model_filter]
        if status_filter != "All":
            view = view[view["status"] == status_filter]

        show_cols = [
            "experiment_id", "started_at", "model", "estimator", "feature_set", "status", "training_rows",
            "walk_forward_n", "mae_improvement_vs_baseline", "directional_accuracy_edge_vs_baseline",
            "r2", "promotion_candidate", "skeptic_reasons",
        ]
        st.dataframe(
            view[[c for c in show_cols if c in view]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "mae_improvement_vs_baseline": st.column_config.NumberColumn("MAE edge", format="%+.1%"),
                "directional_accuracy_edge_vs_baseline": st.column_config.NumberColumn("Direction edge", format="%+.1%"),
                "r2": st.column_config.NumberColumn("OOS R²", format="%.2f"),
                "promotion_candidate": st.column_config.CheckboxColumn("Candidate"),
            },
        )

        chartable = view.dropna(subset=["mae_improvement_vs_baseline", "directional_accuracy_edge_vs_baseline"]).copy()
        if not chartable.empty:
            fig = px.scatter(
                chartable,
                x="mae_improvement_vs_baseline",
                y="directional_accuracy_edge_vs_baseline",
                size="training_rows",
                color="status",
                hover_name="experiment_id",
                hover_data=["model", "estimator", "feature_set", "r2", "promotion_candidate"],
                title="Challenger experiment map",
            )
            fig.update_xaxes(tickformat="+.1%", zeroline=True)
            fig.update_yaxes(tickformat="+.1%", zeroline=True)
            fig.add_vline(x=0, line_dash="dash")
            fig.add_hline(y=0, line_dash="dash")
            st.plotly_chart(fig, use_container_width=True)

        if not view.empty:
            exp_id = st.selectbox("Explain an experiment", list(view["experiment_id"].astype(str)), key="experiment_detail")
            exp = view[view["experiment_id"].astype(str) == exp_id].iloc[0]
            st.markdown(f"**Hypothesis**  \n{exp.get('hypothesis') or '—'}")
            st.markdown(
                f"**Result:** `{exp.get('status')}` · estimator `{exp.get('estimator')}` · feature set `{exp.get('feature_set')}` · "
                f"MAE edge {pct(exp.get('mae_improvement_vs_baseline'), 1, signed=True)} · "
                f"direction edge {pct(exp.get('directional_accuracy_edge_vs_baseline'), 1, signed=True)} · "
                f"OOS R² {num(exp.get('r2'))}."
            )
            reasons = str(exp.get("skeptic_reasons") or "").strip()
            if reasons:
                st.warning(f"Skeptic: {reasons}")
            elif bool(exp.get("promotion_candidate")):
                st.success("This challenger cleared the current research thresholds and is logged as a promotion candidate for review. It has not been promoted automatically.")
            else:
                st.info("The skeptic did not record a rejection reason, but the experiment did not meet the stricter promotion-candidate threshold.")


# ---------------------------------------------------------------------------
# Data health
# ---------------------------------------------------------------------------
with tab_data:
    st.subheader("Point-in-time data health")
    counts = table_counts(DB)
    dates = latest_data_dates(DB)
    cols = st.columns(5)
    cols[0].metric("Symbols", counts.get("symbols", 0))
    cols[1].metric("Price rows", f"{counts.get('prices', 0):,}")
    cols[2].metric("Fundamental rows", f"{counts.get('fundamentals', 0):,}")
    cols[3].metric("Prediction rows", f"{counts.get('predictions', 0):,}")
    cols[4].metric("Experiment rows", f"{counts.get('research_experiments', 0):,}")

    st.dataframe(
        pd.DataFrame([{"Dataset": k, "Latest stored date": v or "—"} for k, v in dates.items()]),
        hide_index=True,
        use_container_width=True,
    )

    st.subheader("Persistent learning state")
    last_cycle = latest_state(provider_state, "continual_learning", "last_cycle", {}) or {}
    if not last_cycle:
        st.info("No completed continual-learning cycle state is stored yet.")
    else:
        short = last_cycle.get("short_horizon", {}) if isinstance(last_cycle, dict) else {}
        risks = last_cycle.get("risk_forecasts", {}) if isinstance(last_cycle, dict) else {}
        research = last_cycle.get("research_loop", {}) if isinstance(last_cycle, dict) else {}
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Cycle matured", int(last_cycle.get("matured", 0) or 0))
        s2.metric("Short-horizon journaled", int(last_cycle.get("short_horizon_journaled", 0) or 0))
        s3.metric("Risk status", str(risks.get("status") or "—") if isinstance(risks, dict) else "—")
        s4.metric("Research-loop status", str(research.get("status") or "—") if isinstance(research, dict) else "—")
        with st.expander("Show raw last-cycle audit state"):
            st.json(last_cycle)

    st.info(
        "The dashboard is intentionally backed by the same local SQLite database as the scheduled learning workflow. "
        "If a forecast or realized outcome is missing here, the UI leaves it blank rather than estimating or inventing it."
    )


st.divider()
st.caption(
    "Research governance: short-horizon and forward-risk models are context-only; autonomous experiments cannot self-promote; "
    "the deterministic equity-research model remains authoritative for valuation assumptions."
)
