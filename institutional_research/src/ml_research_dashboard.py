from __future__ import annotations

"""Read-only helpers for the Streamlit ML Research Lab.

The dashboard intentionally reads the same persistent SQLite evidence used by the continual-
learning stack.  These helpers contain no trading actions and do not retrain or promote models.
They only normalize stored predictions, governance metrics, forward-risk state and autonomous
research experiments into display-friendly tables.
"""

from pathlib import Path
import json
import math
import sqlite3
from typing import Any

import numpy as np
import pandas as pd


RETURN_MODEL_ORDER = [
    "Expected 1D Excess Return",
    "Expected 1W Excess Return",
    "Expected 1M Excess Return",
    "Expected 3M Excess Return",
    "Expected 6M Excess Return",
    "Expected 12M Excess Return",
]

RETURN_MODEL_LABELS = {
    "Expected 1D Excess Return": "1D",
    "Expected 1W Excess Return": "1W",
    "Expected 1M Excess Return": "1M",
    "Expected 3M Excess Return": "3M",
    "Expected 6M Excess Return": "6M",
    "Expected 12M Excess Return": "12M",
}

RISK_MODEL_LABELS = {
    "Expected 1W Realized Volatility": "1W Volatility",
    "Expected 1M Forward Drawdown": "1M Drawdown",
}

NUMERIC_REGISTRY_COLUMNS = [
    "matured_predictions",
    "mae",
    "baseline_mae",
    "skill_vs_baseline",
    "directional_accuracy",
    "information_coefficient",
    "calibration_score",
    "drift_ratio",
    "influence_multiplier",
]


def finite(value: Any, default=None):
    try:
        out = float(value)
        return out if math.isfinite(out) else default
    except Exception:
        return default


def safe_json(value: Any, default=None):
    if value in (None, ""):
        return default
    if isinstance(value, (dict, list, int, float, bool)):
        return value
    try:
        return json.loads(value)
    except Exception:
        return default


def prediction_number(value: Any):
    obj = safe_json(value, value)
    if isinstance(obj, bool):
        return None
    if isinstance(obj, dict):
        for key in ("prediction", "value", "expected_return", "forecast", "median_prediction"):
            if key in obj:
                return finite(obj.get(key))
    return finite(obj)


def table_exists(db_path: Path, table: str) -> bool:
    path = Path(db_path)
    if not path.exists():
        return False
    try:
        with sqlite3.connect(path) as con:
            row = con.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
                (table,),
            ).fetchone()
        return bool(row)
    except Exception:
        return False


def read_sql(db_path: Path, query: str, params: tuple = ()) -> pd.DataFrame:
    path = Path(db_path)
    if not path.exists():
        return pd.DataFrame()
    try:
        with sqlite3.connect(path) as con:
            return pd.read_sql_query(query, con, params=params)
    except Exception:
        return pd.DataFrame()


def load_dashboard_tables(db_path: Path) -> dict[str, pd.DataFrame]:
    """Load only tables used by the visual research layer.

    Missing tables return empty frames so an older database can still open safely after schema
    migration.  The Streamlit page invokes the normal HistoryStore migration before this loader.
    """
    path = Path(db_path)
    queries = {
        "predictions": """
            SELECT id,run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,
                   model_version,realized_at,realized_value,error,created_at,target_type,
                   baseline_value,baseline_name,evaluation_version
            FROM predictions ORDER BY datetime(created_at),id
        """,
        "registry": "SELECT * FROM model_registry ORDER BY model",
        "performance": "SELECT * FROM model_performance ORDER BY evaluated_at,model,id",
        "training": "SELECT * FROM training_runs ORDER BY trained_at,model,id",
        "experiments": "SELECT * FROM research_experiments ORDER BY datetime(started_at),id",
        "provider_state": "SELECT provider,state_key,state_value,updated_at FROM provider_state ORDER BY provider,state_key",
    }
    out: dict[str, pd.DataFrame] = {}
    for name, query in queries.items():
        table = "model_registry" if name == "registry" else (
            "model_performance" if name == "performance" else (
                "training_runs" if name == "training" else (
                    "research_experiments" if name == "experiments" else name
                )
            )
        )
        out[name] = read_sql(path, query) if table_exists(path, table) else pd.DataFrame()
    return out


def prepare_predictions(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return pd.DataFrame()
    out = frame.copy()
    out["prediction_num"] = out.get("prediction", pd.Series(index=out.index, dtype=object)).map(prediction_number)
    for col in ("realized_value", "error", "baseline_value", "horizon_days"):
        if col in out:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("created_at", "realized_at", "as_of"):
        if col in out:
            out[col] = pd.to_datetime(out[col], errors="coerce", utc=True)
    out["as_of_day"] = out["as_of"].dt.date if "as_of" in out else None
    sort_cols = [c for c in ("created_at", "id") if c in out]
    if sort_cols:
        out = out.sort_values(sort_cols)
    subset = [c for c in ("model", "symbol", "as_of_day", "model_version") if c in out]
    if subset:
        out = out.drop_duplicates(subset, keep="last")
    return out.reset_index(drop=True)


def latest_return_forecasts(predictions: pd.DataFrame) -> pd.DataFrame:
    if predictions is None or predictions.empty:
        return pd.DataFrame()
    p = prepare_predictions(predictions)
    p = p[p["model"].isin(RETURN_MODEL_ORDER) & p["prediction_num"].notna()].copy()
    if p.empty:
        return p
    p["horizon"] = p["model"].map(RETURN_MODEL_LABELS)
    p["horizon_order"] = p["model"].map({m: i for i, m in enumerate(RETURN_MODEL_ORDER)})
    p = p.sort_values(["created_at", "id"]).groupby(["model", "symbol"], as_index=False).tail(1)
    return p.sort_values(["symbol", "horizon_order"]).reset_index(drop=True)


def forecast_matrix(latest: pd.DataFrame) -> pd.DataFrame:
    if latest is None or latest.empty:
        return pd.DataFrame()
    pivot = latest.pivot_table(index="symbol", columns="horizon", values="prediction_num", aggfunc="last")
    ordered = [RETURN_MODEL_LABELS[m] for m in RETURN_MODEL_ORDER if RETURN_MODEL_LABELS[m] in pivot.columns]
    pivot = pivot.reindex(columns=ordered)
    return pivot.reset_index().rename_axis(None, axis=1)


def latest_state(provider_state: pd.DataFrame, provider: str, key: str, default=None):
    if provider_state is None or provider_state.empty:
        return default
    view = provider_state[
        (provider_state["provider"].astype(str) == str(provider))
        & (provider_state["state_key"].astype(str) == str(key))
    ]
    if view.empty:
        return default
    row = view.tail(1).iloc[0]
    return safe_json(row.get("state_value"), default)


def risk_forecast_frame(provider_state: pd.DataFrame) -> pd.DataFrame:
    state = latest_state(provider_state, "risk_forecasts", "last_generation", {}) or {}
    models = state.get("models", {}) if isinstance(state, dict) else {}
    rows = []
    for model, details in models.items():
        if not isinstance(details, dict):
            continue
        per_symbol = details.get("per_symbol") or {}
        for symbol, value in per_symbol.items():
            val = finite(value)
            if val is None:
                continue
            rows.append({
                "model": model,
                "risk_label": RISK_MODEL_LABELS.get(model, model),
                "symbol": str(symbol).upper(),
                "value": val,
                "status": details.get("status"),
                "training_rows": details.get("training_rows"),
                "median_prediction": finite(details.get("median_prediction")),
                "as_of": state.get("date"),
            })
    return pd.DataFrame(rows)


def risk_matrix(risk: pd.DataFrame) -> pd.DataFrame:
    if risk is None or risk.empty:
        return pd.DataFrame()
    pivot = risk.pivot_table(index="symbol", columns="risk_label", values="value", aggfunc="last")
    ordered = [label for label in RISK_MODEL_LABELS.values() if label in pivot.columns]
    return pivot.reindex(columns=ordered).reset_index().rename_axis(None, axis=1)


def prepare_registry(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return pd.DataFrame()
    out = frame.copy()
    for col in NUMERIC_REGISTRY_COLUMNS:
        if col in out:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    if "model" in out:
        order = {m: i for i, m in enumerate(RETURN_MODEL_ORDER)}
        out["_order"] = out["model"].map(order).fillna(99)
        out = out.sort_values(["_order", "model"]).drop(columns="_order")
    return out.reset_index(drop=True)


def _metric(payload: dict, key: str):
    return finite(payload.get(key)) if isinstance(payload, dict) else None


def prepare_experiments(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return pd.DataFrame()
    rows = []
    for _, row in frame.iterrows():
        validation = safe_json(row.get("validation_json"), {}) or {}
        skeptic = safe_json(row.get("skeptic_json"), {}) or {}
        reasons = skeptic.get("reasons") if isinstance(skeptic, dict) else None
        rows.append({
            "id": row.get("id"),
            "experiment_id": row.get("experiment_id"),
            "model": row.get("model"),
            "hypothesis": row.get("hypothesis"),
            "estimator": row.get("estimator"),
            "feature_set": row.get("feature_set"),
            "started_at": pd.to_datetime(row.get("started_at"), errors="coerce", utc=True),
            "status": row.get("status"),
            "training_rows": pd.to_numeric(row.get("training_rows"), errors="coerce"),
            "walk_forward_n": _metric(validation, "n"),
            "mae": _metric(validation, "mae"),
            "baseline_mae": _metric(validation, "baseline_mae"),
            "mae_improvement_vs_baseline": _metric(validation, "mae_improvement_vs_baseline"),
            "directional_accuracy": _metric(validation, "directional_accuracy"),
            "baseline_directional_accuracy": _metric(validation, "baseline_directional_accuracy"),
            "directional_accuracy_edge_vs_baseline": _metric(validation, "directional_accuracy_edge_vs_baseline"),
            "r2": _metric(validation, "r2"),
            "promotion_candidate": bool(row.get("promotion_candidate")) or bool(skeptic.get("promotion_candidate")),
            "skeptic_passed": skeptic.get("passed") if isinstance(skeptic, dict) else None,
            "skeptic_reasons": "; ".join(str(x) for x in reasons) if isinstance(reasons, list) else str(reasons or ""),
            "notes": row.get("notes"),
        })
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["started_at", "id"], ascending=[False, False]).reset_index(drop=True)
    return out


def latest_training_run(training: pd.DataFrame, model: str):
    if training is None or training.empty:
        return None
    view = training[training["model"].astype(str) == str(model)].copy()
    if view.empty:
        return None
    view["trained_at_ts"] = pd.to_datetime(view["trained_at"], errors="coerce", utc=True)
    view = view.sort_values(["trained_at_ts", "id" if "id" in view else "trained_at_ts"])
    return view.iloc[-1]


def training_validation_table(training: pd.DataFrame, model: str) -> pd.DataFrame:
    row = latest_training_run(training, model)
    if row is None:
        return pd.DataFrame()
    payload = safe_json(row.get("validation_json"), {}) or {}
    walk = payload.get("walk_forward", payload) if isinstance(payload, dict) else {}
    rows = []
    if isinstance(walk, dict):
        for estimator, metrics in walk.items():
            if not isinstance(metrics, dict):
                continue
            rows.append({
                "estimator": estimator,
                "n": _metric(metrics, "n"),
                "mae": _metric(metrics, "mae"),
                "baseline_mae": _metric(metrics, "baseline_mae"),
                "mae_improvement_vs_baseline": _metric(metrics, "mae_improvement_vs_baseline"),
                "directional_accuracy": _metric(metrics, "directional_accuracy"),
                "directional_accuracy_edge_vs_baseline": _metric(metrics, "directional_accuracy_edge_vs_baseline"),
                "r2": _metric(metrics, "r2"),
            })
    return pd.DataFrame(rows)


def table_counts(db_path: Path) -> dict[str, int]:
    path = Path(db_path)
    names = [
        "symbols", "prices", "fundamentals", "earnings", "features", "predictions",
        "training_runs", "model_performance", "model_registry", "research_experiments",
    ]
    out: dict[str, int] = {}
    if not path.exists():
        return out
    with sqlite3.connect(path) as con:
        for name in names:
            try:
                out[name] = int(con.execute(f"SELECT COUNT(*) FROM {name}").fetchone()[0])
            except Exception:
                out[name] = 0
    return out


def latest_data_dates(db_path: Path) -> dict[str, str | None]:
    path = Path(db_path)
    queries = {
        "Latest price": "SELECT MAX(date) FROM prices",
        "Latest fundamental availability": "SELECT MAX(available_date) FROM fundamentals",
        "Latest prediction": "SELECT MAX(created_at) FROM predictions",
        "Latest model evaluation": "SELECT MAX(evaluated_at) FROM model_performance",
        "Latest experiment": "SELECT MAX(started_at) FROM research_experiments",
    }
    out = {}
    if not path.exists():
        return {k: None for k in queries}
    with sqlite3.connect(path) as con:
        for label, query in queries.items():
            try:
                row = con.execute(query).fetchone()
                out[label] = row[0] if row else None
            except Exception:
                out[label] = None
    return out


def governance_explanation(row: pd.Series | dict | None) -> str:
    if row is None:
        return "No live governance evidence has been recorded yet."
    get = row.get if hasattr(row, "get") else lambda *_: None
    status = str(get("status") or "UNPROVEN")
    n = int(finite(get("matured_predictions"), 0) or 0)
    skill = finite(get("skill_vs_baseline"))
    direction = finite(get("directional_accuracy"))
    drift = finite(get("drift_ratio"))
    parts = [f"Status: {status}. {n} matured live forecast(s) are currently available."]
    if skill is not None:
        parts.append(f"Live MAE skill versus the stored naive baseline is {skill:+.1%}.")
    if direction is not None:
        parts.append(f"Directional accuracy is {direction:.1%}.")
    if drift is not None:
        parts.append(f"Recent/prior error drift ratio is {drift:.2f}; values above 1 indicate worsening recent error.")
    if status == "CHAMPION":
        parts.append("The model has passed the project's live governance gate, but this is still research evidence rather than a trading instruction.")
    elif status == "DEMOTED":
        parts.append("The model has failed one or more live governance checks, so its downstream influence should remain suppressed.")
    elif status == "CHALLENGER":
        parts.append("The model has some live evidence but has not yet met the champion gate.")
    else:
        parts.append("The model does not yet have enough mature live evidence for a strong governance conclusion.")
    return " ".join(parts)


def return_forecast_explanation(horizon: str, value: float | None, benchmark: str = "SPY") -> str:
    if value is None:
        return f"No current {horizon} forecast is available."
    direction = "outperform" if value > 0 else "underperform" if value < 0 else "match"
    return (
        f"The model's current {horizon} estimate is {value:+.2%} excess return versus {benchmark}. "
        f"In other words, it estimates the stock may {direction} {benchmark} by roughly {abs(value):.2%} over that horizon. "
        "This is an ML research signal, not a total-return target or a trade recommendation."
    )


def risk_explanation(volatility: float | None, drawdown: float | None) -> str:
    parts = []
    if volatility is not None:
        parts.append(
            f"The 1-week risk model estimates {volatility:.1%} annualized realized volatility over the next five trading days."
        )
    if drawdown is not None:
        parts.append(
            f"The 1-month downside model estimates a worst decline from today's price of about {drawdown:.1%} over the next 21 trading days."
        )
    if not parts:
        return "No current forward-risk estimate is available for this symbol."
    parts.append("These are context-only risk estimates and do not automatically change portfolio weights.")
    return " ".join(parts)
