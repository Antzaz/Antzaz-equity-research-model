from __future__ import annotations

"""Governed continual learning for the investment-research ML layer.

The project deliberately uses batch continual learning rather than unconstrained online updates:
- every normal ML run can retrain on the expanded point-in-time database;
- old forecasts are closed only after their actual horizon/event becomes observable;
- realized forecast errors are compared with a contemporaneous simple baseline;
- drift, calibration, directional accuracy and information coefficient are monitored;
- model influence is governed through a local champion/challenger registry.

The SQLite database remains local and gitignored.  Nothing in this module executes trades or
changes deterministic valuation outputs.  It only evaluates ML evidence and controls how much
that evidence is allowed to influence downstream portfolio expected-return blending.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import argparse
import json
import math
import sqlite3
from typing import Any

import numpy as np
import pandas as pd


LEARNING_SCHEMA_VERSION = 2
CURRENT_MODEL_VERSION = "ml-layer-v5-continual"
DEFAULT_BENCHMARK = "SPY"
CONFIDENCE_SCORE = {"High": 1.0, "Moderate": 0.65, "Low": 0.30}
SUPERVISED_MODELS = {"Expected 12M Excess Return", "Consensus / Earnings Surprise"}
KNOWN_MODELS = [
    "Expected 12M Excess Return",
    "Consensus / Earnings Surprise",
    "Financial Anomaly Detection",
    "Market Regime Classifier",
    "AI Impact ML",
    "Portfolio ML / Position Sizing",
]

_ACTIVE_STORE = None
_HOOKS_INSTALLED = False


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _finite(value: Any, default=None):
    try:
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _json(value: Any, default=None):
    if value in (None, ""):
        return default
    if isinstance(value, (dict, list, int, float)):
        return value
    try:
        return json.loads(value)
    except Exception:
        return default if default is not None else value


def _prediction_number(value: Any):
    obj = _json(value, value)
    if isinstance(obj, bool):
        return None
    if isinstance(obj, (int, float)):
        return _finite(obj)
    if isinstance(obj, dict):
        for key in ("prediction", "value", "expected_return", "forecast"):
            if key in obj:
                out = _finite(obj.get(key))
                if out is not None:
                    return out
    return _finite(obj)


def _symbol_candidates(symbol: str) -> list[str]:
    s = str(symbol or "").upper().strip()
    out = [s]
    if "." in s and s.count(".") == 1:
        left, right = s.split(".")
        if len(right) == 1:
            out.append(f"{left}-{right}")
    if "-" in s and s.count("-") == 1:
        left, right = s.split("-")
        if len(right) == 1:
            out.append(f"{left}.{right}")
    return list(dict.fromkeys(x for x in out if x))


def ensure_learning_schema(store) -> None:
    """Idempotently migrate an existing HistoryStore database for continual learning."""
    with store.connect() as con:
        existing = {row[1] for row in con.execute("PRAGMA table_info(predictions)").fetchall()}
        additions = {
            "target_type": "TEXT",
            "baseline_value": "REAL",
            "baseline_name": "TEXT",
            "evaluation_version": "TEXT",
        }
        for name, sql_type in additions.items():
            if name not in existing:
                con.execute(f"ALTER TABLE predictions ADD COLUMN {name} {sql_type}")

        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS training_runs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model TEXT NOT NULL,
                model_version TEXT NOT NULL,
                trained_at TEXT NOT NULL,
                status TEXT,
                confidence TEXT,
                prediction TEXT,
                training_rows INTEGER,
                validation_json TEXT,
                drivers_json TEXT,
                details_json TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_training_runs_model_time
                ON training_runs(model, trained_at);

            CREATE TABLE IF NOT EXISTS model_performance(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model TEXT NOT NULL,
                model_version TEXT NOT NULL,
                evaluated_at TEXT NOT NULL,
                sample_count INTEGER NOT NULL,
                mae REAL,
                rmse REAL,
                bias REAL,
                directional_accuracy REAL,
                baseline_mae REAL,
                skill_vs_baseline REAL,
                information_coefficient REAL,
                calibration_score REAL,
                drift_ratio REAL,
                status TEXT NOT NULL,
                influence_multiplier REAL NOT NULL,
                details_json TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_model_performance_model_time
                ON model_performance(model, evaluated_at);

            CREATE TABLE IF NOT EXISTS model_registry(
                model TEXT PRIMARY KEY,
                champion_version TEXT,
                status TEXT NOT NULL,
                matured_predictions INTEGER NOT NULL DEFAULT 0,
                mae REAL,
                baseline_mae REAL,
                skill_vs_baseline REAL,
                directional_accuracy REAL,
                information_coefficient REAL,
                calibration_score REAL,
                drift_ratio REAL,
                influence_multiplier REAL NOT NULL,
                last_evaluated TEXT,
                metrics_json TEXT,
                updated_at TEXT NOT NULL
            );

            CREATE TRIGGER IF NOT EXISTS trg_predictions_continual_v5
            AFTER INSERT ON predictions
            BEGIN
                UPDATE predictions
                SET model_version = CASE
                        WHEN NEW.model_version IS NULL OR NEW.model_version='ml-layer-v4-purged-pit'
                            THEN 'ml-layer-v5-continual'
                        ELSE NEW.model_version END,
                    target_type = CASE
                        WHEN NEW.model='Expected 12M Excess Return' THEN '12m_excess_return'
                        WHEN NEW.model='Consensus / Earnings Surprise' THEN 'next_eps_surprise'
                        ELSE COALESCE(target_type,'context_only') END,
                    baseline_name = CASE
                        WHEN NEW.model='Expected 12M Excess Return' THEN 'Trailing 12M excess-return baseline'
                        WHEN NEW.model='Consensus / Earnings Surprise' THEN 'Prior 4Q mean surprise baseline'
                        ELSE baseline_name END
                WHERE id=NEW.id;
            END;
            """
        )
        con.execute(
            "INSERT INTO metadata(key,value,updated_at) VALUES('learning_schema_version',?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at",
            (str(LEARNING_SCHEMA_VERSION), _now()),
        )

    _ensure_default_registry(store)


def _ensure_default_registry(store) -> None:
    now = _now()
    with store.connect() as con:
        for model in KNOWN_MODELS:
            if model in SUPERVISED_MODELS:
                status = "UNPROVEN"
                multiplier = 0.25
            else:
                status = "CONTEXT_ONLY"
                multiplier = 1.0
            con.execute(
                """INSERT INTO model_registry(
                       model,champion_version,status,matured_predictions,influence_multiplier,
                       last_evaluated,metrics_json,updated_at
                   ) VALUES(?,?,?,?,?,?,?,?)
                   ON CONFLICT(model) DO NOTHING""",
                (model, CURRENT_MODEL_VERSION, status, 0, multiplier, None,
                 json.dumps({"note": "Awaiting realized out-of-sample evidence."}), now),
            )


def _price_series(store, symbol: str) -> pd.Series:
    frames = []
    for candidate in _symbol_candidates(symbol):
        try:
            pf = store.price_frame(candidate)
        except Exception:
            pf = pd.DataFrame()
        if pf is not None and not pf.empty:
            px = pd.to_numeric(pf.get("adj_close"), errors="coerce")
            if px is None or px.dropna().empty:
                px = pd.to_numeric(pf.get("close"), errors="coerce")
            else:
                px = px.fillna(pd.to_numeric(pf.get("close"), errors="coerce"))
            px = px.dropna()
            if not px.empty:
                frames.append(px)
    if not frames:
        return pd.Series(dtype=float)
    out = pd.concat(frames).sort_index()
    out = out[~out.index.duplicated(keep="first")]
    return out


def _as_naive_timestamp(value) -> pd.Timestamp | None:
    try:
        ts = pd.Timestamp(value)
        if pd.isna(ts):
            return None
        if ts.tzinfo is not None:
            ts = ts.tz_convert("UTC").tz_localize(None)
        return ts
    except Exception:
        return None


def _first_at_or_after(series: pd.Series, when: pd.Timestamp):
    if series.empty:
        return None, None
    idx = series.index
    if getattr(idx, "tz", None) is not None:
        series = series.copy()
        series.index = idx.tz_convert("UTC").tz_localize(None)
    view = series.loc[series.index >= when]
    if view.empty:
        return None, None
    return view.index[0], _finite(view.iloc[0])


def _trailing_return(series: pd.Series, as_of: pd.Timestamp, observations: int = 252):
    if series.empty:
        return None
    if getattr(series.index, "tz", None) is not None:
        series = series.copy()
        series.index = series.index.tz_convert("UTC").tz_localize(None)
    view = series.loc[series.index <= as_of].dropna().tail(observations + 1)
    if len(view) < 60:
        return None
    p0 = _finite(view.iloc[0]); p1 = _finite(view.iloc[-1])
    return p1 / p0 - 1 if p0 not in (None, 0) and p1 is not None else None


def _baseline_expected_return(store, symbol: str, benchmark: str, as_of: pd.Timestamp) -> float:
    stock = _trailing_return(_price_series(store, symbol), as_of)
    bench = _trailing_return(_price_series(store, benchmark), as_of)
    if stock is None:
        return 0.0
    return float(stock - (bench or 0.0))


def _earnings_rows(store, symbol: str) -> pd.DataFrame:
    frames = []
    try:
        with store.connect() as con:
            for candidate in _symbol_candidates(symbol):
                df = pd.read_sql_query(
                    "SELECT reported_date,surprise_pct FROM earnings WHERE symbol=? "
                    "AND surprise_pct IS NOT NULL ORDER BY reported_date",
                    con,
                    params=(candidate,),
                )
                if not df.empty:
                    frames.append(df)
    except Exception:
        return pd.DataFrame()
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True).drop_duplicates("reported_date")
    df["reported_date"] = pd.to_datetime(df["reported_date"], errors="coerce")
    return df.dropna(subset=["reported_date"]).sort_values("reported_date")


def _baseline_earnings(store, symbol: str, as_of: pd.Timestamp) -> float:
    df = _earnings_rows(store, symbol)
    if df.empty:
        return 0.0
    prior = pd.to_numeric(df.loc[df["reported_date"] <= as_of, "surprise_pct"], errors="coerce").dropna().tail(4)
    return float(prior.mean()) if not prior.empty else 0.0


def backfill_prediction_metadata(store, default_benchmark: str = DEFAULT_BENCHMARK) -> int:
    ensure_learning_schema(store)
    with store.connect() as con:
        rows = con.execute(
            """SELECT id,symbol,model,as_of,features_json,target_type,baseline_value,baseline_name
               FROM predictions
               WHERE target_type IS NULL OR baseline_value IS NULL OR baseline_name IS NULL"""
        ).fetchall()
    changed = 0
    for row in rows:
        pid, symbol, model, as_of, features_json, target_type, baseline_value, baseline_name = row
        ts = _as_naive_timestamp(as_of)
        if ts is None:
            continue
        details = _json(features_json, {}) or {}
        benchmark = str(details.get("benchmark") or default_benchmark).upper()
        if model == "Expected 12M Excess Return":
            target_type = target_type or "12m_excess_return"
            baseline_name = baseline_name or "Trailing 12M excess-return baseline"
            if baseline_value is None:
                baseline_value = _baseline_expected_return(store, symbol, benchmark, ts)
        elif model == "Consensus / Earnings Surprise":
            target_type = target_type or "next_eps_surprise"
            baseline_name = baseline_name or "Prior 4Q mean surprise baseline"
            if baseline_value is None:
                baseline_value = _baseline_earnings(store, symbol, ts)
        else:
            target_type = target_type or "context_only"
        with store.connect() as con:
            con.execute(
                "UPDATE predictions SET target_type=?,baseline_value=?,baseline_name=? WHERE id=?",
                (target_type, _finite(baseline_value), baseline_name, pid),
            )
        changed += 1
    return changed


def _realized_12m_excess(store, symbol: str, benchmark: str, as_of: pd.Timestamp, horizon_days: int):
    stock = _price_series(store, symbol)
    bench = _price_series(store, benchmark)
    if stock.empty or bench.empty:
        return None, None
    target = as_of + pd.Timedelta(days=int(horizon_days or 365))
    _, s0 = _first_at_or_after(stock, as_of)
    sdate, s1 = _first_at_or_after(stock, target)
    _, b0 = _first_at_or_after(bench, as_of)
    bdate, b1 = _first_at_or_after(bench, target)
    if None in (s0, s1, b0, b1) or s0 == 0 or b0 == 0:
        return None, None
    realized = (s1 / s0 - 1) - (b1 / b0 - 1)
    realized_at = max(sdate, bdate) if sdate is not None and bdate is not None else sdate or bdate
    return float(realized), realized_at


def _realized_next_earnings(store, symbol: str, as_of: pd.Timestamp):
    df = _earnings_rows(store, symbol)
    if df.empty:
        return None, None
    future = df[df["reported_date"] > as_of]
    if future.empty:
        return None, None
    row = future.iloc[0]
    value = _finite(row.get("surprise_pct"))
    return value, row.get("reported_date") if value is not None else (None, None)


def mature_predictions(store, default_benchmark: str = DEFAULT_BENCHMARK) -> dict:
    """Close only forecasts whose ex-post target is observable in the local point-in-time store."""
    ensure_learning_schema(store)
    backfill_prediction_metadata(store, default_benchmark)
    with store.connect() as con:
        rows = con.execute(
            """SELECT id,symbol,model,as_of,horizon_days,prediction,features_json
               FROM predictions WHERE realized_at IS NULL
               AND model IN ('Expected 12M Excess Return','Consensus / Earnings Surprise')
               ORDER BY id"""
        ).fetchall()
    matured = 0; skipped = 0
    for pid, symbol, model, as_of, horizon_days, prediction, features_json in rows:
        pred = _prediction_number(prediction)
        ts = _as_naive_timestamp(as_of)
        if pred is None or ts is None:
            skipped += 1
            continue
        details = _json(features_json, {}) or {}
        benchmark = str(details.get("benchmark") or default_benchmark).upper()
        if model == "Expected 12M Excess Return":
            realized, realized_at = _realized_12m_excess(store, symbol, benchmark, ts, horizon_days or 365)
        else:
            realized, realized_at = _realized_next_earnings(store, symbol, ts)
        if realized is None or realized_at is None:
            continue
        realized_ts = _as_naive_timestamp(realized_at)
        with store.connect() as con:
            con.execute(
                """UPDATE predictions SET realized_at=?,realized_value=?,error=?,evaluation_version=?
                   WHERE id=? AND realized_at IS NULL""",
                (realized_ts.isoformat() if realized_ts is not None else str(realized_at),
                 float(realized), float(realized - pred), "continual-learning-v2", pid),
            )
        matured += 1
    return {"matured": matured, "skipped_invalid": skipped, "pending": max(0, len(rows) - matured - skipped)}


def _prediction_frame(store, model: str | None = None) -> pd.DataFrame:
    ensure_learning_schema(store)
    q = """SELECT id,run_id,symbol,model,as_of,horizon_days,prediction,confidence,model_version,
                  realized_at,realized_value,error,baseline_value,baseline_name,target_type,created_at
           FROM predictions WHERE realized_at IS NOT NULL AND realized_value IS NOT NULL"""
    params: tuple = ()
    if model:
        q += " AND model=?"; params = (model,)
    q += " ORDER BY datetime(created_at),id"
    with store.connect() as con:
        df = pd.read_sql_query(q, con, params=params)
    if df.empty:
        return df
    df["prediction_num"] = df["prediction"].map(_prediction_number)
    df["realized_value"] = pd.to_numeric(df["realized_value"], errors="coerce")
    df["baseline_value"] = pd.to_numeric(df["baseline_value"], errors="coerce")
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce", utc=True)
    df["realized_at"] = pd.to_datetime(df["realized_at"], errors="coerce", utc=True)
    df["as_of_day"] = pd.to_datetime(df["as_of"], errors="coerce", utc=True).dt.date
    # Re-running the same research on the same day must not manufacture sample size.
    df = df.sort_values(["created_at", "id"]).drop_duplicates(
        ["model", "symbol", "as_of_day", "model_version"], keep="last"
    )
    return df.dropna(subset=["prediction_num", "realized_value"])


def _calibration_score(df: pd.DataFrame):
    if len(df) < 5:
        return None
    conf = df["confidence"].map(CONFIDENCE_SCORE).astype(float)
    abs_err = (df["realized_value"] - df["prediction_num"]).abs()
    if conf.nunique() < 2 or abs_err.nunique() < 2:
        return None
    out = conf.corr(-abs_err, method="spearman")
    return _finite(out)


def _performance_metrics(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"n": 0}
    y = pd.to_numeric(df["realized_value"], errors="coerce")
    p = pd.to_numeric(df["prediction_num"], errors="coerce")
    b = pd.to_numeric(df["baseline_value"], errors="coerce")
    ok = y.notna() & p.notna()
    y = y[ok]; p = p[ok]; b = b[ok]
    err = y - p
    mae = float(err.abs().mean()); rmse = float(np.sqrt(np.mean(np.square(err)))); bias = float(err.mean())
    direction = float((np.sign(y) == np.sign(p)).mean()) if len(y) else None
    b_ok = b.notna()
    baseline_mae = float((y[b_ok] - b[b_ok]).abs().mean()) if b_ok.any() else None
    skill = 1 - mae / baseline_mae if baseline_mae not in (None, 0) else None
    ic = _finite(p.corr(y, method="spearman")) if len(y) >= 5 and p.nunique() > 1 and y.nunique() > 1 else None
    calibration = _calibration_score(df.loc[ok])

    drift = None; recent_mae = None; prior_mae = None
    ordered = df.loc[ok].sort_values("realized_at")
    window = min(20, max(5, len(ordered) // 3))
    if len(ordered) >= window * 2:
        ae = (ordered["realized_value"] - ordered["prediction_num"]).abs()
        recent_mae = float(ae.tail(window).mean())
        prior_mae = float(ae.iloc[-2 * window:-window].mean())
        drift = recent_mae / prior_mae if prior_mae > 0 else None
    return {
        "n": int(len(y)), "mae": mae, "rmse": rmse, "bias": bias,
        "directional_accuracy": direction, "baseline_mae": baseline_mae,
        "skill_vs_baseline": skill, "information_coefficient": ic,
        "calibration_score": calibration, "drift_ratio": drift,
        "recent_mae": recent_mae, "prior_mae": prior_mae,
    }


def _governance(metrics: dict) -> tuple[str, float]:
    n = int(metrics.get("n") or 0)
    skill = _finite(metrics.get("skill_vs_baseline"))
    da = _finite(metrics.get("directional_accuracy"))
    ic = _finite(metrics.get("information_coefficient"))
    drift = _finite(metrics.get("drift_ratio"))
    if n < 5:
        return "UNPROVEN", 0.25
    if n >= 8 and ((skill is not None and skill <= -0.05) or (da is not None and da < 0.45) or (drift is not None and drift >= 1.75)):
        return "DEMOTED", 0.0

    skill_score = max(0.0, min(1.0, ((skill or 0.0) + 0.05) / 0.30))
    da_score = max(0.0, min(1.0, ((da or 0.50) - 0.48) / 0.15))
    ic_score = max(0.0, min(1.0, ((ic or 0.0) + 0.02) / 0.22))
    sample_score = min(1.0, n / 40.0)
    drift_penalty = 0.55 if drift is not None and drift > 1.35 else 1.0
    evidence = (0.40 * skill_score + 0.25 * da_score + 0.15 * ic_score + 0.20 * sample_score) * drift_penalty

    champion = n >= 12 and (skill or -1) >= 0.05 and (da or 0) >= 0.53 and (drift is None or drift <= 1.35)
    if champion:
        return "CHAMPION", float(max(0.65, min(1.0, 0.65 + 0.35 * evidence)))
    return "CHALLENGER", float(max(0.20, min(0.65, 0.20 + 0.45 * evidence)))


def evaluate_models(store) -> pd.DataFrame:
    """Evaluate realized supervised forecasts and update the model registry."""
    ensure_learning_schema(store)
    frames = []
    now = _now()
    for model in SUPERVISED_MODELS:
        all_df = _prediction_frame(store, model)
        if all_df.empty:
            continue
        versions = list(dict.fromkeys(str(x or "unknown") for x in all_df["model_version"].fillna("unknown")))
        candidates = []
        for version in versions:
            view = all_df[all_df["model_version"].fillna("unknown").astype(str) == version]
            metrics = _performance_metrics(view); status, influence = _governance(metrics)
            candidates.append((version, metrics, status, influence))
            _insert_performance(store, model, version, metrics, status, influence, now)
        aggregate = _performance_metrics(all_df); agg_status, agg_influence = _governance(aggregate)
        _insert_performance(store, model, "ALL", aggregate, agg_status, agg_influence, now)

        eligible = [x for x in candidates if int(x[1].get("n") or 0) >= 5]
        if eligible:
            def score(item):
                m = item[1]
                return (_finite(m.get("skill_vs_baseline"), -9), _finite(m.get("information_coefficient"), -9), int(m.get("n") or 0))
            champion_version = max(eligible, key=score)[0]
        else:
            champion_version = candidates[-1][0] if candidates else CURRENT_MODEL_VERSION
        _update_registry(store, model, champion_version, aggregate, agg_status, agg_influence, now)
        frames.append({"Model": model, "ChampionVersion": champion_version, "Status": agg_status,
                       "InfluenceMultiplier": agg_influence, **aggregate})
    return pd.DataFrame(frames)


def _insert_performance(store, model, version, metrics, status, influence, now):
    with store.connect() as con:
        con.execute(
            """INSERT INTO model_performance(
                   model,model_version,evaluated_at,sample_count,mae,rmse,bias,directional_accuracy,
                   baseline_mae,skill_vs_baseline,information_coefficient,calibration_score,drift_ratio,
                   status,influence_multiplier,details_json)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (model, version, now, int(metrics.get("n") or 0), _finite(metrics.get("mae")),
             _finite(metrics.get("rmse")), _finite(metrics.get("bias")),
             _finite(metrics.get("directional_accuracy")), _finite(metrics.get("baseline_mae")),
             _finite(metrics.get("skill_vs_baseline")), _finite(metrics.get("information_coefficient")),
             _finite(metrics.get("calibration_score")), _finite(metrics.get("drift_ratio")), status,
             float(influence), json.dumps(metrics, default=str)),
        )


def _update_registry(store, model, champion_version, metrics, status, influence, now):
    with store.connect() as con:
        con.execute(
            """INSERT INTO model_registry(
                   model,champion_version,status,matured_predictions,mae,baseline_mae,skill_vs_baseline,
                   directional_accuracy,information_coefficient,calibration_score,drift_ratio,
                   influence_multiplier,last_evaluated,metrics_json,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(model) DO UPDATE SET
                   champion_version=excluded.champion_version,status=excluded.status,
                   matured_predictions=excluded.matured_predictions,mae=excluded.mae,
                   baseline_mae=excluded.baseline_mae,skill_vs_baseline=excluded.skill_vs_baseline,
                   directional_accuracy=excluded.directional_accuracy,
                   information_coefficient=excluded.information_coefficient,
                   calibration_score=excluded.calibration_score,drift_ratio=excluded.drift_ratio,
                   influence_multiplier=excluded.influence_multiplier,last_evaluated=excluded.last_evaluated,
                   metrics_json=excluded.metrics_json,updated_at=excluded.updated_at""",
            (model, champion_version, status, int(metrics.get("n") or 0), _finite(metrics.get("mae")),
             _finite(metrics.get("baseline_mae")), _finite(metrics.get("skill_vs_baseline")),
             _finite(metrics.get("directional_accuracy")), _finite(metrics.get("information_coefficient")),
             _finite(metrics.get("calibration_score")), _finite(metrics.get("drift_ratio")),
             float(influence), now, json.dumps(metrics, default=str), now),
        )


def record_training_results(store, results) -> int:
    """Persist gated walk-forward/training evidence for model-learning diagnostics."""
    if store is None:
        return 0
    ensure_learning_schema(store)
    count = 0; trained_at = _now()
    for result in results or []:
        try:
            metrics = dict(result.metrics or {})
            details = dict(result.details or {})
            drivers = list(result.drivers or [])
            rows = int(metrics.get("training_rows") or metrics.get("monthly_rows") or len(metrics.get("years") or []) or 0)
            with store.connect() as con:
                con.execute(
                    """INSERT INTO training_runs(
                           model,model_version,trained_at,status,confidence,prediction,training_rows,
                           validation_json,drivers_json,details_json)
                       VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (result.name, CURRENT_MODEL_VERSION, trained_at, result.status, result.confidence,
                     json.dumps(result.prediction, default=str), rows, json.dumps(metrics, default=str),
                     json.dumps(drivers, default=str), json.dumps(details, default=str)),
                )
            count += 1
        except Exception:
            continue
    return count


def run_continual_learning_cycle(store, benchmark: str = DEFAULT_BENCHMARK) -> dict:
    ensure_learning_schema(store)
    metadata = backfill_prediction_metadata(store, benchmark)
    maturity = mature_predictions(store, benchmark)
    performance = evaluate_models(store)
    summary = {
        "ran_at": _now(), "metadata_backfilled": metadata, **maturity,
        "models_evaluated": int(len(performance)),
    }
    try:
        store.set_state("continual_learning", "last_cycle", summary)
    except Exception:
        pass
    return summary


def registry_frame(store) -> pd.DataFrame:
    ensure_learning_schema(store)
    with store.connect() as con:
        return pd.read_sql_query("SELECT * FROM model_registry ORDER BY model", con)


def performance_frame(store) -> pd.DataFrame:
    ensure_learning_schema(store)
    with store.connect() as con:
        return pd.read_sql_query("SELECT * FROM model_performance ORDER BY evaluated_at,model", con)


def training_frame(store) -> pd.DataFrame:
    ensure_learning_schema(store)
    with store.connect() as con:
        return pd.read_sql_query("SELECT * FROM training_runs ORDER BY trained_at,model", con)


def set_active_store(store) -> None:
    global _ACTIVE_STORE
    _ACTIVE_STORE = store


def install_hooks() -> None:
    """Install non-invasive hooks so normal ml_research.py runs learn automatically."""
    global _HOOKS_INSTALLED
    if _HOOKS_INSTALLED:
        return
    from .history_store import HistoryStore
    from . import quality as quality_module

    original_post = HistoryStore.__post_init__
    original_expected_frame = HistoryStore.expected_return_frame
    original_gate_results = quality_module.gate_results

    def post_init(self):
        original_post(self)
        ensure_learning_schema(self)
        set_active_store(self)

    def expected_return_frame(self, *args, **kwargs):
        set_active_store(self)
        try:
            cycle = run_continual_learning_cycle(self, DEFAULT_BENCHMARK)
            if cycle.get("matured"):
                print(f"[ml-learning] matured {cycle['matured']} prior forecast(s); registry refreshed")
        except Exception as exc:
            print(f"[ml-learning] feedback cycle skipped: {exc}")
        return original_expected_frame(self, *args, **kwargs)

    def gate_results(results):
        gated = original_gate_results(results)
        if _ACTIVE_STORE is not None:
            try:
                record_training_results(_ACTIVE_STORE, gated)
            except Exception:
                pass
        return gated

    HistoryStore.__post_init__ = post_init
    HistoryStore.expected_return_frame = expected_return_frame
    quality_module.gate_results = gate_results
    _HOOKS_INSTALLED = True


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Close matured ML forecasts and refresh model-learning governance")
    parser.add_argument("--db", default=str(Path(__file__).resolve().parents[1] / "ml_data" / "ml_history.sqlite"))
    parser.add_argument("--benchmark", default=DEFAULT_BENCHMARK)
    args = parser.parse_args(argv)
    from .history_store import HistoryStore
    store = HistoryStore(Path(args.db))
    summary = run_continual_learning_cycle(store, args.benchmark)
    print(json.dumps(summary, indent=2, default=str))
    print(registry_frame(store).to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
