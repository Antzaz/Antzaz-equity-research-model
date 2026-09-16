from __future__ import annotations

"""Governed autonomous research loop for the ML layer.

The loop behaves like a small research desk rather than a self-modifying trading bot:

- Research agent: reads model-governance state and proposes bounded hypotheses.
- Experiment agent: evaluates approved sklearn challengers and feature ablations.
- Simulation agent: uses purged expanding walk-forward validation only.
- Skeptic agent: rejects challengers that fail sample-size, baseline or directional tests.
- Promotion gate: records candidates for review; it never rewrites production model code.

This module deliberately reuses the existing point-in-time store, validation engine and model
registry.  It adds research evidence without changing deterministic valuation or executing trades.
"""

from datetime import datetime, timezone
import json
import math
from typing import Any

import numpy as np
import pandas as pd
from sklearn.ensemble import ExtraTreesRegressor, HistGradientBoostingRegressor, RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import ElasticNet, Ridge
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from .common import RANDOM_STATE
from .models import EXPECTED_FEATURES
from .short_horizon import (
    FAST_FEATURES,
    SHORT_HORIZONS,
    model_feature_columns,
    training_frame_for_model,
)
from .validation import expanding_walk_forward


RESEARCH_LOOP_VERSION = "ml-research-loop-v1"
EXPERIMENT_CADENCE_DAYS = 7


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _finite(value: Any, default=None):
    try:
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def ensure_research_schema(store) -> None:
    with store.connect() as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS research_experiments(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                experiment_id TEXT NOT NULL UNIQUE,
                model TEXT NOT NULL,
                hypothesis TEXT NOT NULL,
                estimator TEXT NOT NULL,
                feature_set TEXT NOT NULL,
                started_at TEXT NOT NULL,
                status TEXT NOT NULL,
                training_rows INTEGER,
                validation_json TEXT,
                skeptic_json TEXT,
                promotion_candidate INTEGER NOT NULL DEFAULT 0,
                notes TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_research_experiments_model_time
                ON research_experiments(model, started_at);
            """
        )


def _candidate_estimators() -> dict[str, Any]:
    return {
        "hist_gradient_boosting": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", HistGradientBoostingRegressor(
                max_iter=220,
                max_leaf_nodes=16,
                learning_rate=0.04,
                l2_regularization=0.8,
                random_state=RANDOM_STATE,
            )),
        ]),
        "extra_trees": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", ExtraTreesRegressor(
                n_estimators=220,
                min_samples_leaf=8,
                max_features=0.8,
                n_jobs=-1,
                random_state=RANDOM_STATE,
            )),
        ]),
        "random_forest": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", RandomForestRegressor(
                n_estimators=180,
                min_samples_leaf=10,
                max_features=0.7,
                n_jobs=-1,
                random_state=RANDOM_STATE,
            )),
        ]),
        "elastic_net": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scale", StandardScaler()),
            ("model", ElasticNet(alpha=0.015, l1_ratio=0.20, max_iter=5000, random_state=RANDOM_STATE)),
        ]),
        "ridge": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scale", StandardScaler()),
            ("model", Ridge(alpha=4.0)),
        ]),
    }


def _feature_sets(model_name: str, base: list[str]) -> dict[str, list[str]]:
    sets = {"base": list(base)}
    if base == FAST_FEATURES:
        sets["no_volume"] = [c for c in base if c != "volume_z_21d"]
        sets["momentum_volatility"] = [
            c for c in base
            if c.startswith("ret_") or c.startswith("excess_") or c.startswith("vol_")
            or c.startswith("market_") or c.startswith("drawdown_") or c == "beta_63d"
        ]
    elif base == EXPECTED_FEATURES:
        sets["fundamental_only"] = [c for c in base if c not in {"momentum_12m", "momentum_6m", "volatility_6m", "drawdown_12m"}]
        sets["market_plus_quality"] = [
            c for c in base
            if c in {"revenue_growth", "operating_margin", "fcf_margin", "roe", "net_debt_to_revenue",
                     "momentum_12m", "momentum_6m", "volatility_6m", "drawdown_12m"}
        ]
    return {k: v for k, v in sets.items() if len(v) >= 4}


def _registry_snapshot(store) -> list[dict[str, Any]]:
    try:
        with store.connect() as con:
            rows = con.execute(
                """SELECT model,champion_version,status,matured_predictions,skill_vs_baseline,
                          directional_accuracy,information_coefficient,drift_ratio,influence_multiplier
                   FROM model_registry ORDER BY model"""
            ).fetchall()
        keys = ["model", "champion_version", "status", "matured_predictions", "skill_vs_baseline",
                "directional_accuracy", "information_coefficient", "drift_ratio", "influence_multiplier"]
        return [dict(zip(keys, row)) for row in rows]
    except Exception:
        return []


def _research_datasets(store, benchmark: str) -> list[dict[str, Any]]:
    datasets: list[dict[str, Any]] = []
    try:
        frame12 = store.expected_return_frame(min_rows=30)
    except Exception:
        frame12 = pd.DataFrame()
    if not frame12.empty:
        datasets.append({
            "model": "Expected 12M Excess Return",
            "frame": frame12,
            "features": EXPECTED_FEATURES,
            "target": "target_excess_return_12m",
            "min_rows": 30,
        })

    for model_name in SHORT_HORIZONS:
        try:
            frame = training_frame_for_model(store, model_name, benchmark)
        except Exception:
            frame = pd.DataFrame()
        features = model_feature_columns(model_name)
        fast = features == FAST_FEATURES
        datasets.append({
            "model": model_name,
            "frame": frame,
            "features": features,
            "target": "target_excess_return",
            "min_rows": 800 if fast else 40,
        })
    return datasets


def _hypothesis(model: str, estimator: str, feature_set: str, registry: dict[str, Any] | None) -> str:
    status = str((registry or {}).get("status") or "UNPROVEN")
    drift = _finite((registry or {}).get("drift_ratio"))
    reason = "increase out-of-sample skill"
    if status in {"DEMOTED", "CHALLENGER"}:
        reason = "repair weak live skill versus the naive baseline"
    elif drift is not None and drift > 1.35:
        reason = "reduce recent forecast-error drift"
    return f"Test whether {estimator} with the {feature_set} feature set can {reason} for {model}."


def _skeptic(metrics: dict[str, Any], training_rows: int) -> dict[str, Any]:
    n = int(metrics.get("n") or 0)
    mae_edge = _finite(metrics.get("mae_improvement_vs_baseline"))
    da_edge = _finite(metrics.get("directional_accuracy_edge_vs_baseline"))
    r2 = _finite(metrics.get("r2"))
    reasons = []
    if training_rows < 40:
        reasons.append("training sample too small")
    if n < 10:
        reasons.append("walk-forward sample too small")
    if mae_edge is None or mae_edge <= 0:
        reasons.append("no MAE improvement versus expanding-mean baseline")
    if da_edge is not None and da_edge < -0.02:
        reasons.append("directional accuracy materially worse than baseline")
    if r2 is not None and r2 < -0.25:
        reasons.append("strongly negative out-of-sample R2")
    passed = not reasons
    promotion = bool(
        passed
        and n >= 20
        and (mae_edge or 0.0) >= 0.05
        and (da_edge is None or da_edge >= 0.0)
    )
    return {
        "passed": passed,
        "promotion_candidate": promotion,
        "reasons": reasons,
        "walk_forward_n": n,
        "mae_improvement_vs_baseline": mae_edge,
        "directional_accuracy_edge_vs_baseline": da_edge,
        "r2": r2,
    }


def _due(store, force: bool) -> bool:
    if force:
        return True
    try:
        last = store.get_state("ml_research_loop", "last_run", {}) or {}
        stamp = pd.Timestamp(last.get("started_at")) if isinstance(last, dict) and last.get("started_at") else None
        if stamp is None or pd.isna(stamp):
            return True
        if stamp.tzinfo is not None:
            stamp = stamp.tz_convert("UTC").tz_localize(None)
        return (pd.Timestamp.utcnow().tz_localize(None) - stamp).days >= EXPERIMENT_CADENCE_DAYS
    except Exception:
        return True


def run_research_loop(
    store,
    benchmark: str = "SPY",
    *,
    force: bool = False,
    max_experiments: int = 12,
) -> dict:
    """Run bounded autonomous challenger research and persist all evidence."""
    ensure_research_schema(store)
    started_at = _now()
    if not _due(store, force):
        return {"status": "NOT_DUE", "started_at": started_at, "experiments": 0, "promotion_candidates": []}

    registry_rows = _registry_snapshot(store)
    registry = {str(r.get("model")): r for r in registry_rows}
    estimators = _candidate_estimators()
    outcomes = []
    promotion_candidates = []
    experiment_count = 0

    # Rotate the challenger order by ISO week so the loop explores without exploding compute cost.
    week = datetime.now(timezone.utc).isocalendar().week
    estimator_names = list(estimators)
    estimator_names = estimator_names[week % len(estimator_names):] + estimator_names[:week % len(estimator_names)]

    for ds in _research_datasets(store, benchmark):
        if experiment_count >= int(max_experiments):
            break
        frame = ds["frame"]
        if frame is None or len(frame) < int(ds["min_rows"]):
            outcomes.append({"model": ds["model"], "status": "INSUFFICIENT_DATA", "training_rows": 0 if frame is None else int(len(frame))})
            continue
        feature_sets = _feature_sets(ds["model"], ds["features"])
        min_train = max(500 if ds["features"] == FAST_FEATURES else 24, min(5000 if ds["features"] == FAST_FEATURES else 80, len(frame) // 2))
        step = max(1, len(frame) // (14 if ds["features"] == FAST_FEATURES else 20))

        for feature_name, feature_cols in feature_sets.items():
            for estimator_name in estimator_names[:2]:
                if experiment_count >= int(max_experiments):
                    break
                estimator = estimators[estimator_name]
                hypothesis = _hypothesis(ds["model"], estimator_name, feature_name, registry.get(ds["model"]))
                experiment_id = f"{datetime.now(timezone.utc).strftime('%Y%m%d')}-{experiment_count+1:02d}-{abs(hash((ds['model'], estimator_name, feature_name))) % 100000:05d}"
                try:
                    wf = expanding_walk_forward(
                        estimator,
                        frame,
                        feature_cols,
                        ds["target"],
                        min_train=min_train,
                        step=step,
                    )
                    metrics = wf.metrics
                    skeptic = _skeptic(metrics, len(frame))
                    status = "PASS" if skeptic["passed"] else "REJECTED"
                except Exception as exc:
                    metrics = {}
                    skeptic = {"passed": False, "promotion_candidate": False, "reasons": [f"{type(exc).__name__}: {exc}"]}
                    status = "ERROR"

                with store.connect() as con:
                    con.execute(
                        """INSERT OR REPLACE INTO research_experiments(
                               experiment_id,model,hypothesis,estimator,feature_set,started_at,status,
                               training_rows,validation_json,skeptic_json,promotion_candidate,notes)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            experiment_id,
                            ds["model"],
                            hypothesis,
                            estimator_name,
                            feature_name,
                            started_at,
                            status,
                            int(len(frame)),
                            json.dumps(metrics, default=str),
                            json.dumps(skeptic, default=str),
                            1 if skeptic.get("promotion_candidate") else 0,
                            "Research-only challenger; production code is unchanged.",
                        ),
                    )
                item = {
                    "experiment_id": experiment_id,
                    "model": ds["model"],
                    "hypothesis": hypothesis,
                    "estimator": estimator_name,
                    "feature_set": feature_name,
                    "status": status,
                    "training_rows": int(len(frame)),
                    "validation": metrics,
                    "skeptic": skeptic,
                }
                outcomes.append(item)
                experiment_count += 1
                if skeptic.get("promotion_candidate"):
                    promotion_candidates.append(item)

    summary = {
        "status": "PASS",
        "version": RESEARCH_LOOP_VERSION,
        "started_at": started_at,
        "experiments": int(experiment_count),
        "promotion_candidates": promotion_candidates,
        "registry_snapshot": registry_rows,
        "outcomes": outcomes,
        "governance": "Candidates are logged only; no automatic production-code promotion or trading action.",
    }
    try:
        store.set_state("ml_research_loop", "last_run", summary)
    except Exception:
        pass
    return summary
