from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pandas as pd

from institutional_research.src.ml_research_dashboard import (
    forecast_matrix,
    governance_explanation,
    latest_return_forecasts,
    load_dashboard_tables,
    prepare_experiments,
    prepare_predictions,
    risk_forecast_frame,
    risk_matrix,
)


def test_latest_return_forecasts_include_fast_horizons_and_deduplicate():
    raw = pd.DataFrame([
        {
            "id": 1,
            "run_id": "a",
            "symbol": "AAA",
            "model": "Expected 1D Excess Return",
            "as_of": "2026-09-15",
            "horizon_days": 1,
            "prediction": json.dumps(0.01),
            "confidence": "Moderate",
            "features_json": "{}",
            "model_version": "v1",
            "realized_at": None,
            "realized_value": None,
            "error": None,
            "created_at": "2026-09-15T08:00:00+00:00",
            "target_type": "1d_excess_return",
            "baseline_value": 0.0,
            "baseline_name": "baseline",
            "evaluation_version": None,
        },
        {
            "id": 2,
            "run_id": "b",
            "symbol": "AAA",
            "model": "Expected 1D Excess Return",
            "as_of": "2026-09-15",
            "horizon_days": 1,
            "prediction": json.dumps(0.02),
            "confidence": "High",
            "features_json": "{}",
            "model_version": "v1",
            "realized_at": None,
            "realized_value": None,
            "error": None,
            "created_at": "2026-09-15T09:00:00+00:00",
            "target_type": "1d_excess_return",
            "baseline_value": 0.0,
            "baseline_name": "baseline",
            "evaluation_version": None,
        },
        {
            "id": 3,
            "run_id": "c",
            "symbol": "AAA",
            "model": "Expected 1W Excess Return",
            "as_of": "2026-09-15",
            "horizon_days": 7,
            "prediction": json.dumps({"prediction": 0.04}),
            "confidence": "Moderate",
            "features_json": "{}",
            "model_version": "v2",
            "realized_at": None,
            "realized_value": None,
            "error": None,
            "created_at": "2026-09-15T09:00:00+00:00",
            "target_type": "1w_excess_return",
            "baseline_value": 0.01,
            "baseline_name": "baseline",
            "evaluation_version": None,
        },
    ])
    prepared = prepare_predictions(raw)
    latest = latest_return_forecasts(prepared)
    assert set(latest["horizon"]) == {"1D", "1W"}
    one_day = latest[latest["horizon"] == "1D"].iloc[0]
    assert abs(float(one_day["prediction_num"]) - 0.02) < 1e-12
    matrix = forecast_matrix(latest)
    assert list(matrix.columns) == ["symbol", "1D", "1W"]
    assert abs(float(matrix.iloc[0]["1W"]) - 0.04) < 1e-12


def test_risk_state_is_normalized_into_symbol_matrix():
    payload = {
        "date": "2026-09-16",
        "models": {
            "Expected 1W Realized Volatility": {
                "status": "PASS",
                "training_rows": 1200,
                "median_prediction": 0.22,
                "per_symbol": {"AAA": 0.20, "BBB": 0.30},
            },
            "Expected 1M Forward Drawdown": {
                "status": "PASS",
                "training_rows": 1200,
                "median_prediction": -0.08,
                "per_symbol": {"AAA": -0.07, "BBB": -0.11},
            },
        },
    }
    state = pd.DataFrame([
        {
            "provider": "risk_forecasts",
            "state_key": "last_generation",
            "state_value": json.dumps(payload),
            "updated_at": "2026-09-16T08:00:00Z",
        }
    ])
    risk = risk_forecast_frame(state)
    assert len(risk) == 4
    matrix = risk_matrix(risk)
    assert set(matrix.columns) == {"symbol", "1W Volatility", "1M Drawdown"}
    aaa = matrix[matrix["symbol"] == "AAA"].iloc[0]
    assert abs(float(aaa["1W Volatility"]) - 0.20) < 1e-12
    assert abs(float(aaa["1M Drawdown"]) + 0.07) < 1e-12


def test_experiment_parser_surfaces_skeptic_and_candidate():
    raw = pd.DataFrame([
        {
            "id": 1,
            "experiment_id": "exp-1",
            "model": "Expected 1W Excess Return",
            "hypothesis": "Try Extra Trees",
            "estimator": "extra_trees",
            "feature_set": "base",
            "started_at": "2026-09-16T08:00:00Z",
            "status": "PASS",
            "training_rows": 2500,
            "validation_json": json.dumps({
                "n": 200,
                "mae": 0.02,
                "baseline_mae": 0.024,
                "mae_improvement_vs_baseline": 0.1667,
                "directional_accuracy": 0.56,
                "directional_accuracy_edge_vs_baseline": 0.03,
                "r2": 0.05,
            }),
            "skeptic_json": json.dumps({
                "passed": True,
                "promotion_candidate": True,
                "reasons": [],
            }),
            "promotion_candidate": 1,
            "notes": "research only",
        }
    ])
    out = prepare_experiments(raw)
    row = out.iloc[0]
    assert bool(row["promotion_candidate"])
    assert bool(row["skeptic_passed"])
    assert abs(float(row["mae_improvement_vs_baseline"]) - 0.1667) < 1e-12
    assert row["skeptic_reasons"] == ""


def test_governance_explanation_is_plain_english():
    text = governance_explanation({
        "status": "CHALLENGER",
        "matured_predictions": 9,
        "skill_vs_baseline": 0.08,
        "directional_accuracy": 0.56,
        "drift_ratio": 1.10,
    })
    assert "CHALLENGER" in text
    assert "9 matured" in text
    assert "8.0%" in text
    assert "56.0%" in text


def test_load_dashboard_tables_tolerates_partial_database(tmp_path: Path):
    db = tmp_path / "ml.sqlite"
    with sqlite3.connect(db) as con:
        con.executescript(
            """
            CREATE TABLE predictions(
                id INTEGER PRIMARY KEY, run_id TEXT, symbol TEXT, model TEXT, as_of TEXT,
                horizon_days INTEGER, prediction TEXT, confidence TEXT, features_json TEXT,
                model_version TEXT, realized_at TEXT, realized_value REAL, error REAL,
                created_at TEXT, target_type TEXT, baseline_value REAL, baseline_name TEXT,
                evaluation_version TEXT
            );
            CREATE TABLE provider_state(
                provider TEXT, state_key TEXT, state_value TEXT, updated_at TEXT
            );
            """
        )
    tables = load_dashboard_tables(db)
    assert "predictions" in tables
    assert "registry" in tables
    assert tables["registry"].empty
    assert tables["predictions"].empty
