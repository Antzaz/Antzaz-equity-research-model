from __future__ import annotations

from pathlib import Path
import json
import tempfile

import numpy as np
import pandas as pd

from machine_learning.history_store import HistoryStore
from machine_learning.continual_learning import (
    CURRENT_MODEL_VERSION,
    ensure_learning_schema,
    mature_predictions,
    evaluate_models,
    run_continual_learning_cycle,
)
from institutional_research.src.continual_patch import apply_learning_governance


def _price_rows(symbol, dates, values):
    return [
        {"symbol": symbol, "date": d.date().isoformat(), "close": float(v), "adj_close": float(v), "volume": 1_000_000,
         "source": "Yahoo Finance"}
        for d, v in zip(dates, values)
    ]


def _insert_prediction(store, *, symbol="AAA", model="Expected 12M Excess Return", as_of="2024-01-02",
                       prediction=0.10, horizon=365, confidence="High", realized=None,
                       baseline=None, version=CURRENT_MODEL_VERSION, suffix="1"):
    with store.connect() as con:
        con.execute(
            """INSERT INTO predictions(
                   run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,
                   model_version,realized_at,realized_value,error,baseline_value,created_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (f"run-{suffix}", symbol, model, as_of, horizon, json.dumps(prediction), confidence,
             json.dumps({"benchmark": "SPY"}), version,
             ("2025-02-01T00:00:00+00:00" if realized is not None else None),
             realized, (realized - prediction if realized is not None else None), baseline,
             f"2024-01-{min(28, int(suffix) if str(suffix).isdigit() else 1):02d}T12:00:00+00:00"),
        )


def test_expected_return_prediction_matures_from_price_history():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        ensure_learning_schema(store)
        dates = pd.bdate_range("2023-12-01", "2025-03-31")
        stock = np.linspace(100.0, 140.0, len(dates))
        bench = np.linspace(100.0, 115.0, len(dates))
        store.upsert_prices(_price_rows("AAA", dates, stock))
        store.upsert_prices(_price_rows("SPY", dates, bench))
        _insert_prediction(store)

        summary = run_continual_learning_cycle(store, "SPY")
        assert summary["matured"] == 1
        with store.connect() as con:
            row = con.execute(
                "SELECT realized_value,error,baseline_value,target_type,evaluation_version FROM predictions"
            ).fetchone()
        assert row[0] is not None and row[1] is not None
        assert row[2] is not None
        assert row[3] == "12m_excess_return"
        assert row[4] == "continual-learning-v2"
        with store.connect() as con:
            reg = con.execute(
                "SELECT status,matured_predictions,influence_multiplier FROM model_registry "
                "WHERE model='Expected 12M Excess Return'"
            ).fetchone()
        assert reg[0] == "UNPROVEN"
        assert reg[1] == 1
        assert 0 <= reg[2] <= 1


def test_earnings_surprise_prediction_matures_on_next_report():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        ensure_learning_schema(store)
        earnings = [
            ("2023-01-20", 0.01), ("2023-04-20", 0.02), ("2023-07-20", -0.01),
            ("2023-10-20", 0.03), ("2024-01-25", 0.05),
        ]
        store.upsert_earnings([
            {"symbol": "AAA", "reported_date": d, "surprise_pct": s, "source": "Test"}
            for d, s in earnings
        ])
        _insert_prediction(
            store, model="Consensus / Earnings Surprise", as_of="2024-01-01",
            prediction=0.04, horizon=None, suffix="2",
        )
        out = mature_predictions(store, "SPY")
        assert out["matured"] == 1
        with store.connect() as con:
            row = con.execute(
                "SELECT realized_value,baseline_value,target_type FROM predictions"
            ).fetchone()
        assert abs(row[0] - 0.05) < 1e-12
        assert row[1] is not None
        assert row[2] == "next_eps_surprise"


def test_strong_realized_history_promotes_champion():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        ensure_learning_schema(store)
        for i in range(1, 16):
            pred = 0.08 + i * 0.001
            realized = pred + (0.005 if i % 2 else -0.004)
            _insert_prediction(
                store, symbol=f"A{i:02d}", as_of=f"2023-{((i-1)//12)+1:02d}-{((i-1)%12)+1:02d}",
                prediction=pred, realized=realized, baseline=0.0, suffix=str(i),
            )
        perf = evaluate_models(store)
        row = perf[perf["Model"] == "Expected 12M Excess Return"].iloc[0]
        assert row["Status"] == "CHAMPION"
        assert row["InfluenceMultiplier"] >= 0.65
        assert row["skill_vs_baseline"] > 0


def test_bad_realized_history_demotes_model_and_zeroes_portfolio_influence():
    with tempfile.TemporaryDirectory() as td:
        db = Path(td) / "ml.sqlite"
        store = HistoryStore(db)
        ensure_learning_schema(store)
        for i in range(1, 10):
            _insert_prediction(
                store, symbol=f"B{i:02d}", as_of=f"2023-02-{i:02d}", prediction=0.10,
                realized=-0.10, baseline=0.0, suffix=str(i),
            )
        evaluate_models(store)
        with store.connect() as con:
            reg = con.execute(
                "SELECT status,influence_multiplier FROM model_registry "
                "WHERE model='Expected 12M Excess Return'"
            ).fetchone()
        assert reg[0] == "DEMOTED"
        assert reg[1] == 0.0

        inputs = pd.DataFrame([{
            "Ticker": "AAA", "ClassicalExpectedReturn": 0.08, "MLTotalReturn": 0.18,
            "MLBlendWeight": 0.50, "BlendedMLExpectedReturn": 0.13,
        }])
        governed, meta = apply_learning_governance(inputs, {}, db)
        assert governed.loc[0, "PreLearningMLBlendWeight"] == 0.50
        assert governed.loc[0, "MLBlendWeight"] == 0.0
        assert abs(governed.loc[0, "BlendedMLExpectedReturn"] - 0.08) < 1e-12
        assert governed.loc[0, "LearningStatus"] == "DEMOTED"
        assert meta["continual_learning"]["influence_multiplier"] == 0.0
