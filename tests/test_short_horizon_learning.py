from __future__ import annotations

from pathlib import Path
import json
import tempfile

import numpy as np
import pandas as pd

import machine_learning  # noqa: F401 - installs learning extensions
from machine_learning.history_store import HistoryStore
from machine_learning import continual_learning as cl
from machine_learning.short_horizon import SHORT_HORIZONS, generate_short_horizon_predictions


def _price_rows(symbol, dates, values):
    return [
        {
            "symbol": symbol,
            "date": d.date().isoformat(),
            "close": float(v),
            "adj_close": float(v),
            "volume": 1_000_000,
            "source": "Yahoo Finance",
        }
        for d, v in zip(dates, values)
    ]


def _insert_feature(store, symbol: str, as_of: str, seed: float):
    values = {
        "revenue_growth": 0.08 + seed,
        "operating_margin": 0.20 + seed,
        "net_margin": 0.15 + seed,
        "fcf_margin": 0.17 + seed,
        "capex_to_revenue": 0.08,
        "rd_to_revenue": 0.10,
        "roe": 0.18 + seed,
        "net_debt_to_revenue": 0.05,
        "momentum_12m": 0.10 + seed,
        "momentum_6m": 0.05 + seed,
        "volatility_6m": 0.22,
        "drawdown_12m": -0.15,
    }
    with store.connect() as con:
        con.execute(
            """INSERT INTO features(
                   symbol,as_of,target_date,revenue_growth,operating_margin,net_margin,fcf_margin,
                   capex_to_revenue,rd_to_revenue,roe,net_debt_to_revenue,momentum_12m,momentum_6m,
                   volatility_6m,drawdown_12m,target_excess_return_12m,source,built_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                symbol, as_of, "2025-12-31", values["revenue_growth"], values["operating_margin"],
                values["net_margin"], values["fcf_margin"], values["capex_to_revenue"],
                values["rd_to_revenue"], values["roe"], values["net_debt_to_revenue"],
                values["momentum_12m"], values["momentum_6m"], values["volatility_6m"],
                values["drawdown_12m"], 0.10, "Test PIT", "2025-01-01T00:00:00+00:00",
            ),
        )


def test_short_horizon_models_are_registered_but_do_not_replace_12m_model():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        cl.ensure_learning_schema(store)
        with store.connect() as con:
            models = {r[0] for r in con.execute("SELECT model FROM model_registry").fetchall()}
        assert set(SHORT_HORIZONS).issubset(models)
        assert "Expected 12M Excess Return" in models


def test_one_month_prediction_matures_against_true_one_month_excess_return():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        cl.ensure_learning_schema(store)
        dates = pd.bdate_range("2024-01-02", "2024-04-30")
        stock = np.linspace(100.0, 120.0, len(dates))
        bench = np.linspace(100.0, 108.0, len(dates))
        store.upsert_prices(_price_rows("AAA", dates, stock))
        store.upsert_prices(_price_rows("SPY", dates, bench))
        with store.connect() as con:
            con.execute(
                """INSERT INTO predictions(
                       run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,
                       model_version,realized_at,realized_value,error,created_at,target_type,
                       baseline_value,baseline_name,evaluation_version)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    "short-test", "AAA", "Expected 1M Excess Return", "2024-01-02", 30,
                    json.dumps(0.05), "Moderate", json.dumps({"benchmark": "SPY"}),
                    "ml-short-horizon-v1", None, None, None, "2024-01-02T12:00:00+00:00",
                    "1m_excess_return", 0.0, "Trailing 1M excess-return baseline", None,
                ),
            )
        out = cl.mature_predictions(store, "SPY")
        assert out["short_horizon_matured"] == 1
        with store.connect() as con:
            row = con.execute(
                "SELECT realized_value,evaluation_version FROM predictions WHERE model='Expected 1M Excess Return'"
            ).fetchone()
        assert row[0] is not None
        assert row[1] == "continual-learning-short-v1"


def test_generator_journals_real_1m_3m_6m_targets_from_point_in_time_history():
    with tempfile.TemporaryDirectory() as td:
        store = HistoryStore(Path(td) / "ml.sqlite")
        cl.ensure_learning_schema(store)
        dates = pd.bdate_range("2022-01-03", "2025-12-31")
        bench = 100.0 * np.exp(np.linspace(0.0, 0.25, len(dates)))
        store.upsert_prices(_price_rows("SPY", dates, bench))

        symbols = ["AAA", "BBB", "CCC", "DDD"]
        feature_dates = ["2022-03-01", "2022-09-01", "2023-03-01", "2023-09-01", "2024-03-01", "2024-09-02"]
        for j, symbol in enumerate(symbols):
            stock = 90.0 * np.exp(np.linspace(0.0, 0.35 + j * 0.04, len(dates)))
            store.upsert_prices(_price_rows(symbol, dates, stock))
            for i, as_of in enumerate(feature_dates):
                _insert_feature(store, symbol, as_of, seed=(j * 0.01 + i * 0.002))

        out = generate_short_horizon_predictions(
            store, "SPY", min_training_rows=12, journal_spacing_days=0
        )
        assert out["status"] == "PASS"
        assert out["journaled"] > 0
        with store.connect() as con:
            rows = con.execute(
                """SELECT DISTINCT model,target_type,horizon_days FROM predictions
                   WHERE model_version='ml-short-horizon-v1'"""
            ).fetchall()
        seen = {r[0]: (r[1], r[2]) for r in rows}
        assert seen["Expected 1M Excess Return"] == ("1m_excess_return", 30)
        assert seen["Expected 3M Excess Return"] == ("3m_excess_return", 91)
        assert seen["Expected 6M Excess Return"] == ("6m_excess_return", 182)
