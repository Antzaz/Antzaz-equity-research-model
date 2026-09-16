from __future__ import annotations

from pathlib import Path
import tempfile

import numpy as np
import pandas as pd

import machine_learning  # noqa: F401 - installs learning extensions
from machine_learning.history_store import HistoryStore
from machine_learning.continual_learning import ensure_learning_schema
from machine_learning.short_horizon import (
    FAST_FEATURES,
    SHORT_HORIZONS,
    build_fast_training_frame,
)
from machine_learning.risk_forecasts import build_risk_training_frame
from machine_learning.research_loop import ensure_research_schema


def _seed_store(path: Path) -> HistoryStore:
    store = HistoryStore(path)
    dates = pd.bdate_range("2024-01-02", periods=320)
    store.upsert_symbols([
        {"symbol": "AAA", "name": "AAA", "source": "Test"},
        {"symbol": "BBB", "name": "BBB", "source": "Test"},
        {"symbol": "SPY", "name": "SPY", "source": "Test"},
    ])
    for symbol, slope, phase in [("AAA", 0.0007, 0.0), ("BBB", 0.0003, 0.7), ("SPY", 0.0004, 1.1)]:
        t = np.arange(len(dates), dtype=float)
        px = 100.0 * np.exp(slope * t + 0.02 * np.sin(t / 11.0 + phase))
        rows = [
            {
                "symbol": symbol,
                "date": d.date().isoformat(),
                "close": float(v),
                "adj_close": float(v),
                "volume": float(1_000_000 + 100_000 * np.sin(i / 7.0 + phase)),
                "source": "Yahoo Finance",
            }
            for i, (d, v) in enumerate(zip(dates, px))
        ]
        store.upsert_prices(rows)

    # Feature rows are used to define the investable research universe and support longer horizons.
    with store.connect() as con:
        for symbol in ("AAA", "BBB"):
            con.execute(
                """INSERT INTO features(
                       symbol,as_of,target_date,revenue_growth,operating_margin,net_margin,fcf_margin,
                       capex_to_revenue,rd_to_revenue,roe,net_debt_to_revenue,price_to_sales,
                       earnings_yield,fcf_yield,book_to_market,ev_to_sales,momentum_12m,momentum_6m,
                       volatility_6m,drawdown_12m,target_excess_return_12m,source,built_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    symbol,
                    dates[-80].date().isoformat(),
                    dates[-1].date().isoformat(),
                    0.10, 0.20, 0.15, 0.12, 0.05, 0.08, 0.18, 0.05,
                    4.0, 0.04, 0.03, 0.20, 4.5, 0.15, 0.08, 0.25, -0.10,
                    0.05, "Test", dates[-80].isoformat(),
                ),
            )
    return store


def test_1d_and_1w_horizons_are_configured_as_fast_models():
    assert SHORT_HORIZONS["Expected 1D Excess Return"]["trading_days"] == 1
    assert SHORT_HORIZONS["Expected 1W Excess Return"]["trading_days"] == 5
    assert SHORT_HORIZONS["Expected 1D Excess Return"]["feature_mode"] == "fast"
    assert SHORT_HORIZONS["Expected 1W Excess Return"]["feature_mode"] == "fast"


def test_fast_training_targets_are_forward_and_leakage_safe():
    with tempfile.TemporaryDirectory() as td:
        store = _seed_store(Path(td) / "ml.sqlite")
        one_day = build_fast_training_frame(store, "SPY", 1, max_rows=5000)
        one_week = build_fast_training_frame(store, "SPY", 5, max_rows=5000)
        assert not one_day.empty and not one_week.empty
        assert set(FAST_FEATURES).issubset(one_day.columns)
        assert (pd.to_datetime(one_day["target_date"]) > pd.to_datetime(one_day["as_of"])).all()
        assert (pd.to_datetime(one_week["target_date"]) > pd.to_datetime(one_week["as_of"])).all()
        assert one_day["target_excess_return"].notna().all()
        assert one_week["target_excess_return"].notna().all()


def test_forward_risk_targets_are_constructed_from_future_prices_only():
    with tempfile.TemporaryDirectory() as td:
        store = _seed_store(Path(td) / "ml.sqlite")
        frame = build_risk_training_frame(store, "SPY", max_rows=5000)
        assert not frame.empty
        vol = frame.dropna(subset=["target_volatility", "target_date_volatility"])
        dd = frame.dropna(subset=["target_drawdown", "target_date_drawdown"])
        assert not vol.empty and not dd.empty
        assert (pd.to_datetime(vol["target_date_volatility"]) > pd.to_datetime(vol["as_of"])).all()
        assert (pd.to_datetime(dd["target_date_drawdown"]) > pd.to_datetime(dd["as_of"])).all()
        assert (vol["target_volatility"] >= 0).all()
        assert (dd["target_drawdown"] <= 0).all()


def test_learning_registry_and_research_experiment_schema_include_extensions():
    with tempfile.TemporaryDirectory() as td:
        store = _seed_store(Path(td) / "ml.sqlite")
        ensure_learning_schema(store)
        ensure_research_schema(store)
        with store.connect() as con:
            models = {row[0] for row in con.execute("SELECT model FROM model_registry").fetchall()}
            table = con.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='research_experiments'"
            ).fetchone()
        assert "Expected 1D Excess Return" in models
        assert "Expected 1W Excess Return" in models
        assert "Expected 1W Realized Volatility" in models
        assert "Expected 1M Forward Drawdown" in models
        assert table is not None
