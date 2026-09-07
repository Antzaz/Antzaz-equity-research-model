from __future__ import annotations

"""Integration shim for 1M/3M/6M governed learning.

Kept separate so the existing 12M model mathematics and portfolio-optimizer contract remain
unchanged.  This module extends only the continual-learning registry, maturation and daily
maintenance cycle.
"""

import pandas as pd

from . import continual_learning as cl
from .short_horizon import SHORT_HORIZONS, generate_short_horizon_predictions


_INSTALLED = False


def install_short_horizon_learning() -> None:
    global _INSTALLED
    if _INSTALLED:
        return

    short_models = set(SHORT_HORIZONS)
    cl.SUPERVISED_MODELS.update(short_models)
    for model in SHORT_HORIZONS:
        if model not in cl.KNOWN_MODELS:
            cl.KNOWN_MODELS.insert(0, model)

    # Fix a legacy tuple-precedence edge case while preserving the earnings model itself.
    def realized_next_earnings(store, symbol: str, as_of: pd.Timestamp):
        df = cl._earnings_rows(store, symbol)
        if df.empty:
            return None, None
        future = df[df["reported_date"] > as_of]
        if future.empty:
            return None, None
        row = future.iloc[0]
        value = cl._finite(row.get("surprise_pct"))
        if value is None:
            return None, None
        return value, row.get("reported_date")

    cl._realized_next_earnings = realized_next_earnings

    original_mature = cl.mature_predictions

    def mature_predictions(store, default_benchmark: str = cl.DEFAULT_BENCHMARK) -> dict:
        base = original_mature(store, default_benchmark)
        cl.ensure_learning_schema(store)
        placeholders = ",".join("?" for _ in short_models)
        with store.connect() as con:
            rows = con.execute(
                f"""SELECT id,symbol,model,as_of,horizon_days,prediction,features_json
                    FROM predictions WHERE realized_at IS NULL
                    AND model IN ({placeholders}) ORDER BY id""",
                tuple(sorted(short_models)),
            ).fetchall()

        matured = 0
        skipped = 0
        for pid, symbol, model, as_of, horizon_days, prediction, features_json in rows:
            pred = cl._prediction_number(prediction)
            ts = cl._as_naive_timestamp(as_of)
            if pred is None or ts is None:
                skipped += 1
                continue
            details = cl._json(features_json, {}) or {}
            benchmark = str(details.get("benchmark") or default_benchmark).upper()
            realized, realized_at = cl._realized_12m_excess(
                store, symbol, benchmark, ts, int(horizon_days or SHORT_HORIZONS[model]["days"])
            )
            if realized is None or realized_at is None:
                continue
            realized_ts = cl._as_naive_timestamp(realized_at)
            with store.connect() as con:
                con.execute(
                    """UPDATE predictions SET realized_at=?,realized_value=?,error=?,evaluation_version=?
                       WHERE id=? AND realized_at IS NULL""",
                    (
                        realized_ts.isoformat() if realized_ts is not None else str(realized_at),
                        float(realized), float(realized - pred), "continual-learning-short-v1", pid,
                    ),
                )
            matured += 1

        base["matured"] = int(base.get("matured", 0)) + matured
        base["skipped_invalid"] = int(base.get("skipped_invalid", 0)) + skipped
        base["pending"] = int(base.get("pending", 0)) + max(0, len(rows) - matured - skipped)
        base["short_horizon_matured"] = matured
        base["short_horizon_pending"] = max(0, len(rows) - matured - skipped)
        return base

    cl.mature_predictions = mature_predictions

    original_cycle = cl.run_continual_learning_cycle

    def run_continual_learning_cycle(store, benchmark: str = cl.DEFAULT_BENCHMARK) -> dict:
        cl.ensure_learning_schema(store)
        try:
            short = generate_short_horizon_predictions(store, benchmark)
        except Exception as exc:
            short = {"status": "ERROR", "journaled": 0, "error": f"{type(exc).__name__}: {exc}"}
        out = original_cycle(store, benchmark)
        out["short_horizon"] = short
        out["short_horizon_journaled"] = int(short.get("journaled", 0) or 0)
        try:
            store.set_state("continual_learning", "last_cycle", out)
        except Exception:
            pass
        return out

    cl.run_continual_learning_cycle = run_continual_learning_cycle
    _INSTALLED = True
