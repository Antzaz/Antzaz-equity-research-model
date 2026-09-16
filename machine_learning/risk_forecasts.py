from __future__ import annotations

"""Research-only forward risk forecasts built on the fast point-in-time market panel.

These models complement expected-return forecasts with two questions that matter directly to
portfolio construction:

1. What realized volatility is plausible over the next trading week?
2. What peak-to-trough loss from today's price is plausible over the next month?

They remain context-only: they do not execute trades or directly overwrite optimizer inputs.
The deterministic validation evidence is persisted so the research loop can challenge them.
"""

from datetime import datetime, timezone
import json
import math
from typing import Any

import numpy as np
import pandas as pd
from sklearn.ensemble import ExtraTreesRegressor, HistGradientBoostingRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import ElasticNet
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from .common import RANDOM_STATE
from .short_horizon import FAST_FEATURES, _fast_symbol_frame, _feature_symbols, _finite
from .validation import expanding_walk_forward


RISK_MODEL_VERSION = "ml-forward-risk-v1"
RISK_MODELS = {
    "Expected 1W Realized Volatility": {
        "target": "target_volatility",
        "trading_days": 5,
        "label": "1W realized volatility",
    },
    "Expected 1M Forward Drawdown": {
        "target": "target_drawdown",
        "trading_days": 21,
        "label": "1M forward drawdown",
    },
}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _estimators() -> dict[str, Any]:
    return {
        "hgb": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", HistGradientBoostingRegressor(
                max_iter=180,
                max_leaf_nodes=12,
                learning_rate=0.05,
                l2_regularization=0.5,
                random_state=RANDOM_STATE,
            )),
        ]),
        "extra_trees": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", ExtraTreesRegressor(
                n_estimators=180,
                min_samples_leaf=8,
                max_features=0.8,
                n_jobs=-1,
                random_state=RANDOM_STATE,
            )),
        ]),
        "elastic": Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scale", StandardScaler()),
            ("model", ElasticNet(alpha=0.01, l1_ratio=0.20, max_iter=5000, random_state=RANDOM_STATE)),
        ]),
    }


def _future_volatility(returns: pd.Series, n: int) -> pd.Series:
    shifted = pd.to_numeric(returns, errors="coerce").shift(-1)
    return shifted.iloc[::-1].rolling(int(n)).std().iloc[::-1] * np.sqrt(252)


def _future_drawdown(prices: pd.Series, n: int) -> pd.Series:
    px = pd.to_numeric(prices, errors="coerce")
    future_min = px.shift(-1).iloc[::-1].rolling(int(n)).min().iloc[::-1]
    # Drawdown is a loss measure. If every future observation is above today's price,
    # the economically correct forward drawdown is 0%, not a positive return.
    return (future_min / px - 1.0).clip(upper=0.0)


def build_risk_training_frame(store, benchmark: str = "SPY", *, max_rows: int = 60000) -> pd.DataFrame:
    frames = []
    for symbol in _feature_symbols(store):
        if symbol == benchmark.upper():
            continue
        df = _fast_symbol_frame(store, symbol, benchmark)
        if df.empty:
            continue
        sr = pd.to_numeric(df["stock"], errors="coerce").pct_change()
        df["target_volatility"] = _future_volatility(sr, 5)
        df["target_drawdown"] = _future_drawdown(df["stock"], 21)
        idx = pd.Series(df.index, index=df.index)
        df["target_date_volatility"] = pd.to_datetime(idx.shift(-5), errors="coerce")
        df["target_date_drawdown"] = pd.to_datetime(idx.shift(-21), errors="coerce")
        frames.append(df[["symbol", "as_of", "target_volatility", "target_drawdown",
                          "target_date_volatility", "target_date_drawdown"] + FAST_FEATURES].copy())
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True).sort_values(["as_of", "symbol"])
    for c in FAST_FEATURES + ["target_volatility", "target_drawdown"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out.replace([np.inf, -np.inf], np.nan)
    if len(out) > int(max_rows):
        stride = max(1, int(np.ceil(len(out) / float(max_rows))))
        sampled = out.iloc[::stride].copy()
        recent_cut = out["as_of"].max() - pd.Timedelta(days=120)
        recent = out[out["as_of"] >= recent_cut]
        out = pd.concat([sampled, recent], ignore_index=True).drop_duplicates(["symbol", "as_of"])
        out = out.sort_values(["as_of", "symbol"]).tail(int(max_rows))
    return out.reset_index(drop=True)


def current_risk_features(store, benchmark: str = "SPY") -> pd.DataFrame:
    rows = []
    for symbol in _feature_symbols(store):
        if symbol == benchmark.upper():
            continue
        df = _fast_symbol_frame(store, symbol, benchmark)
        if df.empty:
            continue
        row = df.tail(1).iloc[0]
        item = {c: _finite(row.get(c)) for c in FAST_FEATURES}
        item.update({"symbol": symbol, "as_of": pd.Timestamp(row["as_of"])})
        rows.append(item)
    return pd.DataFrame(rows).replace([np.inf, -np.inf], np.nan) if rows else pd.DataFrame()


def _confidence(n: int) -> str:
    return "High" if n >= 5000 else "Moderate" if n >= 1500 else "Low"


def generate_risk_forecasts(store, benchmark: str = "SPY", *, min_rows: int = 800) -> dict:
    today = datetime.now(timezone.utc).date().isoformat()
    try:
        last = store.get_state("risk_forecasts", "last_generation", {}) or {}
        if isinstance(last, dict) and last.get("date") == today:
            return {"status": "ALREADY_RAN_TODAY", "date": today, "models": last.get("models", {})}
    except Exception:
        pass

    train = build_risk_training_frame(store, benchmark)
    current = current_risk_features(store, benchmark)
    if len(train) < int(min_rows) or current.empty:
        return {
            "status": "INSUFFICIENT_DATA",
            "date": today,
            "training_rows": int(len(train)),
            "current_rows": int(len(current)),
            "models": {},
        }

    summary: dict[str, Any] = {"status": "PASS", "date": today, "models": {}}
    for model_name, spec in RISK_MODELS.items():
        target = spec["target"]
        target_date_col = "target_date_volatility" if target == "target_volatility" else "target_date_drawdown"
        frame = train.dropna(subset=[target, target_date_col]).copy()
        frame["target_date"] = frame[target_date_col]
        if len(frame) < int(min_rows):
            summary["models"][model_name] = {"status": "INSUFFICIENT_DATA", "training_rows": int(len(frame))}
            continue

        estimators = _estimators()
        min_train = max(500, min(5000, len(frame) // 2))
        step = max(1, len(frame) // 18)
        validation = {}
        preds = {}
        for name, estimator in estimators.items():
            wf = expanding_walk_forward(
                estimator,
                frame,
                FAST_FEATURES,
                target,
                min_train=min_train,
                step=step,
            )
            estimator.fit(frame[FAST_FEATURES], frame[target])
            validation[name] = wf.metrics
            preds[name] = np.asarray(estimator.predict(current[FAST_FEATURES]), dtype=float)

        weights = {"hgb": 0.45, "extra_trees": 0.35, "elastic": 0.20}
        pred = sum(float(weights[name]) * preds[name] for name in weights)
        if target == "target_volatility":
            pred = np.clip(pred, 0.0, None)
        else:
            pred = np.clip(pred, -1.0, 0.0)

        per_symbol = {
            str(row["symbol"]): float(pred[i])
            for i, (_, row) in enumerate(current.iterrows())
            if math.isfinite(float(pred[i]))
        }
        created_at = _now()
        details = {
            "benchmark": benchmark.upper(),
            "label": spec["label"],
            "trading_days": int(spec["trading_days"]),
            "training_rows": int(len(frame)),
            "walk_forward": validation,
            "ensemble_weights": weights,
            "portfolio_use": "research_only",
        }
        with store.connect() as con:
            con.execute(
                """INSERT INTO training_runs(
                       model,model_version,trained_at,status,confidence,prediction,training_rows,
                       validation_json,drivers_json,details_json)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (
                    model_name,
                    RISK_MODEL_VERSION,
                    created_at,
                    "PASS",
                    _confidence(len(frame)),
                    json.dumps({"median_prediction": float(np.nanmedian(pred))}),
                    int(len(frame)),
                    json.dumps(validation, default=str),
                    "[]",
                    json.dumps(details, default=str),
                ),
            )
        summary["models"][model_name] = {
            "status": "PASS",
            "training_rows": int(len(frame)),
            "median_prediction": float(np.nanmedian(pred)),
            "walk_forward": validation,
            "per_symbol": per_symbol,
        }

    try:
        store.set_state("risk_forecasts", "last_generation", summary)
    except Exception:
        pass
    return summary
