from __future__ import annotations

"""Short-horizon expected-return research models.

These models are deliberately separate from the existing 12-month expected-return model.
They use the same point-in-time feature family but learn genuinely different forward
excess-return targets at 1M, 3M and 6M horizons.  Their predictions are journaled for
continual-learning evaluation, but they do not feed the portfolio optimizer.
"""

from datetime import datetime, timezone
import json
import math
from typing import Any

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import ElasticNet
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from .models import EXPECTED_FEATURES
from .validation import expanding_walk_forward


SHORT_HORIZONS = {
    "Expected 1M Excess Return": {"days": 30, "target_type": "1m_excess_return", "label": "1M", "lookback": 21},
    "Expected 3M Excess Return": {"days": 91, "target_type": "3m_excess_return", "label": "3M", "lookback": 63},
    "Expected 6M Excess Return": {"days": 182, "target_type": "6m_excess_return", "label": "6M", "lookback": 126},
}
SHORT_HORIZON_MODEL_VERSION = "ml-short-horizon-v1"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _finite(value: Any, default=None):
    try:
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


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


def _price_series(store, symbol: str) -> pd.Series:
    frames = []
    for candidate in _symbol_candidates(symbol):
        try:
            pf = store.price_frame(candidate)
        except Exception:
            pf = pd.DataFrame()
        if pf is None or pf.empty:
            continue
        adj = pd.to_numeric(pf.get("adj_close"), errors="coerce")
        close = pd.to_numeric(pf.get("close"), errors="coerce")
        px = adj.fillna(close) if adj is not None else close
        if px is not None and not px.dropna().empty:
            frames.append(px.dropna())
    if not frames:
        return pd.Series(dtype=float)
    out = pd.concat(frames).sort_index()
    out.index = pd.to_datetime(out.index, errors="coerce")
    out = out[~out.index.isna()]
    if getattr(out.index, "tz", None) is not None:
        out.index = out.index.tz_convert("UTC").tz_localize(None)
    return out[~out.index.duplicated(keep="first")]


def _first_at_or_after(series: pd.Series, when: pd.Timestamp):
    if series.empty:
        return None, None
    view = series.loc[series.index >= when]
    if view.empty:
        return None, None
    return view.index[0], _finite(view.iloc[0])


def _forward_excess(stock: pd.Series, bench: pd.Series, as_of: pd.Timestamp, days: int):
    target = as_of + pd.Timedelta(days=int(days))
    _, s0 = _first_at_or_after(stock, as_of)
    sd, s1 = _first_at_or_after(stock, target)
    _, b0 = _first_at_or_after(bench, as_of)
    bd, b1 = _first_at_or_after(bench, target)
    if None in (s0, s1, b0, b1) or s0 == 0 or b0 == 0:
        return None, None
    realized = (s1 / s0 - 1.0) - (b1 / b0 - 1.0)
    target_date = max(sd, bd) if sd is not None and bd is not None else sd or bd
    return float(realized), target_date


def _trailing_excess(stock: pd.Series, bench: pd.Series, as_of: pd.Timestamp, observations: int) -> float:
    s = stock.loc[stock.index <= as_of].dropna().tail(int(observations) + 1)
    b = bench.loc[bench.index <= as_of].dropna().tail(int(observations) + 1)
    if len(s) < max(10, observations // 2) or len(b) < max(10, observations // 2):
        return 0.0
    s0, s1 = _finite(s.iloc[0]), _finite(s.iloc[-1])
    b0, b1 = _finite(b.iloc[0]), _finite(b.iloc[-1])
    if None in (s0, s1, b0, b1) or s0 == 0 or b0 == 0:
        return 0.0
    return float((s1 / s0 - 1.0) - (b1 / b0 - 1.0))


def _feature_history(store) -> pd.DataFrame:
    try:
        with store.connect() as con:
            df = pd.read_sql_query("SELECT * FROM features ORDER BY as_of,symbol", con)
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return df
    df["as_of"] = pd.to_datetime(df["as_of"], errors="coerce")
    return df.dropna(subset=["symbol", "as_of"])


def build_horizon_training_frame(store, benchmark: str, horizon_days: int) -> pd.DataFrame:
    """Derive a true forward excess-return target for each stored point-in-time feature row."""
    features = _feature_history(store)
    if features.empty:
        return pd.DataFrame()
    bench = _price_series(store, benchmark)
    if bench.empty:
        return pd.DataFrame()

    rows = []
    for symbol, group in features.groupby("symbol", sort=False):
        stock = _price_series(store, symbol)
        if stock.empty:
            continue
        for _, row in group.iterrows():
            as_of = pd.Timestamp(row["as_of"])
            target, target_date = _forward_excess(stock, bench, as_of, horizon_days)
            if target is None or target_date is None:
                continue
            out = {c: row.get(c) for c in EXPECTED_FEATURES}
            out.update({
                "symbol": str(symbol).upper(),
                "as_of": as_of,
                "target_date": target_date,
                "target_excess_return": float(target),
            })
            rows.append(out)
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    for c in EXPECTED_FEATURES + ["target_excess_return"]:
        frame[c] = pd.to_numeric(frame[c], errors="coerce")
    return frame.replace([np.inf, -np.inf], np.nan).dropna(subset=["target_excess_return"])


def _drawdown(prices: pd.Series) -> float | None:
    x = pd.to_numeric(prices, errors="coerce").dropna()
    if len(x) < 2:
        return None
    dd = x / x.cummax() - 1.0
    return _finite(dd.min())


def current_feature_frame(store, benchmark: str) -> pd.DataFrame:
    """Build current feature vectors from latest PIT fundamentals plus current market features."""
    try:
        with store.connect() as con:
            latest = pd.read_sql_query(
                """SELECT f.* FROM features f
                   JOIN (SELECT symbol,MAX(as_of) AS as_of FROM features GROUP BY symbol) x
                   ON f.symbol=x.symbol AND f.as_of=x.as_of
                   ORDER BY f.symbol""",
                con,
            )
    except Exception:
        return pd.DataFrame()
    if latest.empty:
        return latest
    bench = _price_series(store, benchmark)
    if bench.empty:
        return pd.DataFrame()
    bench_last = bench.index.max()

    rows = []
    for _, row in latest.iterrows():
        symbol = str(row.get("symbol") or "").upper().strip()
        stock = _price_series(store, symbol)
        if not symbol or stock.empty:
            continue
        as_of = min(stock.index.max(), bench_last)
        stock_view = stock.loc[stock.index <= as_of]
        p12 = stock_view.tail(253)
        p6 = stock_view.tail(127)
        out = {c: row.get(c) for c in EXPECTED_FEATURES}
        if len(p12) >= 126:
            out["momentum_12m"] = float(p12.iloc[-1] / p12.iloc[0] - 1.0)
            out["drawdown_12m"] = _drawdown(p12)
        if len(p6) >= 63:
            out["momentum_6m"] = float(p6.iloc[-1] / p6.iloc[0] - 1.0)
            out["volatility_6m"] = float(p6.pct_change().std() * np.sqrt(252))
        out.update({"symbol": symbol, "as_of": as_of})
        rows.append(out)
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    for c in EXPECTED_FEATURES:
        frame[c] = pd.to_numeric(frame[c], errors="coerce")
    return frame.replace([np.inf, -np.inf], np.nan)


def _estimators():
    hgb = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("model", HistGradientBoostingRegressor(
            max_iter=180, max_leaf_nodes=12, learning_rate=0.05,
            l2_regularization=0.5, random_state=42,
        )),
    ])
    elastic = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("scale", StandardScaler()),
        ("model", ElasticNet(alpha=0.02, l1_ratio=0.25, max_iter=5000, random_state=42)),
    ])
    return hgb, elastic


def _confidence(n: int) -> str:
    return "High" if n >= 120 else "Moderate" if n >= 60 else "Low"


def _journal_spacing_ok(store, model: str, symbol: str, as_of: pd.Timestamp, spacing_days: int) -> bool:
    try:
        with store.connect() as con:
            row = con.execute(
                """SELECT as_of FROM predictions WHERE model=? AND symbol=? AND model_version=?
                   ORDER BY datetime(created_at) DESC,id DESC LIMIT 1""",
                (model, symbol, SHORT_HORIZON_MODEL_VERSION),
            ).fetchone()
        if not row:
            return True
        last = pd.Timestamp(row[0])
        if last.tzinfo is not None:
            last = last.tz_convert("UTC").tz_localize(None)
        return (as_of - last).days >= int(spacing_days)
    except Exception:
        return True


def generate_short_horizon_predictions(
    store,
    benchmark: str = "SPY",
    *,
    min_training_rows: int = 40,
    journal_spacing_days: int = 7,
) -> dict:
    """Train 1M/3M/6M models and journal current forecasts for later live evaluation."""
    today = datetime.now(timezone.utc).date().isoformat()
    try:
        last = store.get_state("short_horizon_learning", "last_generation", {}) or {}
        if isinstance(last, dict) and last.get("date") == today:
            return {"status": "ALREADY_RAN_TODAY", "journaled": 0, "models": last.get("models", {})}
    except Exception:
        pass

    current = current_feature_frame(store, benchmark)
    if current.empty:
        return {"status": "NO_CURRENT_FEATURES", "journaled": 0, "models": {}}
    bench = _price_series(store, benchmark)
    if bench.empty:
        return {"status": "NO_BENCHMARK_HISTORY", "journaled": 0, "models": {}}

    total_journaled = 0
    model_summary: dict[str, dict] = {}
    for model_name, spec in SHORT_HORIZONS.items():
        train = build_horizon_training_frame(store, benchmark, int(spec["days"]))
        if len(train) < int(min_training_rows):
            model_summary[model_name] = {"status": "INSUFFICIENT_DATA", "training_rows": int(len(train)), "journaled": 0}
            continue

        hgb, elastic = _estimators()
        min_train = max(30, min(80, len(train) // 2))
        step = max(1, len(train) // 25)
        wf_h = expanding_walk_forward(
            hgb, train, EXPECTED_FEATURES, "target_excess_return",
            min_train=min_train, step=step,
        )
        wf_e = expanding_walk_forward(
            elastic, train, EXPECTED_FEATURES, "target_excess_return",
            min_train=min_train, step=step,
        )
        hgb.fit(train[EXPECTED_FEATURES], train["target_excess_return"])
        elastic.fit(train[EXPECTED_FEATURES], train["target_excess_return"])
        p_h = hgb.predict(current[EXPECTED_FEATURES])
        p_e = elastic.predict(current[EXPECTED_FEATURES])
        pred = 0.65 * np.asarray(p_h, dtype=float) + 0.35 * np.asarray(p_e, dtype=float)
        confidence = _confidence(len(train))
        created_at = _now()
        journaled = 0

        for i, (_, row) in enumerate(current.iterrows()):
            symbol = str(row["symbol"]).upper()
            as_of = pd.Timestamp(row["as_of"])
            if not _journal_spacing_ok(store, model_name, symbol, as_of, journal_spacing_days):
                continue
            stock = _price_series(store, symbol)
            baseline = _trailing_excess(stock, bench, as_of, int(spec["lookback"]))
            features_payload = {
                "benchmark": benchmark.upper(),
                "horizon": spec["label"],
                "horizon_days": int(spec["days"]),
                "feature_snapshot": {c: _finite(row.get(c)) for c in EXPECTED_FEATURES},
                "walk_forward": {"hgb": wf_h.metrics, "elastic": wf_e.metrics},
            }
            with store.connect() as con:
                con.execute(
                    """INSERT INTO predictions(
                           run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,
                           model_version,realized_at,realized_value,error,created_at,target_type,
                           baseline_value,baseline_name,evaluation_version)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        f"short-{today}", symbol, model_name, as_of.isoformat(), int(spec["days"]),
                        json.dumps(float(pred[i])), confidence, json.dumps(features_payload, default=str),
                        SHORT_HORIZON_MODEL_VERSION, None, None, None, created_at,
                        spec["target_type"], float(baseline),
                        f"Trailing {spec['label']} excess-return baseline", None,
                    ),
                )
            journaled += 1

        validation = {
            "horizon": spec["label"],
            "horizon_days": int(spec["days"]),
            "training_rows": int(len(train)),
            "hgb_walk_forward": wf_h.metrics,
            "elastic_walk_forward": wf_e.metrics,
            "ensemble_weights": {"hgb": 0.65, "elastic": 0.35},
            "portfolio_use": "research_only",
        }
        with store.connect() as con:
            con.execute(
                """INSERT INTO training_runs(
                       model,model_version,trained_at,status,confidence,prediction,training_rows,
                       validation_json,drivers_json,details_json)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (
                    model_name, SHORT_HORIZON_MODEL_VERSION, created_at, "PASS", confidence,
                    json.dumps({"median_prediction": float(np.nanmedian(pred))}), int(len(train)),
                    json.dumps(validation, default=str), "[]",
                    json.dumps({"horizon_days": int(spec["days"]), "portfolio_use": "research_only"}),
                ),
            )
        total_journaled += journaled
        model_summary[model_name] = {
            "status": "PASS", "training_rows": int(len(train)), "journaled": int(journaled),
            "hgb_walk_forward": wf_h.metrics, "elastic_walk_forward": wf_e.metrics,
        }

    summary = {"status": "PASS", "date": today, "journaled": int(total_journaled), "models": model_summary}
    try:
        store.set_state("short_horizon_learning", "last_generation", summary)
    except Exception:
        pass
    return summary
