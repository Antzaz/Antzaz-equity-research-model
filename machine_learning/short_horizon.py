from __future__ import annotations

"""Short-horizon expected-return research models.

This module extends the existing governed return stack with genuinely short trading horizons
while preserving the current 12-month model and portfolio-optimizer contract.

- 1D and 1W models use fast market features derived only from information available at the
  decision date (price, volume, momentum, volatility, drawdown, beta and market context).
- 1M/3M/6M models continue to use the project's point-in-time fundamental/market feature store.
- Every forecast is journaled for later maturation and champion/challenger evaluation.
- None of these research-only horizons execute trades or overwrite valuation assumptions.
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
from .models import EXPECTED_FEATURES
from .validation import expanding_walk_forward


FAST_FEATURES = [
    "ret_1d",
    "ret_5d",
    "ret_21d",
    "excess_1d",
    "excess_5d",
    "excess_21d",
    "vol_5d",
    "vol_21d",
    "vol_63d",
    "drawdown_21d",
    "drawdown_63d",
    "beta_63d",
    "volume_z_21d",
    "market_ret_1d",
    "market_ret_5d",
    "market_vol_21d",
]

SHORT_HORIZONS = {
    "Expected 1D Excess Return": {
        "days": 1,
        "trading_days": 1,
        "target_type": "1d_excess_return",
        "label": "1D",
        "lookback": 1,
        "feature_mode": "fast",
        "journal_spacing_days": 1,
    },
    "Expected 1W Excess Return": {
        "days": 7,
        "trading_days": 5,
        "target_type": "1w_excess_return",
        "label": "1W",
        "lookback": 5,
        "feature_mode": "fast",
        "journal_spacing_days": 1,
    },
    "Expected 1M Excess Return": {
        "days": 30,
        "target_type": "1m_excess_return",
        "label": "1M",
        "lookback": 21,
        "feature_mode": "fundamental",
        "journal_spacing_days": 7,
    },
    "Expected 3M Excess Return": {
        "days": 91,
        "target_type": "3m_excess_return",
        "label": "3M",
        "lookback": 63,
        "feature_mode": "fundamental",
        "journal_spacing_days": 14,
    },
    "Expected 6M Excess Return": {
        "days": 182,
        "target_type": "6m_excess_return",
        "label": "6M",
        "lookback": 126,
        "feature_mode": "fundamental",
        "journal_spacing_days": 30,
    },
}
SHORT_HORIZON_MODEL_VERSION = "ml-short-horizon-v2-fast"


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


def _price_frame(store, symbol: str) -> pd.DataFrame:
    frames = []
    for candidate in _symbol_candidates(symbol):
        try:
            pf = store.price_frame(candidate)
        except Exception:
            pf = pd.DataFrame()
        if pf is None or pf.empty:
            continue
        out = pf.copy()
        out.index = pd.to_datetime(out.index, errors="coerce")
        out = out[~out.index.isna()]
        if getattr(out.index, "tz", None) is not None:
            out.index = out.index.tz_convert("UTC").tz_localize(None)
        frames.append(out)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames).sort_index()
    return out[~out.index.duplicated(keep="first")]


def _price_series(store, symbol: str) -> pd.Series:
    pf = _price_frame(store, symbol)
    if pf.empty:
        return pd.Series(dtype=float)
    adj = pd.to_numeric(pf.get("adj_close"), errors="coerce")
    close = pd.to_numeric(pf.get("close"), errors="coerce")
    px = adj.fillna(close) if adj is not None else close
    return px.dropna() if px is not None else pd.Series(dtype=float)


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


def _forward_excess_trading(stock: pd.Series, bench: pd.Series, as_of: pd.Timestamp, trading_days: int):
    aligned = pd.concat([stock.rename("stock"), bench.rename("bench")], axis=1, join="inner").dropna()
    if aligned.empty:
        return None, None
    aligned = aligned.loc[aligned.index >= as_of]
    n = int(trading_days)
    if len(aligned) <= n:
        return None, None
    s0 = _finite(aligned.iloc[0]["stock"])
    b0 = _finite(aligned.iloc[0]["bench"])
    s1 = _finite(aligned.iloc[n]["stock"])
    b1 = _finite(aligned.iloc[n]["bench"])
    if None in (s0, s1, b0, b1) or s0 == 0 or b0 == 0:
        return None, None
    realized = (s1 / s0 - 1.0) - (b1 / b0 - 1.0)
    return float(realized), aligned.index[n]


def _trailing_excess(stock: pd.Series, bench: pd.Series, as_of: pd.Timestamp, observations: int) -> float:
    n = max(1, int(observations))
    aligned = pd.concat([stock.rename("stock"), bench.rename("bench")], axis=1, join="inner").dropna()
    aligned = aligned.loc[aligned.index <= as_of].tail(n + 1)
    if len(aligned) < 2:
        return 0.0
    s0, s1 = _finite(aligned.iloc[0]["stock"]), _finite(aligned.iloc[-1]["stock"])
    b0, b1 = _finite(aligned.iloc[0]["bench"]), _finite(aligned.iloc[-1]["bench"])
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


def _feature_symbols(store) -> list[str]:
    df = _feature_history(store)
    if not df.empty:
        return sorted(df["symbol"].astype(str).str.upper().unique().tolist())
    try:
        return [str(x).upper() for x in store.symbols()]
    except Exception:
        return []


def build_horizon_training_frame(store, benchmark: str, horizon_days: int) -> pd.DataFrame:
    """Derive forward excess-return targets for the existing PIT fundamental feature rows."""
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


def _fast_symbol_frame(store, symbol: str, benchmark: str) -> pd.DataFrame:
    pf = _price_frame(store, symbol)
    bf = _price_frame(store, benchmark)
    if pf.empty or bf.empty:
        return pd.DataFrame()

    stock = pd.to_numeric(pf.get("adj_close"), errors="coerce")
    if stock is None or stock.dropna().empty:
        stock = pd.to_numeric(pf.get("close"), errors="coerce")
    else:
        stock = stock.fillna(pd.to_numeric(pf.get("close"), errors="coerce"))
    bench = pd.to_numeric(bf.get("adj_close"), errors="coerce")
    if bench is None or bench.dropna().empty:
        bench = pd.to_numeric(bf.get("close"), errors="coerce")
    else:
        bench = bench.fillna(pd.to_numeric(bf.get("close"), errors="coerce"))
    volume = pd.to_numeric(pf.get("volume"), errors="coerce") if "volume" in pf else pd.Series(index=pf.index, dtype=float)

    df = pd.concat(
        [stock.rename("stock"), bench.rename("bench"), volume.rename("volume")],
        axis=1,
        join="inner",
    ).sort_index()
    df = df.dropna(subset=["stock", "bench"])
    if len(df) < 80:
        return pd.DataFrame()

    sr = df["stock"].pct_change()
    br = df["bench"].pct_change()
    df["ret_1d"] = df["stock"].pct_change(1)
    df["ret_5d"] = df["stock"].pct_change(5)
    df["ret_21d"] = df["stock"].pct_change(21)
    df["excess_1d"] = sr - br
    df["excess_5d"] = df["stock"].pct_change(5) - df["bench"].pct_change(5)
    df["excess_21d"] = df["stock"].pct_change(21) - df["bench"].pct_change(21)
    df["vol_5d"] = sr.rolling(5).std() * np.sqrt(252)
    df["vol_21d"] = sr.rolling(21).std() * np.sqrt(252)
    df["vol_63d"] = sr.rolling(63).std() * np.sqrt(252)
    df["drawdown_21d"] = df["stock"] / df["stock"].rolling(21).max() - 1.0
    df["drawdown_63d"] = df["stock"] / df["stock"].rolling(63).max() - 1.0
    cov = sr.rolling(63).cov(br)
    var = br.rolling(63).var().replace(0, np.nan)
    df["beta_63d"] = cov / var
    vol_mean = df["volume"].rolling(21).mean()
    vol_std = df["volume"].rolling(21).std().replace(0, np.nan)
    df["volume_z_21d"] = (df["volume"] - vol_mean) / vol_std
    df["market_ret_1d"] = br
    df["market_ret_5d"] = df["bench"].pct_change(5)
    df["market_vol_21d"] = br.rolling(21).std() * np.sqrt(252)
    df["symbol"] = str(symbol).upper()
    df["as_of"] = df.index
    return df.replace([np.inf, -np.inf], np.nan)


def build_fast_training_frame(
    store,
    benchmark: str,
    trading_days: int,
    *,
    max_rows: int = 60000,
) -> pd.DataFrame:
    """Build a cross-sectional daily PIT panel for 1D/1W excess-return research."""
    frames = []
    td = int(trading_days)
    for symbol in _feature_symbols(store):
        if symbol == benchmark.upper():
            continue
        df = _fast_symbol_frame(store, symbol, benchmark)
        if df.empty:
            continue
        df["target_excess_return"] = (
            df["stock"].shift(-td) / df["stock"] - 1.0
            - (df["bench"].shift(-td) / df["bench"] - 1.0)
        )
        target_dates = pd.Series(df.index, index=df.index).shift(-td)
        df["target_date"] = pd.to_datetime(target_dates, errors="coerce")
        use = df[["symbol", "as_of", "target_date", "target_excess_return"] + FAST_FEATURES].copy()
        use = use.dropna(subset=["target_excess_return", "target_date"])
        frames.append(use)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True).sort_values(["as_of", "symbol"])
    for c in FAST_FEATURES + ["target_excess_return"]:
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


def current_fast_feature_frame(store, benchmark: str) -> pd.DataFrame:
    rows = []
    for symbol in _feature_symbols(store):
        if symbol == benchmark.upper():
            continue
        df = _fast_symbol_frame(store, symbol, benchmark)
        if df.empty:
            continue
        row = df.tail(1).iloc[0]
        out = {c: _finite(row.get(c)) for c in FAST_FEATURES}
        out.update({"symbol": str(symbol).upper(), "as_of": pd.Timestamp(row["as_of"])})
        rows.append(out)
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    for c in FAST_FEATURES:
        frame[c] = pd.to_numeric(frame[c], errors="coerce")
    return frame.replace([np.inf, -np.inf], np.nan)


def current_feature_frame(store, benchmark: str) -> pd.DataFrame:
    """Build current PIT fundamental vectors with freshly calculated market features."""
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


def model_feature_columns(model_name: str) -> list[str]:
    spec = SHORT_HORIZONS[model_name]
    return FAST_FEATURES if spec.get("feature_mode") == "fast" else EXPECTED_FEATURES


def training_frame_for_model(store, model_name: str, benchmark: str = "SPY") -> pd.DataFrame:
    spec = SHORT_HORIZONS[model_name]
    if spec.get("feature_mode") == "fast":
        return build_fast_training_frame(store, benchmark, int(spec["trading_days"]))
    return build_horizon_training_frame(store, benchmark, int(spec["days"]))


def current_frame_for_model(store, model_name: str, benchmark: str = "SPY") -> pd.DataFrame:
    spec = SHORT_HORIZONS[model_name]
    if spec.get("feature_mode") == "fast":
        return current_fast_feature_frame(store, benchmark)
    return current_feature_frame(store, benchmark)


def _estimators(include_tree_challenger: bool = False):
    hgb = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("model", HistGradientBoostingRegressor(
            max_iter=180,
            max_leaf_nodes=12,
            learning_rate=0.05,
            l2_regularization=0.5,
            random_state=RANDOM_STATE,
        )),
    ])
    elastic = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("scale", StandardScaler()),
        ("model", ElasticNet(alpha=0.02, l1_ratio=0.25, max_iter=5000, random_state=RANDOM_STATE)),
    ])
    if not include_tree_challenger:
        return {"hgb": hgb, "elastic": elastic}
    extra = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("model", ExtraTreesRegressor(
            n_estimators=180,
            min_samples_leaf=8,
            max_features=0.8,
            n_jobs=-1,
            random_state=RANDOM_STATE,
        )),
    ])
    return {"hgb": hgb, "elastic": elastic, "extra_trees": extra}


def _confidence(n: int, fast: bool = False) -> str:
    if fast:
        return "High" if n >= 5000 else "Moderate" if n >= 1500 else "Low"
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


def _ensemble_weights(names: list[str], fast: bool) -> dict[str, float]:
    if fast and set(names) >= {"hgb", "elastic", "extra_trees"}:
        return {"hgb": 0.45, "extra_trees": 0.35, "elastic": 0.20}
    return {"hgb": 0.65, "elastic": 0.35}


def generate_short_horizon_predictions(
    store,
    benchmark: str = "SPY",
    *,
    min_training_rows: int = 40,
    journal_spacing_days: int | None = None,
) -> dict:
    """Train 1D/1W/1M/3M/6M models and journal current forecasts for live grading."""
    today = datetime.now(timezone.utc).date().isoformat()
    try:
        last = store.get_state("short_horizon_learning", "last_generation", {}) or {}
        if isinstance(last, dict) and last.get("date") == today:
            return {"status": "ALREADY_RAN_TODAY", "journaled": 0, "models": last.get("models", {})}
    except Exception:
        pass

    bench = _price_series(store, benchmark)
    if bench.empty:
        return {"status": "NO_BENCHMARK_HISTORY", "journaled": 0, "models": {}}

    total_journaled = 0
    model_summary: dict[str, dict] = {}
    for model_name, spec in SHORT_HORIZONS.items():
        fast = spec.get("feature_mode") == "fast"
        features = model_feature_columns(model_name)
        train = training_frame_for_model(store, model_name, benchmark)
        current = current_frame_for_model(store, model_name, benchmark)
        required = max(int(min_training_rows), 800 if fast else int(min_training_rows))
        if len(train) < required or current.empty:
            model_summary[model_name] = {
                "status": "INSUFFICIENT_DATA",
                "training_rows": int(len(train)),
                "current_rows": int(len(current)),
                "minimum_rows": int(required),
                "journaled": 0,
            }
            continue

        estimators = _estimators(include_tree_challenger=fast)
        min_train = max(500 if fast else 30, min(5000 if fast else 80, len(train) // 2))
        step = max(1, len(train) // (18 if fast else 25))
        walk_forward = {}
        fitted = {}
        predictions = {}
        for name, estimator in estimators.items():
            wf = expanding_walk_forward(
                estimator,
                train,
                features,
                "target_excess_return",
                min_train=min_train,
                step=step,
            )
            estimator.fit(train[features], train["target_excess_return"])
            walk_forward[name] = wf.metrics
            fitted[name] = estimator
            predictions[name] = np.asarray(estimator.predict(current[features]), dtype=float)

        weights = _ensemble_weights(list(fitted), fast)
        pred = np.zeros(len(current), dtype=float)
        total_weight = 0.0
        for name, weight in weights.items():
            if name in predictions:
                pred += float(weight) * predictions[name]
                total_weight += float(weight)
        if total_weight <= 0:
            model_summary[model_name] = {"status": "ERROR", "training_rows": int(len(train)), "journaled": 0}
            continue
        pred /= total_weight

        confidence = _confidence(len(train), fast=fast)
        created_at = _now()
        journaled = 0
        spacing = int(spec.get("journal_spacing_days", 1)) if journal_spacing_days is None else int(journal_spacing_days)

        for i, (_, row) in enumerate(current.iterrows()):
            symbol = str(row["symbol"]).upper()
            as_of = pd.Timestamp(row["as_of"])
            if not _journal_spacing_ok(store, model_name, symbol, as_of, spacing):
                continue
            stock = _price_series(store, symbol)
            baseline = _trailing_excess(stock, bench, as_of, int(spec["lookback"]))
            features_payload = {
                "benchmark": benchmark.upper(),
                "horizon": spec["label"],
                "horizon_days": int(spec["days"]),
                "trading_days": int(spec.get("trading_days", 0) or 0),
                "feature_mode": spec.get("feature_mode"),
                "feature_snapshot": {c: _finite(row.get(c)) for c in features},
                "walk_forward": walk_forward,
                "ensemble_weights": weights,
            }
            with store.connect() as con:
                con.execute(
                    """INSERT INTO predictions(
                           run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,
                           model_version,realized_at,realized_value,error,created_at,target_type,
                           baseline_value,baseline_name,evaluation_version)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        f"short-{today}",
                        symbol,
                        model_name,
                        as_of.isoformat(),
                        int(spec["days"]),
                        json.dumps(float(pred[i])),
                        confidence,
                        json.dumps(features_payload, default=str),
                        SHORT_HORIZON_MODEL_VERSION,
                        None,
                        None,
                        None,
                        created_at,
                        spec["target_type"],
                        float(baseline),
                        f"Trailing {spec['label']} excess-return baseline",
                        None,
                    ),
                )
            journaled += 1

        validation = {
            "horizon": spec["label"],
            "horizon_days": int(spec["days"]),
            "trading_days": int(spec.get("trading_days", 0) or 0),
            "feature_mode": spec.get("feature_mode"),
            "training_rows": int(len(train)),
            "walk_forward": walk_forward,
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
                    SHORT_HORIZON_MODEL_VERSION,
                    created_at,
                    "PASS",
                    confidence,
                    json.dumps({"median_prediction": float(np.nanmedian(pred))}),
                    int(len(train)),
                    json.dumps(validation, default=str),
                    "[]",
                    json.dumps({
                        "horizon_days": int(spec["days"]),
                        "trading_days": int(spec.get("trading_days", 0) or 0),
                        "portfolio_use": "research_only",
                    }),
                ),
            )
        total_journaled += journaled
        model_summary[model_name] = {
            "status": "PASS",
            "training_rows": int(len(train)),
            "journaled": int(journaled),
            "walk_forward": walk_forward,
            "ensemble_weights": weights,
        }

    summary = {"status": "PASS", "date": today, "journaled": int(total_journaled), "models": model_summary}
    try:
        store.set_state("short_horizon_learning", "last_generation", summary)
    except Exception:
        pass
    return summary
