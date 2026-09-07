from __future__ import annotations

"""Maximum-data classical + ML portfolio construction.

The optimizer keeps prediction and allocation separate:
- public/local history estimates risk;
- classical/manual assumptions estimate a non-ML return prior;
- the existing Expected 12M Excess Return model is used only as a confidence-shrunk input;
- SLSQP solves transparent constrained portfolio problems.

Local ml_data/ml_history.sqlite is preferred because it can contain up to 20 years of
point-in-time history. Public price history supplied by run_research.py is the fallback.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import json
import math
import sqlite3
import sys

import numpy as np
import pandas as pd
from scipy.optimize import minimize
from sklearn.covariance import LedoitWolf


TRADING_DAYS = 252
CONFIDENCE_WEIGHT = {"High": 1.0, "Moderate": 0.65, "Low": 0.30}
ROOT = Path(__file__).resolve().parents[2]


def _finite(value, default=None):
    try:
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _provider_candidates(ticker: str) -> list[str]:
    t = str(ticker or "").upper().strip()
    out = [t]
    if "." in t and t.count(".") == 1:
        left, right = t.split(".")
        if len(right) == 1:
            out.append(f"{left}-{right}")
    return list(dict.fromkeys(x for x in out if x))


def _annualized_geometric(returns: pd.Series, trading_days: int = TRADING_DAYS):
    r = pd.to_numeric(returns, errors="coerce").dropna()
    if len(r) < 60:
        return None
    growth = float((1 + r).prod())
    if growth <= 0:
        return None
    return growth ** (trading_days / len(r)) - 1


def _window_return_prior(r: pd.Series, trading_days: int = TRADING_DAYS):
    """Robust historical prior from several horizons instead of one noisy sample mean."""
    x = pd.to_numeric(r, errors="coerce").dropna()
    if len(x) < 126:
        return None
    estimates = []
    for n in (756, 1260, 2520, 5040):
        view = x.tail(min(n, len(x)))
        value = _annualized_geometric(view, trading_days)
        if value is not None and np.isfinite(value):
            estimates.append(value)
    if not estimates:
        return None
    return float(np.median(estimates))


def _read_price_series_from_db(db_path: Path, ticker: str, max_years: int):
    if not db_path.exists():
        return pd.Series(dtype=float)
    try:
        with sqlite3.connect(db_path) as con:
            frames = []
            for symbol in _provider_candidates(ticker):
                df = pd.read_sql_query(
                    "SELECT date,close,adj_close,source FROM prices WHERE symbol=? ORDER BY date",
                    con,
                    params=(symbol,),
                )
                if not df.empty:
                    frames.append(df)
            if not frames:
                return pd.Series(dtype=float)
        df = pd.concat(frames, ignore_index=True)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        rank = {"Yahoo Finance": 0, "Alpha Vantage": 1}
        df["_rank"] = df["source"].map(rank).fillna(99)
        df = df.sort_values(["date", "_rank"]).drop_duplicates("date", keep="first")
        px = pd.to_numeric(df["adj_close"], errors="coerce").fillna(
            pd.to_numeric(df["close"], errors="coerce")
        )
        s = pd.Series(px.values, index=df["date"]).dropna().sort_index()
        if s.empty:
            return s
        cutoff = s.index.max() - pd.DateOffset(years=int(max_years))
        return s.loc[s.index >= cutoff]
    except Exception:
        return pd.Series(dtype=float)


def maximum_history_returns(
    tickers: list[str],
    fallback_returns: pd.DataFrame,
    history_db: str | Path | None,
    max_years: int = 20,
    min_observations: int = 126,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Use maximum local history per security, falling back to public downloaded history."""
    db_path = Path(history_db) if history_db else Path()
    series = []
    coverage = []
    fallback = fallback_returns.copy()
    fallback.index = pd.to_datetime(fallback.index, errors="coerce")
    for ticker in tickers:
        px = _read_price_series_from_db(db_path, ticker, max_years) if history_db else pd.Series(dtype=float)
        source = "ML SQLite"
        if len(px) >= min_observations:
            r = px.pct_change().dropna().rename(ticker)
        else:
            source = "Public price fallback"
            r = pd.to_numeric(fallback.get(ticker), errors="coerce").dropna()
            if not r.empty:
                cutoff = r.index.max() - pd.DateOffset(years=int(max_years))
                r = r.loc[r.index >= cutoff].rename(ticker)
        if len(r) >= min_observations:
            series.append(r)
        coverage.append({
            "Ticker": ticker,
            "Source": source if len(r) >= min_observations else "Insufficient history",
            "Observations": int(len(r)),
            "Start": r.index.min().date().isoformat() if len(r) else None,
            "End": r.index.max().date().isoformat() if len(r) else None,
            "ApproxYears": float(len(r) / TRADING_DAYS) if len(r) else 0.0,
        })
    panel = pd.concat(series, axis=1).sort_index() if series else pd.DataFrame()
    return panel, pd.DataFrame(coverage)


def _nearest_psd(matrix: np.ndarray, floor: float = 1e-10) -> np.ndarray:
    a = np.asarray(matrix, dtype=float)
    a = (a + a.T) / 2
    vals, vecs = np.linalg.eigh(a)
    vals = np.maximum(vals, floor)
    out = vecs @ np.diag(vals) @ vecs.T
    return (out + out.T) / 2


def robust_covariance(
    returns: pd.DataFrame,
    trading_days: int = TRADING_DAYS,
    diagonal_shrink: float = 0.20,
) -> tuple[pd.DataFrame, dict]:
    """Maximum-history pairwise covariance blended with Ledoit-Wolf common-overlap covariance."""
    r = returns.copy()
    cols = list(r.columns)
    n = len(cols)
    if n < 2:
        return pd.DataFrame(), {"method": "insufficient"}
    pair = np.zeros((n, n), dtype=float)
    pair_obs = []
    for i, a in enumerate(cols):
        xa = pd.to_numeric(r[a], errors="coerce")
        for j, b in enumerate(cols):
            if j < i:
                pair[i, j] = pair[j, i]
                continue
            xb = pd.to_numeric(r[b], errors="coerce")
            xy = pd.concat([xa, xb], axis=1).dropna()
            obs = len(xy)
            if i != j:
                pair_obs.append(obs)
            if obs < 60:
                value = 0.0 if i != j else float(xa.var(ddof=1))
            else:
                value = float(xy.cov().iloc[0, 1])
            pair[i, j] = pair[j, i] = value
    pair *= trading_days
    diag = np.diag(np.diag(pair))
    pair_shrunk = (1 - diagonal_shrink) * pair + diagonal_shrink * diag
    pair_shrunk = _nearest_psd(pair_shrunk)

    complete = r.dropna(how="any")
    blend = 0.0
    if len(complete) >= 126:
        lw = LedoitWolf().fit(complete.values).covariance_ * trading_days
        blend = min(0.45, max(0.20, len(complete) / (5 * trading_days) * 0.20))
        cov = (1 - blend) * pair_shrunk + blend * lw
    else:
        cov = pair_shrunk
    cov = _nearest_psd(cov)
    meta = {
        "method": "20Y pairwise shrinkage + Ledoit-Wolf common overlap",
        "assets": n,
        "complete_overlap_rows": int(len(complete)),
        "median_pairwise_rows": int(np.median(pair_obs)) if pair_obs else int(len(complete)),
        "ledoit_wolf_weight": float(blend),
        "diagonal_shrinkage": float(diagonal_shrink),
    }
    return pd.DataFrame(cov, index=cols, columns=cols), meta


def _parse_json(value):
    if value in (None, ""):
        return None
    try:
        return json.loads(value)
    except Exception:
        return value


def _latest_prediction_rows(db_path: Path, tickers: list[str]) -> pd.DataFrame:
    if not db_path.exists():
        return pd.DataFrame()
    aliases = {alias: t for t in tickers for alias in _provider_candidates(t)}
    placeholders = ",".join(["?"] * len(aliases))
    if not placeholders:
        return pd.DataFrame()
    try:
        with sqlite3.connect(db_path) as con:
            df = pd.read_sql_query(
                f"""SELECT symbol,model,as_of,prediction,confidence,created_at
                    FROM predictions
                    WHERE symbol IN ({placeholders})
                    AND model='Expected 12M Excess Return'
                    ORDER BY datetime(created_at) DESC, id DESC""",
                con,
                params=tuple(aliases),
            )
        if df.empty:
            return df
        df["Ticker"] = df["symbol"].map(aliases)
        return df.dropna(subset=["Ticker"]).drop_duplicates("Ticker", keep="first")
    except Exception:
        return pd.DataFrame()


def _train_missing_ml_predictions(db_path: Path, missing: list[str]) -> pd.DataFrame:
    """Best-effort shared ExpectedReturnModel fit for holdings without a stored prediction."""
    if not missing or not db_path.exists():
        return pd.DataFrame()
    try:
        if str(ROOT) not in sys.path:
            sys.path.insert(0, str(ROOT))
        from machine_learning.models import ExpectedReturnModel
        with sqlite3.connect(db_path) as con:
            train = pd.read_sql_query(
                "SELECT * FROM features WHERE target_excess_return_12m IS NOT NULL ORDER BY as_of,symbol",
                con,
            )
            latest = pd.read_sql_query(
                """SELECT f.* FROM features f
                   JOIN (SELECT symbol,MAX(as_of) AS as_of FROM features GROUP BY symbol) x
                   ON f.symbol=x.symbol AND f.as_of=x.as_of""",
                con,
            )
        if len(train) < 30 or latest.empty:
            return pd.DataFrame()
        train["as_of"] = pd.to_datetime(train["as_of"], errors="coerce")
        train["target_date"] = pd.to_datetime(train["target_date"], errors="coerce")
        model = ExpectedReturnModel()
        rows = []
        for ticker in missing:
            cand = latest[latest["symbol"].isin(_provider_candidates(ticker))]
            if cand.empty:
                continue
            current = cand.iloc[-1].to_dict()
            result = model.fit_predict(train, current)
            if result.status != "PASS" or result.prediction is None:
                continue
            rows.append({
                "Ticker": ticker,
                "prediction": float(result.prediction),
                "confidence": result.confidence,
                "as_of": current.get("as_of"),
                "MLSource": "On-demand shared point-in-time model",
            })
        return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame()


def _confidence_scalar(label: str | None, as_of=None) -> float:
    base = CONFIDENCE_WEIGHT.get(str(label), 0.0)
    if base <= 0:
        return 0.0
    try:
        dt = pd.Timestamp(as_of)
        if dt.tzinfo is None:
            dt = dt.tz_localize("UTC")
        now = pd.Timestamp(datetime.now(timezone.utc))
        age_days = max(0.0, (now - dt).total_seconds() / 86400)
        freshness = math.exp(-age_days / 730.0)
    except Exception:
        freshness = 0.75
    return float(base * freshness)


def expected_return_inputs(
    returns: pd.DataFrame,
    benchmark_returns: pd.Series,
    manual: pd.DataFrame,
    tickers: list[str],
    history_db: str | Path | None,
) -> tuple[pd.DataFrame, dict]:
    """Classical prior + confidence-shrunk ML expected return inputs."""
    raw = {t: _window_return_prior(returns[t]) if t in returns else None for t in tickers}
    valid = [v for v in raw.values() if v is not None and np.isfinite(v)]
    cross_median = float(np.median(valid)) if valid else 0.08
    historical = {t: 0.55 * raw[t] + 0.45 * cross_median if raw[t] is not None else cross_median for t in tickers}
    benchmark_mu = _window_return_prior(benchmark_returns)
    if benchmark_mu is None:
        benchmark_mu = cross_median

    manual_map = {}
    if manual is not None and not manual.empty and "Ticker" in manual and "ExpectedReturn" in manual:
        manual_map = (
            manual.assign(Ticker=manual["Ticker"].astype(str).str.upper())
            .set_index("Ticker")["ExpectedReturn"].to_dict()
        )

    db_path = Path(history_db) if history_db else Path()
    saved = _latest_prediction_rows(db_path, tickers) if history_db else pd.DataFrame()
    saved_map = {}
    if not saved.empty:
        for _, row in saved.iterrows():
            pred = _parse_json(row.get("prediction"))
            if isinstance(pred, dict):
                pred = pred.get("prediction") or pred.get("value")
            pred = _finite(pred)
            if pred is not None:
                saved_map[row["Ticker"]] = {
                    "prediction": pred,
                    "confidence": row.get("confidence"),
                    "as_of": row.get("as_of") or row.get("created_at"),
                    "MLSource": "Stored walk-forward ML prediction",
                }

    missing = [t for t in tickers if t not in saved_map]
    trained = _train_missing_ml_predictions(db_path, missing) if history_db else pd.DataFrame()
    trained_map = {}
    if not trained.empty:
        trained_map = {row["Ticker"]: row.to_dict() for _, row in trained.iterrows()}

    rows = []
    for t in tickers:
        hist = float(historical[t])
        man = _finite(manual_map.get(t))
        classical = 0.60 * hist + 0.40 * man if man is not None else hist
        ml = saved_map.get(t) or trained_map.get(t)
        ml_excess = _finite((ml or {}).get("prediction"))
        confidence = (ml or {}).get("confidence")
        scalar = _confidence_scalar(confidence, (ml or {}).get("as_of"))
        ml_weight = min(0.55, 0.55 * scalar)
        ml_total = benchmark_mu + ml_excess if ml_excess is not None else None
        blended = (1 - ml_weight) * classical + ml_weight * ml_total if ml_total is not None else classical
        rows.append({
            "Ticker": t,
            "HistoricalPrior": hist,
            "ManualExpectedReturn": man,
            "ClassicalExpectedReturn": classical,
            "MLExcessReturn12M": ml_excess,
            "MLTotalReturn": ml_total,
            "MLConfidence": confidence,
            "MLConfidenceScalar": scalar,
            "MLBlendWeight": ml_weight,
            "BlendedMLExpectedReturn": blended,
            "MLSource": (ml or {}).get("MLSource") if ml else "No ML prediction available",
        })
    meta = {
        "benchmark_expected_return": float(benchmark_mu),
        "ml_covered_assets": int(sum(x["MLExcessReturn12M"] is not None for x in rows)),
        "assets": len(tickers),
        "policy": "ML is a confidence-shrunk expected-return input; optimizer remains deterministic.",
    }
    return pd.DataFrame(rows), meta


def _latest_regime(db_path: Path) -> dict:
    if not db_path.exists():
        return {"regime": None, "confidence": None, "source": "No ML history database"}
    try:
        with sqlite3.connect(db_path) as con:
            row = con.execute(
                """SELECT prediction,confidence,as_of,created_at
                   FROM predictions WHERE model='Market Regime Classifier'
                   ORDER BY datetime(created_at) DESC,id DESC LIMIT 1"""
            ).fetchone()
        if not row:
            return {"regime": None, "confidence": None, "source": "No stored regime prediction"}
        regime = _parse_json(row[0])
        if isinstance(regime, dict):
            regime = regime.get("prediction") or regime.get("regime")
        return {
            "regime": str(regime) if regime else None,
            "confidence": row[1],
            "as_of": row[2] or row[3],
            "source": "Stored Market Regime Classifier",
        }
    except Exception:
        return {"regime": None, "confidence": None, "source": "Regime query failed"}


def regime_aware_covariance(
    returns: pd.DataFrame,
    benchmark_returns: pd.Series,
    long_cov: pd.DataFrame,
    regime: dict,
    max_blend: float = 0.45,
) -> tuple[pd.DataFrame, dict]:
    """Blend long-run robust covariance with covariance from historically similar market states."""
    label = str(regime.get("regime") or "")
    if returns.empty or long_cov.empty or not label:
        return long_cov.copy(), {**regime, "blend_weight": 0.0, "regime_rows": 0}

    b = pd.to_numeric(benchmark_returns, errors="coerce").reindex(returns.index)
    mom = (1 + b.fillna(0)).rolling(126).apply(np.prod, raw=True) - 1
    vol = b.rolling(63).std() * np.sqrt(TRADING_DAYS)
    vol_med = vol.median()
    mom_med = mom.median()
    low = label.lower()
    if "risk-off" in low or "crisis" in low:
        mask = (mom < 0) & (vol >= vol_med)
    elif "recession" in low or "disinflation" in low:
        mask = (mom < 0) & (vol < vol_med)
    elif "inflation" in low or "stagflation" in low:
        mask = (mom <= mom_med) & (vol >= vol_med)
    elif "growth" in low or "risk-on" in low:
        mask = (mom > mom_med) & (vol <= vol_med)
    else:
        mask = (mom.between(mom.quantile(0.25), mom.quantile(0.75))) & (
            vol.between(vol.quantile(0.25), vol.quantile(0.75))
        )

    state = returns.loc[mask.fillna(False)]
    if len(state.dropna(how="all")) < 126:
        state = returns.tail(min(len(returns), 756))
    state_cov, _ = robust_covariance(state)
    if state_cov.empty:
        return long_cov.copy(), {**regime, "blend_weight": 0.0, "regime_rows": int(len(state))}

    c = _confidence_scalar(regime.get("confidence"), regime.get("as_of"))
    blend = float(min(max_blend, max_blend * c))
    aligned = state_cov.reindex(index=long_cov.index, columns=long_cov.columns)
    cov = (1 - blend) * long_cov.values + blend * aligned.values
    cov = _nearest_psd(cov)
    return pd.DataFrame(cov, index=long_cov.index, columns=long_cov.columns), {
        **regime,
        "blend_weight": blend,
        "regime_rows": int(len(state.dropna(how="all"))),
    }


@dataclass
class ConstraintSet:
    bounds: list[tuple[float, float]]
    constraints: list[dict]
    effective_max_position: float
    effective_max_sector: float
    turnover_limit: float | None


def _constraints(
    tickers: list[str],
    holdings: pd.DataFrame,
    expected_returns: pd.DataFrame,
    max_position: float,
    max_sector: float,
    turnover_limit: float | None,
) -> tuple[ConstraintSet, np.ndarray]:
    n = len(tickers)
    current = holdings.set_index("Ticker")["Weight"].reindex(tickers).fillna(0).to_numpy(float)
    current = current / current.sum() if current.sum() else np.repeat(1 / n, n)

    expected = expected_returns.set_index("Ticker") if expected_returns is not None and not expected_returns.empty else pd.DataFrame()
    effective_max = max(float(max_position), 1.0 / n + 1e-6)
    bounds = []
    for t in tickers:
        lo, hi = 0.0, effective_max
        if not expected.empty and t in expected.index:
            lo_v = _finite(expected.loc[t].get("MinWeight"))
            hi_v = _finite(expected.loc[t].get("MaxWeight"))
            if lo_v is not None:
                lo = max(0.0, lo_v)
            if hi_v is not None:
                hi = min(1.0, hi_v)
        bounds.append((lo, max(lo, hi)))

    cons = [{"type": "eq", "fun": lambda w: float(np.sum(w) - 1.0)}]
    sectors = holdings.set_index("Ticker").get("Sector", pd.Series(dtype=object)).reindex(tickers).fillna("Unknown")
    unique_sectors = list(dict.fromkeys(sectors.tolist()))
    effective_sector = 1.0 if len(unique_sectors) <= 1 else float(max_sector)
    if effective_sector < 1:
        for sector in unique_sectors:
            idx = np.array([i for i, t in enumerate(tickers) if sectors.loc[t] == sector], dtype=int)
            cons.append({
                "type": "ineq",
                "fun": lambda w, idx=idx, cap=effective_sector: float(cap - np.sum(w[idx])),
            })
    if turnover_limit is not None and turnover_limit > 0:
        cons.append({
            "type": "ineq",
            "fun": lambda w, cur=current, cap=float(turnover_limit):
                float(cap - 0.5 * np.abs(w - cur).sum()),
        })
    return ConstraintSet(bounds, cons, effective_max, effective_sector, turnover_limit), current


def _initial_weights(current: np.ndarray, bounds: list[tuple[float, float]]) -> np.ndarray:
    lo = np.array([b[0] for b in bounds])
    hi = np.array([b[1] for b in bounds])
    x = np.clip(current, lo, hi)
    if x.sum() <= 0:
        x = np.repeat(1 / len(x), len(x))
    x = x / x.sum()
    for _ in range(20):
        over = x > hi + 1e-10
        if not over.any():
            break
        excess = float((x[over] - hi[over]).sum())
        x[over] = hi[over]
        room = (~over) & (x < hi - 1e-10)
        if not room.any():
            break
        capacity = hi[room] - x[room]
        x[room] += excess * capacity / capacity.sum()
    return x / x.sum()


def _solve(name, objective, x0, cs: ConstraintSet, extra_constraints=None):
    cons = list(cs.constraints) + list(extra_constraints or [])
    res = minimize(
        objective, x0=x0, method="SLSQP", bounds=cs.bounds, constraints=cons,
        options={"maxiter": 1500, "ftol": 1e-12},
    )
    return name, res


def _metrics(weights, mu, cov, risk_free_rate):
    ret = float(weights @ mu)
    vol = float(np.sqrt(max(weights @ cov @ weights, 0)))
    sharpe = (ret - risk_free_rate) / vol if vol > 0 else np.nan
    return ret, vol, sharpe


def optimize_suite(
    holdings: pd.DataFrame,
    expected_bounds: pd.DataFrame,
    expected_inputs: pd.DataFrame,
    classical_cov: pd.DataFrame,
    regime_cov: pd.DataFrame,
    risk_free_rate: float,
    max_position: float,
    max_sector: float,
    turnover_limit: float | None,
    risk_aversion: float,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    tickers = list(classical_cov.columns)
    if len(tickers) < 2:
        return pd.DataFrame(), pd.DataFrame(), {"status": "insufficient"}
    inputs = expected_inputs.set_index("Ticker").reindex(tickers)
    classical_mu = inputs["ClassicalExpectedReturn"].to_numpy(float)
    ml_mu = inputs["BlendedMLExpectedReturn"].to_numpy(float)
    cov_c = classical_cov.loc[tickers, tickers].to_numpy(float)
    cov_m = regime_cov.loc[tickers, tickers].to_numpy(float)
    cs, current = _constraints(
        tickers, holdings, expected_bounds, max_position, max_sector, turnover_limit
    )
    x0 = _initial_weights(current, cs.bounds)

    solutions: list[tuple[str, np.ndarray, np.ndarray, np.ndarray]] = []
    solutions.append(("Current Portfolio", current, classical_mu, cov_c))
    solutions.append(("Equal Weight", np.repeat(1 / len(tickers), len(tickers)), classical_mu, cov_c))

    name, res = _solve("Minimum Variance", lambda w: float(w @ cov_c @ w), x0, cs)
    if res.success:
        solutions.append((name, res.x, classical_mu, cov_c))

    def neg_sharpe(mu, cov):
        def obj(w):
            ret, vol, _ = _metrics(w, mu, cov, risk_free_rate)
            return -((ret - risk_free_rate) / vol) if vol > 0 else 1e6
        return obj

    name, res = _solve("Maximum Sharpe — Classical", neg_sharpe(classical_mu, cov_c), x0, cs)
    if res.success:
        solutions.append((name, res.x, classical_mu, cov_c))

    name, res = _solve(
        "Mean-Variance — Classical",
        lambda w: float(-(w @ classical_mu - 0.5 * risk_aversion * (w @ cov_c @ w))),
        x0, cs,
    )
    if res.success:
        solutions.append((name, res.x, classical_mu, cov_c))

    name, res = _solve("ML Maximum Sharpe", neg_sharpe(ml_mu, cov_c), x0, cs)
    if res.success:
        solutions.append((name, res.x, ml_mu, cov_c))

    name, res = _solve(
        "ML Mean-Variance",
        lambda w: float(-(w @ ml_mu - 0.5 * risk_aversion * (w @ cov_c @ w))),
        x0, cs,
    )
    if res.success:
        solutions.append((name, res.x, ml_mu, cov_c))

    name, res = _solve("Regime-Aware ML Maximum Sharpe", neg_sharpe(ml_mu, cov_m), x0, cs)
    if res.success:
        solutions.append((name, res.x, ml_mu, cov_m))

    summary_rows, weight_rows = [], []
    sectors = holdings.set_index("Ticker").get("Sector", pd.Series(dtype=object)).reindex(tickers).fillna("Unknown")
    for name, w, mu, cov in solutions:
        pret, pvol, sharpe = _metrics(w, mu, cov, risk_free_rate)
        turnover = 0.5 * float(np.abs(w - current).sum())
        summary_rows.append({
            "Portfolio": name,
            "ExpectedReturn": pret,
            "ExpectedVolatility": pvol,
            "ExpectedSharpe": sharpe,
            "Turnover": turnover,
            "MaxWeight": float(w.max()),
            "EffectiveHoldings": float(1 / np.sum(w ** 2)) if np.sum(w ** 2) > 0 else np.nan,
        })
        for i, t in enumerate(tickers):
            weight_rows.append({
                "Portfolio": name,
                "Ticker": t,
                "Sector": sectors.loc[t],
                "CurrentWeight": float(current[i]),
                "TargetWeight": float(w[i]),
                "WeightChange": float(w[i] - current[i]),
            })
    metadata = {
        "status": "PASS",
        "effective_max_position": cs.effective_max_position,
        "effective_max_sector": cs.effective_max_sector,
        "turnover_limit": cs.turnover_limit,
        "risk_aversion": risk_aversion,
    }
    return pd.DataFrame(summary_rows), pd.DataFrame(weight_rows), metadata


def efficient_frontier(
    mu: pd.Series,
    cov: pd.DataFrame,
    holdings: pd.DataFrame,
    expected_bounds: pd.DataFrame,
    max_position: float,
    max_sector: float,
    turnover_limit: float | None,
    risk_free_rate: float,
    points: int = 50,
    label: str = "Classical Frontier",
) -> pd.DataFrame:
    tickers = list(cov.columns)
    m = mu.reindex(tickers).to_numpy(float)
    c = cov.loc[tickers, tickers].to_numpy(float)
    cs, current = _constraints(
        tickers, holdings, expected_bounds, max_position, max_sector, turnover_limit
    )
    x0 = _initial_weights(current, cs.bounds)

    _, rmin = _solve("min-return", lambda w: float(w @ m), x0, cs)
    _, rmax = _solve("max-return", lambda w: float(-(w @ m)), x0, cs)
    if not (rmin.success and rmax.success):
        return pd.DataFrame()
    lo = float(rmin.x @ m)
    hi = float(rmax.x @ m)
    if hi <= lo:
        return pd.DataFrame()
    rows = []
    last = x0
    for target in np.linspace(lo, hi, max(10, int(points))):
        extra = [{"type": "eq", "fun": lambda w, target=target: float(w @ m - target)}]
        _, res = _solve("frontier", lambda w: float(w @ c @ w), last, cs, extra)
        if not res.success:
            continue
        last = res.x
        ret, vol, sharpe = _metrics(res.x, m, c, risk_free_rate)
        rows.append({
            "Frontier": label,
            "TargetReturn": target,
            "ExpectedReturn": ret,
            "ExpectedVolatility": vol,
            "ExpectedSharpe": sharpe,
        })
    return pd.DataFrame(rows)


def build_portfolio_optimization(
    asset_returns: pd.DataFrame,
    benchmark_returns: pd.Series,
    holdings: pd.DataFrame,
    manual_expected_returns: pd.DataFrame,
    risk_free_rate: float,
    config: dict,
    history_db: str | Path | None,
) -> dict:
    cfg = config or {}
    tickers = holdings["Ticker"].astype(str).str.upper().tolist()
    max_years = int(cfg.get("max_history_years", 20))
    max_position = float(cfg.get("max_position", 0.25))
    max_sector = float(cfg.get("max_sector", 0.45))
    turnover_limit = _finite(cfg.get("turnover_limit", 0.50))
    risk_aversion = float(cfg.get("risk_aversion", 4.0))
    frontier_points = int(cfg.get("frontier_points", 50))

    returns, coverage = maximum_history_returns(
        tickers, asset_returns, history_db, max_years=max_years
    )
    usable = [t for t in tickers if t in returns.columns and returns[t].notna().sum() >= 126]
    if len(usable) < 2:
        return {
            "summary": pd.DataFrame(),
            "weights": pd.DataFrame(),
            "frontier": pd.DataFrame(),
            "expected_returns": pd.DataFrame(),
            "covariance": pd.DataFrame(),
            "regime_covariance": pd.DataFrame(),
            "correlation": pd.DataFrame(),
            "coverage": coverage,
            "metadata": {"status": "INSUFFICIENT_DATA"},
        }
    returns = returns[usable]
    holdings = holdings[holdings["Ticker"].isin(usable)].copy()
    holdings["Weight"] = holdings["Weight"] / holdings["Weight"].sum()
    manual = manual_expected_returns[
        manual_expected_returns["Ticker"].isin(usable)
    ].copy() if manual_expected_returns is not None and not manual_expected_returns.empty else pd.DataFrame()

    long_cov, cov_meta = robust_covariance(returns)
    exp_inputs, exp_meta = expected_return_inputs(
        returns, benchmark_returns, manual, usable, history_db
    )
    regime = _latest_regime(Path(history_db)) if history_db else {
        "regime": None, "confidence": None, "source": "No ML DB"
    }
    regime_cov, regime_meta = regime_aware_covariance(
        returns, benchmark_returns, long_cov, regime,
        max_blend=float(cfg.get("regime_covariance_blend", 0.45)),
    )

    summary, weights, opt_meta = optimize_suite(
        holdings, manual, exp_inputs, long_cov, regime_cov, risk_free_rate,
        max_position=max_position, max_sector=max_sector,
        turnover_limit=turnover_limit, risk_aversion=risk_aversion,
    )
    classical_mu = exp_inputs.set_index("Ticker")["ClassicalExpectedReturn"]
    ml_mu = exp_inputs.set_index("Ticker")["BlendedMLExpectedReturn"]
    frontier_c = efficient_frontier(
        classical_mu, long_cov, holdings, manual, max_position, max_sector,
        turnover_limit, risk_free_rate, points=frontier_points, label="Classical Frontier"
    )
    frontier_ml = efficient_frontier(
        ml_mu, regime_cov, holdings, manual, max_position, max_sector,
        turnover_limit, risk_free_rate, points=frontier_points, label="ML / Regime Frontier"
    )
    frontier = pd.concat([frontier_c, frontier_ml], ignore_index=True)

    corr = returns.corr(min_periods=126)
    metadata = {
        **opt_meta,
        "max_history_years": max_years,
        "history_assets": len(usable),
        "covariance": cov_meta,
        "expected_returns": exp_meta,
        "regime": regime_meta,
    }
    return {
        "summary": summary,
        "weights": weights,
        "frontier": frontier,
        "expected_returns": exp_inputs,
        "covariance": long_cov.reset_index().rename(columns={"index": "Ticker"}),
        "regime_covariance": regime_cov.reset_index().rename(columns={"index": "Ticker"}),
        "correlation": corr.reset_index().rename(columns={"index": "Ticker"}),
        "coverage": coverage,
        "metadata": metadata,
    }
