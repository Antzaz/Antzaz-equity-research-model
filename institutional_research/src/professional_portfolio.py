from __future__ import annotations

from pathlib import Path
import math
import numpy as np
import pandas as pd

try:
    from scipy.optimize import minimize
except Exception:
    minimize = None


def benchmark_relative_metrics(
    portfolio_returns: pd.Series,
    benchmark_returns: pd.Series,
    risk_free_rate: float = 0.04,
    trading_days: int = 252,
) -> dict:
    common = portfolio_returns.dropna().index.intersection(benchmark_returns.dropna().index)
    p = portfolio_returns.loc[common].astype(float)
    b = benchmark_returns.loc[common].astype(float)
    if len(common) < 2:
        return {}
    active = p - b
    p_ann = (1 + p).prod() ** (trading_days / len(p)) - 1
    b_ann = (1 + b).prod() ** (trading_days / len(b)) - 1
    active_ann = p_ann - b_ann
    te = active.std(ddof=1) * math.sqrt(trading_days)
    ir = active_ann / te if te > 0 else np.nan
    corr = p.corr(b)
    up = b > 0
    down = b < 0
    up_capture = p[up].mean() / b[up].mean() if up.any() and b[up].mean() != 0 else np.nan
    down_capture = p[down].mean() / b[down].mean() if down.any() and b[down].mean() != 0 else np.nan
    hit_rate = float((p > b).mean())
    return {
        "portfolio_annualized_return": float(p_ann),
        "benchmark_annualized_return": float(b_ann),
        "active_annualized_return": float(active_ann),
        "tracking_error": float(te),
        "information_ratio": float(ir) if np.isfinite(ir) else None,
        "benchmark_correlation": float(corr) if np.isfinite(corr) else None,
        "up_capture": float(up_capture) if np.isfinite(up_capture) else None,
        "down_capture": float(down_capture) if np.isfinite(down_capture) else None,
        "daily_active_hit_rate": hit_rate,
    }


def concentration_metrics(holdings: pd.DataFrame) -> dict:
    w = pd.to_numeric(holdings["Weight"], errors="coerce").fillna(0).clip(lower=0)
    w = w / w.sum() if w.sum() else w
    hhi = float((w ** 2).sum())
    effective_n = float(1 / hhi) if hhi > 0 else np.nan
    ordered = w.sort_values(ascending=False)
    sector = (
        holdings.assign(_w=w.values, _sector=holdings["Sector"].fillna("Unknown"))
        .groupby("_sector")["_w"].sum()
    )
    shhi = float((sector ** 2).sum()) if not sector.empty else np.nan
    effective_sectors = float(1 / shhi) if shhi and shhi > 0 else np.nan
    rc = pd.to_numeric(holdings.get("RiskContributionPct"), errors="coerce")
    return {
        "top_1_weight": float(ordered.iloc[:1].sum()) if len(ordered) else None,
        "top_3_weight": float(ordered.iloc[:3].sum()) if len(ordered) else None,
        "top_5_weight": float(ordered.iloc[:5].sum()) if len(ordered) else None,
        "herfindahl_index": hhi,
        "effective_number_of_holdings": effective_n if np.isfinite(effective_n) else None,
        "sector_herfindahl_index": shhi if np.isfinite(shhi) else None,
        "effective_number_of_sectors": effective_sectors if np.isfinite(effective_sectors) else None,
        "largest_risk_contribution": float(rc.max()) if rc.notna().any() else None,
    }


def risk_budget_table(holdings: pd.DataFrame) -> pd.DataFrame:
    cols = ["Ticker", "Weight", "RiskContributionPct"]
    out = holdings[[c for c in cols if c in holdings.columns]].copy()
    if "RiskContributionPct" not in out:
        out["RiskContributionPct"] = np.nan
    out["RiskVsCapital"] = out["RiskContributionPct"] - out["Weight"]
    out["RiskToCapitalRatio"] = np.where(
        out["Weight"] > 0, out["RiskContributionPct"] / out["Weight"], np.nan
    )
    return out.sort_values("RiskContributionPct", ascending=False)


def liquidity_analysis(
    holdings: pd.DataFrame,
    info: dict[str, dict],
    participation_rate: float = 0.10,
) -> pd.DataFrame:
    rows = []
    participation_rate = max(0.001, min(1.0, float(participation_rate)))
    for _, row in holdings.iterrows():
        t = row["Ticker"]
        d = info.get(t) or {}
        px = row.get("CurrentPrice")
        avg_volume = d.get("averageVolume") or d.get("averageVolume10days") or d.get("volume")
        adv = float(px) * float(avg_volume) if px and avg_volume else np.nan
        mv = row.get("MarketValue")
        days = float(mv) / (adv * participation_rate) if pd.notna(mv) and adv and adv > 0 else np.nan
        rows.append({
            "Ticker": t,
            "MarketValue": mv,
            "AverageDailyVolumeShares": avg_volume,
            "AverageDailyDollarVolume": adv,
            "ParticipationRate": participation_rate,
            "EstimatedDaysToLiquidate": days,
            "PositionPctADV": float(mv) / adv if pd.notna(mv) and adv and adv > 0 else np.nan,
        })
    return pd.DataFrame(rows)


def factor_proxy_sensitivity(
    portfolio_returns: pd.Series,
    proxy_returns: pd.DataFrame,
    trading_days: int = 252,
) -> pd.DataFrame:
    rows = []
    p = portfolio_returns.dropna()
    for col in proxy_returns.columns:
        f = proxy_returns[col].dropna()
        idx = p.index.intersection(f.index)
        if len(idx) < 60:
            continue
        x = f.loc[idx].astype(float)
        y = p.loc[idx].astype(float)
        var = np.var(x, ddof=1)
        beta = np.cov(y, x, ddof=1)[0, 1] / var if var > 0 else np.nan
        corr = y.corr(x)
        residual = y - beta * x if np.isfinite(beta) else y * np.nan
        residual_vol = residual.std(ddof=1) * math.sqrt(trading_days)
        rows.append({
            "Proxy": col,
            "BetaToProxy": beta,
            "Correlation": corr,
            "ResidualVolatility": residual_vol,
            "Observations": len(idx),
        })
    return pd.DataFrame(rows)


def rolling_risk_table(
    portfolio_returns: pd.Series,
    benchmark_returns: pd.Series,
    trading_days: int = 252,
) -> pd.DataFrame:
    p = portfolio_returns.dropna()
    b = benchmark_returns.reindex(p.index)
    rows = []
    for window, label in [(21, "1M"), (63, "3M"), (126, "6M"), (252, "1Y")]:
        if len(p) < window:
            continue
        rp = p.rolling(window)
        ann_vol = rp.std(ddof=1) * math.sqrt(trading_days)
        active = (p - b).rolling(window).std(ddof=1) * math.sqrt(trading_days)
        roll_ret = (1 + p).rolling(window).apply(np.prod, raw=True) - 1
        rows.append({
            "Window": label,
            "LatestReturn": float(roll_ret.dropna().iloc[-1]) if not roll_ret.dropna().empty else np.nan,
            "LatestVolatility": float(ann_vol.dropna().iloc[-1]) if not ann_vol.dropna().empty else np.nan,
            "LatestTrackingError": float(active.dropna().iloc[-1]) if not active.dropna().empty else np.nan,
            "WorstRollingReturn": float(roll_ret.min()) if roll_ret.notna().any() else np.nan,
            "BestRollingReturn": float(roll_ret.max()) if roll_ret.notna().any() else np.nan,
        })
    return pd.DataFrame(rows)


def historical_stress_windows(
    portfolio_returns: pd.Series,
    benchmark_returns: pd.Series,
    scenarios: list[dict],
) -> pd.DataFrame:
    rows = []
    for sc in scenarios or []:
        try:
            start = pd.Timestamp(sc["start"])
            end = pd.Timestamp(sc["end"])
        except Exception:
            continue
        p = portfolio_returns.loc[(portfolio_returns.index >= start) & (portfolio_returns.index <= end)].dropna()
        b = benchmark_returns.loc[(benchmark_returns.index >= start) & (benchmark_returns.index <= end)].dropna()
        if p.empty:
            rows.append({"Scenario": sc.get("name", f"{start.date()} to {end.date()}"),
                         "Start": start.date(), "End": end.date(),
                         "PortfolioReturn": np.nan, "BenchmarkReturn": np.nan,
                         "ActiveReturn": np.nan, "Observations": 0})
            continue
        pr = float((1 + p).prod() - 1)
        br = float((1 + b).prod() - 1) if not b.empty else np.nan
        rows.append({
            "Scenario": sc.get("name", f"{start.date()} to {end.date()}"),
            "Start": start.date(),
            "End": end.date(),
            "PortfolioReturn": pr,
            "BenchmarkReturn": br,
            "ActiveReturn": pr - br if np.isfinite(br) else np.nan,
            "Observations": len(p),
        })
    return pd.DataFrame(rows)


def static_return_attribution(
    asset_returns: pd.DataFrame,
    weights: pd.Series,
) -> pd.DataFrame:
    r = asset_returns.dropna(how="all").copy()
    w = weights.reindex(r.columns).fillna(0).astype(float)
    w = w / w.sum() if w.sum() else w
    rows = []
    for t in r.columns:
        x = r[t].dropna()
        total = float((1 + x).prod() - 1) if not x.empty else np.nan
        arithmetic_contribution = float((r[t].fillna(0) * w[t]).sum())
        rows.append({
            "Ticker": t,
            "Weight": float(w[t]),
            "AssetTotalReturn": total,
            "StaticArithmeticContribution": arithmetic_contribution,
        })
    out = pd.DataFrame(rows)
    total_abs = out["StaticArithmeticContribution"].abs().sum()
    out["ContributionShareAbs"] = (
        out["StaticArithmeticContribution"].abs() / total_abs if total_abs else np.nan
    )
    return out.sort_values("StaticArithmeticContribution", ascending=False)


def load_expected_returns(path: str | Path, tickers: list[str]) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return pd.DataFrame(columns=["Ticker", "ExpectedReturn", "Conviction", "MinWeight", "MaxWeight"])
    df = pd.read_csv(path, comment="#")
    for c in ["Ticker", "ExpectedReturn", "Conviction", "MinWeight", "MaxWeight"]:
        if c not in df.columns:
            df[c] = np.nan
    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    df = df[df["Ticker"].isin(tickers)].copy()
    for c in ["ExpectedReturn", "Conviction", "MinWeight", "MaxWeight"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def _bounds_for_tickers(tickers, expected, max_position):
    by = expected.set_index("Ticker") if not expected.empty else pd.DataFrame()
    bounds = []
    for t in tickers:
        lo = 0.0
        hi = float(max_position)
        if not expected.empty and t in by.index:
            lo_v = by.loc[t, "MinWeight"]
            hi_v = by.loc[t, "MaxWeight"]
            if pd.notna(lo_v):
                lo = max(0.0, float(lo_v))
            if pd.notna(hi_v):
                hi = min(1.0, float(hi_v))
        hi = max(lo, hi)
        bounds.append((lo, hi))
    return bounds


def optimize_portfolios(
    asset_returns: pd.DataFrame,
    holdings: pd.DataFrame,
    expected_returns: pd.DataFrame,
    risk_free_rate: float,
    max_position: float = 0.25,
) -> pd.DataFrame:
    tickers = list(asset_returns.columns)
    r = asset_returns.dropna(how="any")
    if len(tickers) < 2 or len(r) < 60 or minimize is None:
        return pd.DataFrame()
    cov = r.cov().to_numpy() * 252
    current = holdings.set_index("Ticker")["Weight"].reindex(tickers).fillna(0).to_numpy(dtype=float)
    current = current / current.sum()
    bounds = _bounds_for_tickers(tickers, expected_returns, max_position)
    if sum(b[1] for b in bounds) < 1 - 1e-9:
        return pd.DataFrame()

    cons = [{"type": "eq", "fun": lambda w: np.sum(w) - 1}]
    x0 = np.clip(current, [b[0] for b in bounds], [b[1] for b in bounds])
    if x0.sum() <= 0:
        x0 = np.repeat(1 / len(tickers), len(tickers))
    x0 = x0 / x0.sum()

    def vol(w):
        return float(np.sqrt(max(w @ cov @ w, 0)))

    def risk_parity_obj(w):
        pv = vol(w)
        if pv <= 0:
            return 1e6
        m = cov @ w / pv
        c = w * m
        target = np.repeat(pv / len(w), len(w))
        return float(np.sum((c - target) ** 2))

    solutions = []
    for name, obj in [("Minimum Variance", lambda w: w @ cov @ w),
                      ("Equal Risk Contribution", risk_parity_obj)]:
        res = minimize(obj, x0=x0, bounds=bounds, constraints=cons, method="SLSQP")
        if res.success:
            solutions.append((name, res.x))

    exp = expected_returns.set_index("Ticker")["ExpectedReturn"].reindex(tickers) if not expected_returns.empty else pd.Series(index=tickers, dtype=float)
    if exp.notna().all():
        mu = exp.to_numpy(dtype=float)
        def neg_sharpe(w):
            v = vol(w)
            return -((float(w @ mu) - risk_free_rate) / v) if v > 0 else 1e6
        res = minimize(neg_sharpe, x0=x0, bounds=bounds, constraints=cons, method="SLSQP")
        if res.success:
            solutions.append(("Max Expected Sharpe", res.x))

    rows = []
    total_mv = holdings["MarketValue"].sum() if holdings["MarketValue"].notna().all() else np.nan
    px = holdings.set_index("Ticker")["CurrentPrice"]
    for name, w in solutions:
        turnover = 0.5 * float(np.abs(w - current).sum())
        pv = vol(w)
        pret = float(w @ exp.to_numpy(dtype=float)) if exp.notna().all() else np.nan
        for t, cw, tw in zip(tickers, current, w):
            target_value = total_mv * tw if np.isfinite(total_mv) else np.nan
            delta_value = total_mv * (tw - cw) if np.isfinite(total_mv) else np.nan
            delta_shares = delta_value / px.get(t) if np.isfinite(delta_value) and px.get(t) else np.nan
            rows.append({
                "Portfolio": name,
                "Ticker": t,
                "CurrentWeight": cw,
                "TargetWeight": tw,
                "WeightChange": tw - cw,
                "ExpectedPortfolioReturn": pret,
                "ExpectedVolatility": pv,
                "OneWayTurnover": turnover,
                "TargetMarketValue": target_value,
                "TradeValue": delta_value,
                "EstimatedTradeShares": delta_shares,
            })
    return pd.DataFrame(rows)


def constraint_report(
    holdings: pd.DataFrame,
    risk_summary: dict,
    benchmark_metrics: dict,
    liquidity: pd.DataFrame,
    config: dict,
) -> pd.DataFrame:
    c = config or {}
    rows = []
    def add(name, value, limit, rule, unit=""):
        breach = False
        if value is not None and np.isfinite(value) and limit is not None:
            breach = value > limit if rule == "<=" else value < limit
        rows.append({"Constraint": name, "Current": value, "Limit": limit, "Rule": rule, "Unit": unit, "Breach": breach})

    add("Maximum single-name weight", float(holdings["Weight"].max()), c.get("max_position"), "<=", "%")
    sectors = holdings.assign(_sector=holdings["Sector"].fillna("Unknown")).groupby("_sector")["Weight"].sum()
    add("Maximum sector weight", float(sectors.max()) if not sectors.empty else np.nan, c.get("max_sector"), "<=", "%")
    rc = pd.to_numeric(holdings.get("RiskContributionPct"), errors="coerce")
    add("Maximum single-name risk contribution", float(rc.max()) if rc.notna().any() else np.nan, c.get("max_risk_contribution"), "<=", "%")
    add("Portfolio beta", risk_summary.get("beta"), c.get("max_beta"), "<=", "x")
    add("Tracking error", benchmark_metrics.get("tracking_error"), c.get("max_tracking_error"), "<=", "%")
    days = pd.to_numeric(liquidity.get("EstimatedDaysToLiquidate"), errors="coerce")
    add("Maximum days to liquidate", float(days.max()) if days.notna().any() else np.nan, c.get("max_days_to_liquidate"), "<=", "days")
    return pd.DataFrame(rows)


def active_share_from_file(portfolio_weights: pd.Series, benchmark_weights_path: str | Path) -> pd.DataFrame:
    path = Path(benchmark_weights_path)
    if not path.exists():
        return pd.DataFrame()
    b = pd.read_csv(path, comment="#")
    if not {"Ticker", "Weight"}.issubset(b.columns):
        return pd.DataFrame()
    b["Ticker"] = b["Ticker"].astype(str).str.upper().str.strip()
    b["Weight"] = pd.to_numeric(b["Weight"], errors="coerce").fillna(0)
    if b["Weight"].sum() <= 0:
        return pd.DataFrame()
    b["Weight"] /= b["Weight"].sum()
    p = portfolio_weights.copy()
    p.index = p.index.astype(str).str.upper()
    universe = sorted(set(p.index).union(b["Ticker"]))
    bw = b.set_index("Ticker")["Weight"].reindex(universe).fillna(0)
    pw = p.reindex(universe).fillna(0)
    detail = pd.DataFrame({"Ticker": universe, "PortfolioWeight": pw.values, "BenchmarkWeight": bw.values})
    detail["ActiveWeight"] = detail["PortfolioWeight"] - detail["BenchmarkWeight"]
    detail["AbsActiveWeight"] = detail["ActiveWeight"].abs()
    detail.attrs["active_share"] = 0.5 * float(detail["AbsActiveWeight"].sum())
    return detail
