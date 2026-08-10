from __future__ import annotations

"""Portfolio alpha diagnostics.

This module separates raw benchmark outperformance from risk/factor-adjusted residual return.
It supports:
- benchmark CAPM / Jensen alpha;
- Fama-French 3-factor alpha;
- Carhart 4-factor alpha (FF3 + momentum);
- Fama-French 5-factor alpha;
- a public ETF style-proxy model using the project's configured factor proxies;
- alpha t-statistics / p-values;
- rolling 1Y and 3Y alpha;
- arithmetic return decomposition into risk-free, factor-explained and residual alpha.

Important: the broader project applies current portfolio weights to historical asset returns.
Unless point-in-time weights / transactions are supplied, these are current-weight backcast
research diagnostics rather than realized historical manager alpha.
"""

from io import BytesIO
import math
import re
import zipfile

import numpy as np
import pandas as pd
import requests

try:
    from scipy.stats import t as student_t
except Exception:
    student_t = None


FF3_DAILY_URL = (
    "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
    "F-F_Research_Data_Factors_daily_CSV.zip"
)
MOM_DAILY_URL = (
    "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
    "F-F_Momentum_Factor_daily_CSV.zip"
)
FF5_DAILY_URL = (
    "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
    "F-F_Research_Data_5_Factors_2x3_daily_CSV.zip"
)


def _p_value(t_stat: float, dof: int) -> float | None:
    if not np.isfinite(t_stat) or dof <= 0:
        return None
    if student_t is not None:
        return float(2 * student_t.sf(abs(t_stat), dof))
    # Large-sample normal approximation fallback.
    return float(math.erfc(abs(t_stat) / math.sqrt(2)))


def _annualize_mean(x: pd.Series, trading_days: int) -> float:
    return float(pd.to_numeric(x, errors="coerce").mean() * trading_days)


def _download_french_zip(url: str, timeout: int = 30) -> pd.DataFrame:
    response = requests.get(url, timeout=timeout)
    response.raise_for_status()
    with zipfile.ZipFile(BytesIO(response.content)) as zf:
        names = [n for n in zf.namelist() if n.lower().endswith((".csv", ".txt"))]
        if not names:
            raise ValueError(f"No CSV/TXT file found in {url}")
        raw = zf.read(names[0])
    text = raw.decode("latin-1", errors="replace")
    lines = text.splitlines()

    header_idx = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith(",") and any(k in stripped.lower() for k in ("mkt-rf", "mom", "smb", "rmw")):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not find factor header in {url}")

    headers = ["Date"] + [x.strip() for x in lines[header_idx].split(",")[1:]]
    rows = []
    for line in lines[header_idx + 1 :]:
        parts = [x.strip() for x in line.split(",")]
        if not parts or not re.fullmatch(r"\d{8}", parts[0]):
            if rows:
                break
            continue
        vals = parts[: len(headers)]
        if len(vals) < len(headers):
            vals += [""] * (len(headers) - len(vals))
        rows.append(vals)

    if not rows:
        raise ValueError(f"No daily factor rows parsed from {url}")

    df = pd.DataFrame(rows, columns=headers)
    df["Date"] = pd.to_datetime(df["Date"], format="%Y%m%d", errors="coerce")
    df = df.dropna(subset=["Date"]).set_index("Date")
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df.loc[df[col].abs() >= 90, col] = np.nan
        df[col] = df[col] / 100.0
    return df.sort_index()


def load_french_daily_factors() -> tuple[pd.DataFrame, list[str]]:
    """Return daily FF factors in decimal returns plus a list of download warnings."""
    warnings: list[str] = []
    ff3 = mom = ff5 = None
    for label, url in (("FF3", FF3_DAILY_URL), ("Momentum", MOM_DAILY_URL), ("FF5", FF5_DAILY_URL)):
        try:
            df = _download_french_zip(url)
            if label == "FF3":
                ff3 = df
            elif label == "Momentum":
                mom = df
            else:
                ff5 = df
        except Exception as exc:
            warnings.append(f"{label} download failed: {exc}")

    frames = []
    if ff3 is not None:
        keep = [c for c in ["Mkt-RF", "SMB", "HML", "RF"] if c in ff3.columns]
        frames.append(ff3[keep])
    if mom is not None:
        mom_col = next((c for c in mom.columns if c.strip().lower() in {"mom", "umd"}), None)
        if mom_col:
            frames.append(mom[[mom_col]].rename(columns={mom_col: "Mom"}))
    if ff5 is not None:
        keep = [c for c in ["RMW", "CMA"] if c in ff5.columns]
        if keep:
            frames.append(ff5[keep])

    if not frames:
        return pd.DataFrame(), warnings
    out = pd.concat(frames, axis=1)
    out = out.loc[:, ~out.columns.duplicated()].sort_index()
    return out, warnings


def _ols(y: pd.Series, factors: pd.DataFrame, trading_days: int, model_name: str) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    data = pd.concat([y.rename("Y"), factors], axis=1).dropna()
    factor_cols = [c for c in data.columns if c != "Y"]
    n = len(data)
    k = len(factor_cols) + 1
    if n <= max(k + 5, 30):
        return {}, pd.DataFrame(), pd.DataFrame()

    yv = data["Y"].to_numpy(dtype=float)
    xv = data[factor_cols].to_numpy(dtype=float)
    X = np.column_stack([np.ones(n), xv])
    coef, _, _, _ = np.linalg.lstsq(X, yv, rcond=None)
    fitted = X @ coef
    resid = yv - fitted
    dof = n - k
    sse = float(resid @ resid)
    tss = float(((yv - yv.mean()) ** 2).sum())
    r2 = 1 - sse / tss if tss > 0 else np.nan
    sigma2 = sse / dof if dof > 0 else np.nan
    try:
        cov = sigma2 * np.linalg.inv(X.T @ X)
        se = np.sqrt(np.diag(cov))
    except Exception:
        se = np.full(k, np.nan)

    alpha_daily = float(coef[0])
    alpha_ann = alpha_daily * trading_days
    alpha_se_ann = float(se[0] * trading_days) if np.isfinite(se[0]) else np.nan
    alpha_t = float(coef[0] / se[0]) if np.isfinite(se[0]) and se[0] > 0 else np.nan
    p_value = _p_value(alpha_t, dof)

    summary = {
        "Model": model_name,
        "AnnualizedAlpha": alpha_ann,
        "AlphaStdErrorAnnualized": alpha_se_ann if np.isfinite(alpha_se_ann) else None,
        "AlphaTStat": alpha_t if np.isfinite(alpha_t) else None,
        "AlphaPValue": p_value,
        "R2": float(r2) if np.isfinite(r2) else None,
        "ResidualVolatilityAnnualized": float(np.std(resid, ddof=k) * math.sqrt(trading_days)) if dof > 0 else None,
        "Observations": int(n),
        "StartDate": data.index.min().date().isoformat(),
        "EndDate": data.index.max().date().isoformat(),
        "Significant5Pct": bool(p_value is not None and p_value < 0.05),
        "Interpretation": (
            "Statistically significant positive alpha" if p_value is not None and p_value < 0.05 and alpha_ann > 0
            else "Statistically significant negative alpha" if p_value is not None and p_value < 0.05 and alpha_ann < 0
            else "Positive alpha, not statistically significant" if alpha_ann > 0
            else "Negative alpha, not statistically significant" if alpha_ann < 0
            else "No measurable alpha"
        ),
    }

    loadings = []
    decomposition = []
    for j, col in enumerate(factor_cols, start=1):
        beta = float(coef[j])
        factor_ann = _annualize_mean(data[col], trading_days)
        contribution = beta * factor_ann
        loadings.append({
            "Model": model_name,
            "Factor": col,
            "Beta": beta,
            "StdError": float(se[j]) if np.isfinite(se[j]) else None,
            "TStat": float(coef[j] / se[j]) if np.isfinite(se[j]) and se[j] > 0 else None,
            "AnnualizedFactorMean": factor_ann,
            "AnnualizedContribution": contribution,
        })
        decomposition.append({
            "Model": model_name,
            "Component": col,
            "AnnualizedContribution": contribution,
            "Type": "Factor explained",
        })
    decomposition.append({
        "Model": model_name,
        "Component": "Alpha",
        "AnnualizedContribution": alpha_ann,
        "Type": "Residual alpha",
    })

    return summary, pd.DataFrame(loadings), pd.DataFrame(decomposition)


def _capm_inputs(
    portfolio_returns: pd.Series,
    benchmark_returns: pd.Series,
    risk_free_rate: float,
    trading_days: int,
) -> tuple[pd.Series, pd.DataFrame, pd.Series]:
    idx = portfolio_returns.dropna().index.intersection(benchmark_returns.dropna().index)
    p = portfolio_returns.loc[idx].astype(float)
    b = benchmark_returns.loc[idx].astype(float)
    rf_daily = (1 + float(risk_free_rate)) ** (1 / trading_days) - 1
    rf = pd.Series(rf_daily, index=idx, name="RF")
    y = p - rf
    x = pd.DataFrame({"BenchmarkExcess": b - rf}, index=idx)
    return y, x, rf


def _french_model(
    portfolio_returns: pd.Series,
    french: pd.DataFrame,
    factors: list[str],
) -> tuple[pd.Series, pd.DataFrame, pd.Series] | None:
    required = ["RF"] + factors
    if french.empty or any(c not in french.columns for c in required):
        return None
    idx = portfolio_returns.dropna().index.intersection(french.dropna(subset=required).index)
    if len(idx) < 30:
        return None
    p = portfolio_returns.loc[idx].astype(float)
    rf = french.loc[idx, "RF"].astype(float)
    y = p - rf
    x = french.loc[idx, factors].astype(float)
    return y, x, rf


def build_style_proxy_factors(
    proxy_returns: pd.DataFrame,
    risk_free_rate: float,
    trading_days: int,
) -> pd.DataFrame:
    """Build transparent public long/relative style proxy series.

    These are diagnostics, not academic factor returns. Market is SPY excess return when
    available; the other factors are ETF returns relative to SPY so that the regression asks
    whether the portfolio has incremental small-cap, value, growth, momentum, quality or
    low-volatility exposure beyond broad market beta.
    """
    if proxy_returns is None or proxy_returns.empty:
        return pd.DataFrame()
    p = proxy_returns.copy()
    rf_daily = (1 + float(risk_free_rate)) ** (1 / trading_days) - 1
    market_col = "Market_SPY" if "Market_SPY" in p.columns else next((c for c in p.columns if "market" in c.lower()), None)
    if not market_col:
        return pd.DataFrame()
    market = p[market_col]
    out = pd.DataFrame(index=p.index)
    out["Market"] = market - rf_daily
    mappings = {
        "SmallCap": "SmallCap_IWM",
        "Value": "Value_IWD",
        "Growth": "Growth_IWF",
        "Momentum": "Momentum_MTUM",
        "Quality": "Quality_QUAL",
        "LowVol": "LowVol_USMV",
    }
    for factor, source in mappings.items():
        if source in p.columns:
            out[factor] = p[source] - market
    return out


def _rolling_model(
    y: pd.Series,
    factors: pd.DataFrame,
    model_name: str,
    windows: dict[str, int],
    trading_days: int,
) -> pd.DataFrame:
    data = pd.concat([y.rename("Y"), factors], axis=1).dropna()
    rows = []
    for label, window in windows.items():
        if len(data) < window:
            continue
        # Weekly sampling keeps the output compact while preserving a useful rolling history.
        for end in range(window - 1, len(data), 5):
            sample = data.iloc[end - window + 1 : end + 1]
            summary, _, _ = _ols(sample["Y"], sample.drop(columns="Y"), trading_days, model_name)
            if not summary:
                continue
            beta_market = None
            factor_name = "BenchmarkExcess" if "BenchmarkExcess" in sample.columns else ("Mkt-RF" if "Mkt-RF" in sample.columns else None)
            if factor_name:
                # Refit coefficient cheaply to surface rolling market beta.
                X = np.column_stack([np.ones(len(sample)), sample.drop(columns="Y").to_numpy(dtype=float)])
                coef, _, _, _ = np.linalg.lstsq(X, sample["Y"].to_numpy(dtype=float), rcond=None)
                cols = list(sample.drop(columns="Y").columns)
                beta_market = float(coef[1 + cols.index(factor_name)]) if factor_name in cols else None
            rows.append({
                "Date": sample.index[-1],
                "Model": model_name,
                "Window": label,
                "AnnualizedAlpha": summary.get("AnnualizedAlpha"),
                "AlphaTStat": summary.get("AlphaTStat"),
                "AlphaPValue": summary.get("AlphaPValue"),
                "MarketBeta": beta_market,
                "R2": summary.get("R2"),
                "Observations": summary.get("Observations"),
            })
    return pd.DataFrame(rows)


def analyze_alpha(
    portfolio_returns: pd.Series,
    benchmark_returns: pd.Series,
    proxy_returns: pd.DataFrame | None,
    risk_free_rate: float = 0.04,
    trading_days: int = 252,
    rolling_windows: dict[str, int] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    """Return summary, loadings, rolling alpha, decomposition and metadata."""
    rolling_windows = rolling_windows or {"1Y": trading_days, "3Y": trading_days * 3}
    summaries = []
    loadings = []
    decompositions = []
    rolling_frames = []
    metadata = {"french_factor_warnings": [], "method_note": (
        "Alpha uses current portfolio weights applied to historical returns unless the project is supplied with point-in-time weights. "
        "CAPM is benchmark-relative Jensen alpha. FF3/Carhart/FF5 use Kenneth French daily research factors. "
        "The style-proxy model is a public ETF diagnostic, not an academic or commercial risk model."
    )}

    # Benchmark CAPM / Jensen alpha.
    y_capm, x_capm, rf_capm = _capm_inputs(portfolio_returns, benchmark_returns, risk_free_rate, trading_days)
    summary, loading, decomp = _ols(y_capm, x_capm, trading_days, "CAPM - Benchmark")
    if summary:
        summary["RiskFreeAnnualizedArithmetic"] = _annualize_mean(rf_capm, trading_days)
        summaries.append(summary); loadings.append(loading); decompositions.append(decomp)
        rolling_frames.append(_rolling_model(y_capm, x_capm, "CAPM - Benchmark", rolling_windows, trading_days))

    # Academic factor models from Kenneth French.
    french, warnings = load_french_daily_factors()
    metadata["french_factor_warnings"] = warnings
    academic = [
        ("Fama-French 3", ["Mkt-RF", "SMB", "HML"]),
        ("Carhart 4", ["Mkt-RF", "SMB", "HML", "Mom"]),
        ("Fama-French 5", ["Mkt-RF", "SMB", "HML", "RMW", "CMA"]),
    ]
    for model_name, factor_cols in academic:
        model = _french_model(portfolio_returns, french, factor_cols)
        if model is None:
            continue
        y, x, rf = model
        summary, loading, decomp = _ols(y, x, trading_days, model_name)
        if not summary:
            continue
        summary["RiskFreeAnnualizedArithmetic"] = _annualize_mean(rf, trading_days)
        summaries.append(summary); loadings.append(loading); decompositions.append(decomp)
        if model_name in {"Carhart 4"}:
            rolling_frames.append(_rolling_model(y, x, model_name, rolling_windows, trading_days))

    # Public style-proxy residual alpha, including growth / quality controls.
    proxy_factors = build_style_proxy_factors(proxy_returns if proxy_returns is not None else pd.DataFrame(), risk_free_rate, trading_days)
    if not proxy_factors.empty:
        idx = portfolio_returns.dropna().index.intersection(proxy_factors.dropna(how="any").index)
        if len(idx) >= 60:
            rf_daily = (1 + float(risk_free_rate)) ** (1 / trading_days) - 1
            y = portfolio_returns.loc[idx].astype(float) - rf_daily
            x = proxy_factors.loc[idx].astype(float)
            summary, loading, decomp = _ols(y, x, trading_days, "Public Style Proxy")
            if summary:
                summary["RiskFreeAnnualizedArithmetic"] = rf_daily * trading_days
                # Surface multicollinearity warning via condition number.
                try:
                    condition = float(np.linalg.cond(np.column_stack([np.ones(len(x)), x.to_numpy(dtype=float)])))
                except Exception:
                    condition = np.nan
                summary["ConditionNumber"] = condition if np.isfinite(condition) else None
                summaries.append(summary); loadings.append(loading); decompositions.append(decomp)

    summary_df = pd.DataFrame(summaries)
    loadings_df = pd.concat(loadings, ignore_index=True) if loadings else pd.DataFrame()
    rolling_df = pd.concat([x for x in rolling_frames if x is not None and not x.empty], ignore_index=True) if rolling_frames else pd.DataFrame()
    decomposition_df = pd.concat(decompositions, ignore_index=True) if decompositions else pd.DataFrame()

    # Add aggregate explained-vs-alpha rows and raw benchmark active return context.
    if not summary_df.empty and not decomposition_df.empty:
        aggregates = []
        for model in summary_df["Model"]:
            part = decomposition_df[decomposition_df["Model"] == model]
            factor_explained = float(part.loc[part["Type"] == "Factor explained", "AnnualizedContribution"].sum())
            alpha = float(part.loc[part["Type"] == "Residual alpha", "AnnualizedContribution"].sum())
            rf_ann = summary_df.loc[summary_df["Model"] == model, "RiskFreeAnnualizedArithmetic"].iloc[0]
            aggregates.extend([
                {"Model": model, "Component": "Risk-free", "AnnualizedContribution": rf_ann, "Type": "Baseline"},
                {"Model": model, "Component": "Total factor explained", "AnnualizedContribution": factor_explained, "Type": "Factor explained total"},
                {"Model": model, "Component": "Residual alpha", "AnnualizedContribution": alpha, "Type": "Residual alpha summary"},
            ])
        decomposition_df = pd.concat([decomposition_df, pd.DataFrame(aggregates)], ignore_index=True)

    return summary_df, loadings_df, rolling_df, decomposition_df, metadata
