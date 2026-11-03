from __future__ import annotations

"""Rebuild a clearly labeled core portfolio dashboard from LOCAL cached market data.

This is a safe recovery path, not a substitute for the full institutional research run:
- never makes network calls, alters holdings, or invents unavailable benchmark figures;
- computes a current-weight historical *model backcast*, not realized portfolio returns;
- saves enough holdings, portfolio, sector, correlation and risk output for the standard
  Streamlit dashboard when its private outputs/latest folder is missing;
- never replaces an existing full analysis, and validates availability before export.
"""

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path

import numpy as np
import pandas as pd

from src.data import cached_prices, price_cache_status
from src.export import write_outputs
from src.portfolio import load_portfolio, build_holdings
from src.professional_portfolio import concentration_metrics, risk_budget_table


ROOT = Path(__file__).resolve().parent
PORTFOLIO = ROOT / "portfolio.csv"
LATEST = ROOT / "outputs" / "latest"
BENCHMARK = "SPY"
MODE = "RECOVERY_CACHED_PRICES"


def _valid_number(value):
    try:
        v = float(value)
        return v if np.isfinite(v) else None
    except (TypeError, ValueError):
        return None


def _cached_company_info(base: Path) -> dict:
    info_file = base / ".offline_cache" / "company_info.json"
    if not info_file.is_file():
        return {}
    try:
        data = json.loads(info_file.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _summary_statistics(returns: pd.Series, risk_free_rate: float = 0.04) -> dict:
    r = returns.dropna()
    if len(r) < 20:
        return {
            "observations": int(len(r)),
            "annualized_return": None,
            "annualized_volatility": None,
            "sharpe": None,
            "sortino": None,
            "max_drawdown": None,
        }
    growth = (1 + r).cumprod()
    ann_ret = float(growth.iloc[-1] ** (252.0 / len(r)) - 1)
    vol = float(r.std(ddof=1) * np.sqrt(252))
    dd = float((growth / growth.cummax() - 1).min())
    downside = r[r < 0]
    down_vol = float(downside.std(ddof=1) * np.sqrt(252)) if len(downside) > 1 else 0.0
    value_at_risk = float(r.quantile(.05))
    tail = r[r <= value_at_risk]
    return {
        "observations": int(len(r)),
        "annualized_return": ann_ret,
        "annualized_volatility": vol,
        "sharpe": (ann_ret - risk_free_rate) / vol if vol > 0 else None,
        "sortino": (ann_ret - risk_free_rate) / down_vol if down_vol > 0 else None,
        "max_drawdown": dd,
        "daily_var_95": value_at_risk,
        "daily_expected_shortfall_95": float(tail.mean()) if not tail.empty else None,
    }


def recover(
    base: Path = ROOT,
    *,
    prices: pd.DataFrame | None = None,
    min_observations: int = 3,
) -> dict:
    base = Path(base)
    portfolio_file = base / "portfolio.csv"
    if not portfolio_file.is_file():
        raise RuntimeError("Portfolio CSV is missing from this computer")
    portfolio = load_portfolio(portfolio_file)
    tickers = list(portfolio["Ticker"])
    if not tickers:
        raise RuntimeError("Portfolio has no holdings")

    if prices is None:
        if base.resolve() == ROOT.resolve():
            prices = cached_prices()
        else:
            file = base / ".offline_cache" / "adjusted_closes.csv"
            if not file.is_file():
                raise RuntimeError("Local price cache does not exist")
            prices = pd.read_csv(file, index_col=0, parse_dates=True)
    if prices is None or prices.empty:
        raise RuntimeError("Local price cache is empty; run an online sync once")
    prices = prices.copy()
    prices.index = pd.to_datetime(prices.index, errors="coerce", utc=True).tz_localize(None)
    prices = prices[~prices.index.isna()].sort_index()
    prices.columns = [str(c).strip().upper() for c in prices.columns]
    prices = prices.apply(pd.to_numeric, errors="coerce")
    missing = [t for t in tickers if t not in prices or prices[t].dropna().empty]
    if missing:
        raise RuntimeError(
            f"Locally saved price cache is missing {len(missing)} current holding(s)"
        )
    observed = prices[tickers].dropna(how="all")
    if len(observed) < min_observations:
        raise RuntimeError("Cached history is too short even for a basic portfolio dashboard")

    # Each position uses its own most recently observed market close. Report gaps,
    # and never create zero/placeholder prices for a missing ticker.
    current = prices[tickers].ffill().iloc[-1]
    if current.isna().any() or (current <= 0).any():
        raise RuntimeError("One or more cached holding prices are invalid")
    holdings = build_holdings(portfolio, current, _cached_company_info(base))
    if not holdings["Weight"].notna().all():
        raise RuntimeError("Could not calculate complete portfolio weights")

    # Compare securities on the same actual observation dates. This is a static-weight
    # backcast, not realized returns. No forward-fill across listing-history gaps.
    aligned = observed.dropna(how="any")
    returns = aligned.pct_change(fill_method=None).dropna(how="any")
    if len(returns) < 2:
        raise RuntimeError("Fewer than two common daily return observations in local cache")
    weight_by_ticker = holdings.set_index("Ticker")["Weight"].reindex(tickers).astype(float)
    modeled = returns.mul(weight_by_ticker, axis="columns").sum(axis=1)
    summary = _summary_statistics(modeled)
    summary["weight_method"] = holdings.attrs.get("weight_method", "unknown")
    summary["holdings"] = len(tickers)
    summary["benchmark"] = BENCHMARK
    # The portfolio mixes at least one non-USD listing with US securities.
    # Summing local listing currencies would invent a single-currency NAV. The
    # user's locally supplied quantities and currency-specific prices are shown
    # per holding; a cross-currency NAV requires historical/current FX.
    currencies = holdings["Currency"].dropna().astype(str).str.strip()
    currency_uniform = (
        len(currencies) == len(holdings)
        and currencies.nunique() == 1
        and currencies.iloc[0] != ""
    )
    summary["current_market_value"] = (
        _valid_number(holdings["MarketValue"].sum())
        if currency_uniform else None
    )

    bench = pd.Series(dtype=float)
    if BENCHMARK in prices.columns:
        bench = prices[BENCHMARK].pct_change(fill_method=None).reindex(modeled.index)
    benchmark_available = bench.notna().sum() >= 20
    if benchmark_available:
        matched = pd.DataFrame({"portfolio": modeled, "benchmark": bench}).dropna()
        if len(matched) >= 20:
            p, b = matched["portfolio"], matched["benchmark"]
            active = p - b
            te = float(active.std(ddof=1) * np.sqrt(252))
            b_variance = float(b.var(ddof=1))
            summary.update({
                "benchmark_annualized_return": float((1 + b).prod() ** (252 / len(b)) - 1),
                "active_annualized_return": float((1 + p).prod() ** (252 / len(p)) - 1)
                - float((1 + b).prod() ** (252 / len(b)) - 1),
                "tracking_error": te,
                "information_ratio": (
                    (float((1 + p).prod() ** (252 / len(p)) - 1)
                     - float((1 + b).prod() ** (252 / len(b)) - 1)) / te
                    if te > 0 else None
                ),
                "beta": float(p.cov(b) / b_variance) if b_variance > 0 else None,
                "benchmark_correlation": _valid_number(p.corr(b)),
                "daily_active_hit_rate": float((p > b).mean()),
            })
        else:
            benchmark_available = False

    cov = returns.cov() * 252.0
    w = weight_by_ticker.to_numpy()
    variance = float(w @ cov.to_numpy() @ w)
    vol = float(np.sqrt(max(variance, 0.0)))
    if vol > 0:
        marginal = cov.to_numpy() @ w / vol
        component = w * marginal
        rc = component / vol
    else:
        marginal = component = rc = np.full(len(tickers), np.nan)
    risk = pd.DataFrame({
        "Ticker": tickers,
        "Weight": w,
        "MarginalRisk": marginal,
        "ComponentRisk": component,
        "RiskContributionPct": rc,
    })
    holdings = holdings.merge(risk[["Ticker", "RiskContributionPct"]], on="Ticker", how="left")
    hold_return = returns
    asset_risk = pd.DataFrame({
        "Ticker": tickers,
        "AnnualizedVolatility": hold_return.std(ddof=1).to_numpy() * np.sqrt(252),
        "Beta": [None] * len(tickers),
        "HistoricalReturn": (1 + hold_return).prod().to_numpy() - 1,
        "MaxDrawdown": [
            float((((1 + hold_return[t]).cumprod()) /
                ((1 + hold_return[t]).cumprod()).cummax() - 1).min())
            for t in tickers
        ],
    })
    holdings = holdings.merge(asset_risk, on="Ticker", how="left")
    summary.update(concentration_metrics(holdings))

    timeseries = pd.DataFrame({
        "Date": modeled.index,
        "PortfolioReturn": modeled.to_numpy(),
        "PortfolioGrowth": (1 + modeled).cumprod().to_numpy(),
    })
    if benchmark_available:
        timeseries["BenchmarkReturn"] = bench.to_numpy()
        timeseries["BenchmarkGrowth"] = (1 + bench.fillna(0)).cumprod().to_numpy()
        timeseries["ActiveReturn"] = modeled.to_numpy() - bench.to_numpy()
    else:
        timeseries["BenchmarkReturn"] = np.nan
        timeseries["BenchmarkGrowth"] = np.nan
        timeseries["ActiveReturn"] = np.nan

    sector = (
        holdings.assign(Sector=holdings["Sector"].fillna("Unknown"))
        .groupby("Sector", as_index=False)["Weight"].sum()
        .sort_values("Weight", ascending=False)
    )
    covariance = cov.reset_index().rename(columns={"index": "Ticker"})
    correlation = returns.corr().reset_index().rename(columns={"index": "Ticker"})
    input_path = base / "outputs" / "latest" / "summary.json"
    if input_path.exists():
        # Existing complete/latest analyses are always more authoritative than a
        # lightweight recovery run, including an earlier recovery snapshot.
        try:
            existing = json.loads(input_path.read_text(encoding="utf-8"))
            if existing.get("analysis_mode") != MODE:
                return {"status": "SKIPPED_EXISTING_FULL_ANALYSIS", "holdings": len(tickers)}
        except (OSError, ValueError):
            raise RuntimeError("Existing analysis summary is unreadable; repair it manually")

    table_map = {
        "holdings_analysis": holdings,
        "portfolio_timeseries": timeseries,
        "sector_exposure": sector,
        "correlation_matrix": correlation,
        "covariance_matrix": covariance,
        "risk_contribution": risk,
        "risk_budget": risk_budget_table(holdings),
        "concentration_summary": pd.DataFrame([
            {"Metric": key, "Value": value}
            for key, value in summary.items() if key.startswith(("top_", "effective_", "herfindahl_"))
        ]),
    }
    market_date = observed.index.max().date().isoformat()
    report = {
        "analysis_mode": MODE,
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "price_as_of": market_date,
        "portfolio": summary,
        "data_quality": {
            "status": "RECOVERY_ONLY",
            "method": "Current-weight historic return backcast from cached adjusted closes",
            "market_date": market_date,
            "benchmark_available": bool(benchmark_available),
            "single_reporting_currency": bool(currency_uniform),
            "advanced_models": "NOT_RUN",
            "realized_portfolio_returns": "NOT_CALCULATED",
            "warnings": [
                "This is a cached-price recovery snapshot, NOT the full research pipeline.",
                "Historical returns apply present weights to historical adjusted prices.",
                "No historical transaction ledger or cash flow has been inferred.",
                "Unobserved benchmark statistics and advanced outputs remain unavailable.",
                "Holding market values are quoted in listing currencies; without full FX "
                "normalization, portfolio weights are only a listing-currency approximation.",
            ],
        },
        "portfolio_optimization": {"status": "NOT_RUN_RECOVERY_ONLY"},
        "realized_performance": {"status": "NO_LEDGER"},
        "decision_loop": {"method_note": "Advanced decision outputs need an online full research build."},
    }
    _, output = write_outputs(base / "outputs", table_map, report)
    if not (output / "summary.json").is_file():
        raise RuntimeError("Recovery exporter did not produce the required summary.json")
    if not (output / "holdings_analysis.csv").is_file():
        raise RuntimeError("Recovery exporter did not produce holdings_analysis.csv")
    return {
        "status": "RECOVERED_OFFLINE",
        "holdings": len(tickers),
        "common_return_days": len(modeled),
        "price_as_of": market_date,
        "benchmark_available": bool(benchmark_available),
    }


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Rebuild basic local dashboard from already cached prices")
    p.add_argument("--base", type=Path, default=ROOT)
    args = p.parse_args(argv)
    status = recover(args.base)
    print(json.dumps(status, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
