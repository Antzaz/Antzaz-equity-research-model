from __future__ import annotations

import json
from pathlib import Path
import sys

import pandas as pd

from src.data import download_prices, fetch_info
from src.portfolio import load_portfolio, build_holdings
from src.risk import portfolio_risk, per_asset_risk
from src.factors import build_factor_scores, factor_exposure
from src.monte_carlo import bootstrap_portfolio
from src.reverse_dcf import reverse_dcf_table
from src.stress import beta_stress_test
from src.forecast_tracker import analyze_forecasts
from src.export import write_outputs


BASE = Path(__file__).resolve().parent


def main():
    with open(BASE / "config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    portfolio = load_portfolio(BASE / "portfolio.csv")
    benchmark = config["benchmark"].upper()
    tickers = portfolio["Ticker"].tolist()

    print("Downloading market data...")
    prices = download_prices(tickers + [benchmark], period=config["history_period"])
    missing = [t for t in tickers + [benchmark] if t not in prices.columns]
    if missing:
        raise RuntimeError(f"Missing price history for: {', '.join(missing)}")

    print("Downloading company snapshots...")
    info = fetch_info(tickers)

    current_prices = prices[tickers].ffill().iloc[-1]
    holdings = build_holdings(portfolio, current_prices, info)
    weight_method = holdings.attrs.get("weight_method", "unknown")

    asset_prices = prices[tickers]
    asset_returns = asset_prices.pct_change()
    benchmark_returns = prices[benchmark].pct_change().dropna()
    weights = holdings.set_index("Ticker")["Weight"]

    risk_summary, portfolio_returns, covariance, risk_contribution = portfolio_risk(
        asset_returns=asset_returns,
        weights=weights,
        benchmark_returns=benchmark_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        trading_days=int(config["trading_days"]),
    )

    asset_risk = per_asset_risk(
        prices=prices[tickers + [benchmark]],
        benchmark=benchmark,
        trading_days=int(config["trading_days"]),
    )
    holdings = holdings.merge(asset_risk, on="Ticker", how="left")
    holdings = holdings.merge(
        risk_contribution[["Ticker", "RiskContributionPct"]],
        on="Ticker",
        how="left",
    )

    factors = build_factor_scores(
        holdings=holdings,
        prices=asset_prices,
        info=info,
        factor_weights=config["factor_weights"],
        trading_days=int(config["trading_days"]),
    )
    factor_portfolio = factor_exposure(factors)

    mc_cfg = config["monte_carlo"]
    mc_summary, mc_distribution = bootstrap_portfolio(
        asset_returns=asset_returns,
        weights=weights,
        simulations=int(mc_cfg["simulations"]),
        horizon_days=int(mc_cfg["horizon_days"]),
        seed=int(mc_cfg["seed"]),
    )

    reverse_dcf = reverse_dcf_table(
        tickers=tickers,
        info=info,
        assumptions=config["reverse_dcf"],
    )

    stress = beta_stress_test(
        holdings=holdings,
        asset_risk=asset_risk,
        scenarios=config["stress_scenarios"],
    )

    forecasts = analyze_forecasts(BASE / "forecasts.csv")

    correlation = asset_returns.corr()
    correlation_out = correlation.reset_index().rename(columns={"index": "Ticker"})
    covariance_out = covariance.reset_index().rename(columns={"index": "Ticker"})

    port_series = pd.DataFrame({
        "Date": portfolio_returns.index,
        "PortfolioReturn": portfolio_returns.values,
        "PortfolioGrowth": (1 + portfolio_returns).cumprod().values,
    })
    bench_common = benchmark_returns.reindex(portfolio_returns.index)
    port_series["BenchmarkReturn"] = bench_common.values
    port_series["BenchmarkGrowth"] = (1 + bench_common.fillna(0)).cumprod().values

    sector = (
        holdings.assign(Sector=holdings["Sector"].fillna("Unknown"))
        .groupby("Sector", as_index=False)["Weight"].sum()
        .sort_values("Weight", ascending=False)
    )

    summaries = {
        "portfolio": {
            **risk_summary,
            "benchmark": benchmark,
            "holdings": int(len(holdings)),
            "weight_method": weight_method,
            "current_market_value": float(holdings["MarketValue"].sum())
                if holdings["MarketValue"].notna().all() else None,
        },
        "monte_carlo": mc_summary,
        "reverse_dcf_assumptions": config["reverse_dcf"],
    }

    tables = {
        "holdings_analysis": holdings,
        "portfolio_timeseries": port_series,
        "correlation_matrix": correlation_out,
        "covariance_matrix": covariance_out,
        "risk_contribution": risk_contribution,
        "factor_scores": factors,
        "factor_exposure": factor_portfolio,
        "sector_exposure": sector,
        "monte_carlo_distribution": mc_distribution,
        "reverse_dcf": reverse_dcf,
        "stress_tests": stress,
        "forecast_accuracy": forecasts,
    }

    snapshot, latest = write_outputs(BASE / "outputs", tables, summaries)

    print("Analysis complete.")
    print(f"Latest outputs: {latest}")
    print(f"Snapshot:       {snapshot}")
    print()
    print("Next:")
    print("  python -m streamlit run dashboard.py")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
