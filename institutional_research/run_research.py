from __future__ import annotations

import json
from pathlib import Path
import sys

import pandas as pd

from src.data import download_prices, fetch_info
from src.portfolio import load_portfolio, build_holdings
from src.risk import portfolio_risk, per_asset_risk
from src.factors import build_factor_scores, factor_exposure
from src.alpha_analysis import analyze_alpha
from src.monte_carlo import bootstrap_portfolio
from src.reverse_dcf import reverse_dcf_table
from src.stress import beta_stress_test
from src.forecast_tracker import analyze_forecasts
from src.export import write_outputs
from src.professional_portfolio import (
    benchmark_relative_metrics,
    concentration_metrics,
    risk_budget_table,
    liquidity_analysis,
    factor_proxy_sensitivity,
    rolling_risk_table,
    historical_stress_windows,
    static_return_attribution,
    load_expected_returns,
    optimize_portfolios,
    constraint_report,
    active_share_from_file,
)


BASE = Path(__file__).resolve().parent


def main():
    with open(BASE / "config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    portfolio_path = BASE / "portfolio.csv"
    if not portfolio_path.exists():
        template = BASE / "portfolio_template.csv"
        if template.exists():
            import shutil
            shutil.copyfile(template, portfolio_path)
        raise ValueError(
            "portfolio.csv was not found. A local copy has been created from "
            "portfolio_template.csv. Add your holdings to portfolio.csv and run again."
        )

    portfolio = load_portfolio(portfolio_path)
    benchmark = config["benchmark"].upper()
    tickers = portfolio["Ticker"].tolist()
    proxy_map = {k: str(v).upper() for k, v in config.get("factor_proxies", {}).items()}
    all_market_tickers = list(dict.fromkeys(tickers + [benchmark] + list(proxy_map.values())))

    print("Downloading market data...")
    prices = download_prices(all_market_tickers, period=config["history_period"])
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

    # Professional portfolio-construction layer.
    relative = benchmark_relative_metrics(
        portfolio_returns,
        benchmark_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        trading_days=int(config["trading_days"]),
    )
    concentration = concentration_metrics(holdings)
    risk_budget = risk_budget_table(holdings)

    constraints_cfg = config.get("portfolio_constraints", {})
    liquidity = liquidity_analysis(
        holdings,
        info,
        participation_rate=float(constraints_cfg.get("liquidity_participation_rate", 0.10)),
    )

    proxy_available = {
        name: ticker for name, ticker in proxy_map.items() if ticker in prices.columns
    }
    proxy_returns = pd.DataFrame()
    if proxy_available:
        proxy_prices = prices[list(dict.fromkeys(proxy_available.values()))]
        proxy_returns_raw = proxy_prices.pct_change()
        proxy_returns = pd.DataFrame(index=proxy_returns_raw.index)
        for name, ticker in proxy_available.items():
            proxy_returns[name] = proxy_returns_raw[ticker]
        factor_proxy = factor_proxy_sensitivity(
            portfolio_returns,
            proxy_returns,
            trading_days=int(config["trading_days"]),
        )
    else:
        factor_proxy = pd.DataFrame()

    # Alpha / factor-adjusted return layer.
    print("Running CAPM and multi-factor alpha analysis...")
    alpha_cfg = config.get("alpha_analysis", {})
    rolling_windows = alpha_cfg.get(
        "rolling_windows",
        {"1Y": int(config["trading_days"]), "3Y": int(config["trading_days"]) * 3},
    )
    rolling_windows = {str(k): int(v) for k, v in rolling_windows.items() if int(v) > 30}
    alpha_summary, alpha_loadings, rolling_alpha, alpha_decomposition, alpha_metadata = analyze_alpha(
        portfolio_returns=portfolio_returns,
        benchmark_returns=benchmark_returns,
        proxy_returns=proxy_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        trading_days=int(config["trading_days"]),
        rolling_windows=rolling_windows,
    )
    if not alpha_summary.empty:
        alpha_summary["RawActiveAnnualizedReturn"] = relative.get("active_annualized_return")
        capm = alpha_summary[alpha_summary["Model"] == "CAPM - Benchmark"]
        if not capm.empty:
            row = capm.iloc[0]
            risk_summary["annualized_alpha"] = row.get("AnnualizedAlpha")
            risk_summary["alpha_t_stat"] = row.get("AlphaTStat")
            risk_summary["alpha_p_value"] = row.get("AlphaPValue")
            risk_summary["alpha_r_squared"] = row.get("R2")
            risk_summary["alpha_method"] = "Jensen/CAPM regression vs configured benchmark"

    rolling_risk = rolling_risk_table(
        portfolio_returns,
        benchmark_returns,
        trading_days=int(config["trading_days"]),
    )
    historical_stress = historical_stress_windows(
        portfolio_returns,
        benchmark_returns,
        config.get("historical_stress_windows", []),
    )
    attribution = static_return_attribution(asset_returns, weights)

    expected_returns = load_expected_returns(BASE / "expected_returns.csv", tickers)
    optimizations = optimize_portfolios(
        asset_returns,
        holdings,
        expected_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        max_position=float(constraints_cfg.get("max_position", 0.25)),
    )
    constraints = constraint_report(
        holdings,
        risk_summary,
        relative,
        liquidity,
        constraints_cfg,
    )

    active_share_detail = active_share_from_file(
        weights,
        BASE / "benchmark_weights.csv",
    )
    active_share = (
        active_share_detail.attrs.get("active_share")
        if not active_share_detail.empty
        else None
    )

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
    port_series["ActiveReturn"] = (
        port_series["PortfolioReturn"] - port_series["BenchmarkReturn"]
    )

    sector = (
        holdings.assign(Sector=holdings["Sector"].fillna("Unknown"))
        .groupby("Sector", as_index=False)["Weight"].sum()
        .sort_values("Weight", ascending=False)
    )

    summaries = {
        "portfolio": {
            **risk_summary,
            **relative,
            **concentration,
            "benchmark": benchmark,
            "holdings": int(len(holdings)),
            "weight_method": weight_method,
            "active_share": active_share,
            "current_market_value": float(holdings["MarketValue"].sum())
                if holdings["MarketValue"].notna().all() else None,
        },
        "monte_carlo": mc_summary,
        "reverse_dcf_assumptions": config["reverse_dcf"],
        "portfolio_constraints": constraints_cfg,
        "alpha_analysis": alpha_metadata,
        "factor_proxy_note": (
            "ETF proxy sensitivities are public-data diagnostics, not a commercial "
            "multi-factor risk model."
        ),
        "attribution_note": (
            "Return attribution and alpha use static current weights unless point-in-time "
            "portfolio weights / transactions are supplied; they are research diagnostics, "
            "not realized manager-performance attribution."
        ),
    }

    concentration_table = pd.DataFrame(
        [{"Metric": k, "Value": v} for k, v in concentration.items()]
    )
    relative_table = pd.DataFrame(
        [{"Metric": k, "Value": v} for k, v in relative.items()]
    )

    tables = {
        "holdings_analysis": holdings,
        "portfolio_timeseries": port_series,
        "correlation_matrix": correlation_out,
        "covariance_matrix": covariance_out,
        "risk_contribution": risk_contribution,
        "risk_budget": risk_budget,
        "concentration_summary": concentration_table,
        "benchmark_relative": relative_table,
        "liquidity_analysis": liquidity,
        "rolling_risk": rolling_risk,
        "historical_stress_windows": historical_stress,
        "return_attribution": attribution,
        "alpha_summary": alpha_summary,
        "alpha_factor_loadings": alpha_loadings,
        "rolling_alpha": rolling_alpha,
        "alpha_return_decomposition": alpha_decomposition,
        "factor_scores": factors,
        "factor_exposure": factor_portfolio,
        "factor_proxy_sensitivity": factor_proxy,
        "sector_exposure": sector,
        "monte_carlo_distribution": mc_distribution,
        "reverse_dcf": reverse_dcf,
        "stress_tests": stress,
        "forecast_accuracy": forecasts,
        "expected_returns_inputs": expected_returns,
        "portfolio_optimizations": optimizations,
        "constraint_report": constraints,
        "active_share_detail": active_share_detail,
    }

    snapshot, latest = write_outputs(BASE / "outputs", tables, summaries)

    print("Analysis complete.")
    print(f"Latest outputs: {latest}")
    print(f"Snapshot:       {snapshot}")
    if alpha_metadata.get("french_factor_warnings"):
        print("Alpha note: some Kenneth French factor downloads were unavailable:")
        for warning in alpha_metadata["french_factor_warnings"]:
            print(f"  - {warning}")
        print("CAPM and any available factor models were still exported.")
    print()
    if active_share is None:
        print("Optional: add benchmark_weights.csv to calculate true Active Share.")
    if expected_returns.empty:
        print("Optional: add expected_returns.csv to enable expected-return / max-Sharpe sizing.")
    print("Next:")
    print("  python -m streamlit run dashboard.py")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
