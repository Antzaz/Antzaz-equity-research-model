from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
import sys

import numpy as np
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
from src.portfolio_optimization import build_portfolio_optimization
from src.decision_loop import (
    load_transaction_ledger,
    ledger_tickers,
    download_realized_market_data,
    build_realized_portfolio,
    aggregate_realized_attribution,
    brinson_sector_attribution,
    decision_journal_analytics,
    high_level_decision_learning,
    research_expected_return_bridge,
    position_sizing_ranges,
    thesis_budget,
    research_fundamental_scenarios,
    custom_thesis_scenarios,
    expected_return_decomposition,
    transaction_cost_rebalance,
)
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
ROOT = BASE.parent
ML_HISTORY_DB = ROOT / "ml_data" / "ml_history.sqlite"


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

    # Optional transaction history can contain securities no longer held. Include them in
    # market history so realized performance is not survivorship-truncated to today's holdings.
    ledger = load_transaction_ledger(BASE / "transaction_ledger.csv")
    historical_tickers = ledger_tickers(ledger)
    proxy_map = {k: str(v).upper() for k, v in config.get("factor_proxies", {}).items()}
    all_market_tickers = list(dict.fromkeys(tickers + historical_tickers + [benchmark] + list(proxy_map.values())))

    print("Downloading maximum available public market history...")
    prices = download_prices(all_market_tickers, period=config["history_period"])
    missing = [t for t in tickers + [benchmark] if t not in prices.columns]
    if missing:
        raise RuntimeError(f"Missing price history for: {', '.join(missing)}")

    print("Downloading company snapshots...")
    info = fetch_info(list(dict.fromkeys(tickers + historical_tickers)))

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

    # Offline research -> portfolio bridge. Manual expected_returns.csv remains authoritative
    # where supplied; model-derived values fill gaps and are confidence-shrunk first.
    decision_cfg = config.get("decision_loop", {})
    research_bridge = research_expected_return_bridge(
        tickers=tickers,
        root=ROOT,
        info=info,
        benchmark_expected_return=float(decision_cfg.get("benchmark_expected_return", 0.08)),
        convergence_years=float(decision_cfg.get("valuation_convergence_years", 3.0)),
    )

    # Point-in-time transaction accounting. This is intentionally separate from the existing
    # current-weight backcast analytics.
    realized_raw_prices = pd.DataFrame()
    realized_actions = {}
    realized_download_warnings = []
    realized = {
        "summary": {"status": "NO_LEDGER"},
        "timeseries": pd.DataFrame(),
        "weights": pd.DataFrame(),
        "attribution": pd.DataFrame(),
        "transactions": ledger,
        "warnings": [],
    }
    if not ledger.empty:
        print("Reconstructing point-in-time portfolio from transaction ledger...")
        realized_raw_prices, realized_actions, realized_download_warnings = download_realized_market_data(
            historical_tickers,
            period=config["history_period"],
        )
        realized = build_realized_portfolio(
            ledger=ledger,
            raw_prices=realized_raw_prices,
            adjusted_benchmark=prices[benchmark] if benchmark in prices else None,
            corporate_actions=realized_actions,
        )
        realized["warnings"] = list(realized.get("warnings", [])) + realized_download_warnings

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

    manual_expected_returns = load_expected_returns(BASE / "expected_returns.csv", tickers)
    research_expected_inputs = pd.DataFrame()
    if not research_bridge.empty:
        research_expected_inputs = research_bridge.dropna(subset=["ConfidenceAdjustedExpectedReturn"]).copy()
        if not research_expected_inputs.empty:
            research_expected_inputs = pd.DataFrame({
                "Ticker": research_expected_inputs["Ticker"],
                "ExpectedReturn": research_expected_inputs["ConfidenceAdjustedExpectedReturn"],
                "Conviction": research_expected_inputs["Confidence"] * 5.0,
                "MinWeight": np.nan,
                "MaxWeight": np.nan,
                "Thesis": "Confidence-shrunk expected return from latest local equity-research workbook",
                "SourceDate": datetime.now().date().isoformat(),
            })
    if manual_expected_returns.empty:
        expected_returns = research_expected_inputs.copy()
    elif research_expected_inputs.empty:
        expected_returns = manual_expected_returns.copy()
    else:
        # Explicit local manual inputs override the automatic research bridge ticker-by-ticker.
        expected_returns = pd.concat([manual_expected_returns, research_expected_inputs], ignore_index=True)
        expected_returns = expected_returns.drop_duplicates("Ticker", keep="first")

    # Legacy optimizer remains exported for backwards compatibility.
    optimizations = optimize_portfolios(
        asset_returns,
        holdings,
        expected_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        max_position=float(constraints_cfg.get("max_position", 0.25)),
    )

    # Maximum-data optimizer: up to 20Y local ML history, robust covariance, efficient
    # frontiers, confidence-shrunk expected-return ML and regime-aware risk.
    print("Running maximum-data classical + ML portfolio optimization...")
    optimizer_cfg = {
        **constraints_cfg,
        **config.get("portfolio_optimization", {}),
    }
    portfolio_opt = build_portfolio_optimization(
        asset_returns=asset_returns,
        benchmark_returns=benchmark_returns,
        holdings=holdings,
        manual_expected_returns=expected_returns,
        risk_free_rate=float(config["risk_free_rate"]),
        config=optimizer_cfg,
        history_db=ML_HISTORY_DB,
    )
    opt_meta = portfolio_opt.get("metadata", {})
    print(
        "Portfolio optimizer: "
        f"status={opt_meta.get('status')}, history_assets={opt_meta.get('history_assets', 0)}, "
        f"regime={(opt_meta.get('regime') or {}).get('regime') or 'unavailable'}"
    )

    # Close the research -> sizing -> decision -> outcome loop.
    sizing = position_sizing_ranges(
        holdings,
        research_bridge,
        max_position=float(constraints_cfg.get("max_position", 0.25)),
        half_width=float(decision_cfg.get("position_range_half_width", 0.025)),
        liquidity=liquidity,
        max_days_to_liquidate=float(constraints_cfg.get("max_days_to_liquidate", 5.0)),
    )
    thesis_budget_table = thesis_budget(holdings, research_bridge)
    fundamental_scenarios = research_fundamental_scenarios(holdings, research_bridge)
    custom_fundamental_scenarios = custom_thesis_scenarios(
        BASE / "fundamental_scenarios.csv",
        holdings,
    )
    expected_return_detail, expected_return_summary = expected_return_decomposition(holdings, research_bridge)
    rebalance_costs = transaction_cost_rebalance(
        optimizer_weights=portfolio_opt.get("weights"),
        holdings=holdings,
        liquidity=liquidity,
        research_bridge=research_bridge,
        portfolio_name=str(decision_cfg.get("rebalance_target_portfolio", "Regime-Aware ML Maximum Sharpe")),
        rebalance_threshold=float(constraints_cfg.get("rebalance_threshold", 0.03)),
        spread_bps=float(decision_cfg.get("estimated_spread_bps", 5.0)),
        impact_bps_at_10pct_adv=float(decision_cfg.get("impact_bps_at_10pct_adv", 20.0)),
        min_benefit_cost_ratio=float(decision_cfg.get("minimum_benefit_cost_ratio", 1.5)),
    )

    decision_detail, decision_summary = decision_journal_analytics(
        BASE / "portfolio_decision_journal.csv",
        prices,
        benchmark,
    )
    decision_learning = high_level_decision_learning(decision_detail)

    realized_attribution = aggregate_realized_attribution(realized.get("attribution", pd.DataFrame()))
    sector_map = {t: (info.get(t) or {}).get("sector") or "Unknown" for t in historical_tickers}
    brinson = brinson_sector_attribution(
        realized.get("weights", pd.DataFrame()),
        realized_raw_prices,
        sector_map,
        BASE / "benchmark_sector_history.csv",
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
        "portfolio_optimization": opt_meta,
        "realized_performance": realized.get("summary", {}),
        "decision_loop": {
            "research_bridge_coverage": int((research_bridge.get("SourceStatus") == "PASS").sum()) if not research_bridge.empty and "SourceStatus" in research_bridge else 0,
            "transaction_ledger_rows": int(len(ledger)),
            "realized_warnings": realized.get("warnings", []),
            "decision_journal_rows": int(len(decision_detail)),
            "brinson_status": "PASS" if not brinson.empty else "REVIEW — add point-in-time benchmark_sector_history.csv",
            "method_note": "Research-model expected returns are confidence-shrunk; realized performance requires explicit transactions and external cash flows; Brinson requires point-in-time benchmark sector history.",
        },
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
        "manual_expected_returns": manual_expected_returns,
        "research_expected_return_bridge": research_bridge,
        "expected_returns_inputs": expected_returns,
        "position_sizing_ranges": sizing,
        "thesis_budget": thesis_budget_table,
        "fundamental_research_scenarios": fundamental_scenarios,
        "custom_thesis_scenarios": custom_fundamental_scenarios,
        "expected_return_decomposition": expected_return_detail,
        "expected_return_decomposition_summary": expected_return_summary,
        "rebalance_cost_analysis": rebalance_costs,
        "realized_portfolio_timeseries": realized.get("timeseries"),
        "point_in_time_weights": realized.get("weights"),
        "realized_security_attribution": realized_attribution,
        "realized_daily_attribution": realized.get("attribution"),
        "decision_journal_outcomes": decision_detail,
        "decision_journal_summary": decision_summary,
        "decision_learning": decision_learning,
        "brinson_sector_attribution": brinson,
        "portfolio_optimizations": optimizations,
        "optimizer_summary": portfolio_opt.get("summary"),
        "optimizer_weights": portfolio_opt.get("weights"),
        "efficient_frontier": portfolio_opt.get("frontier"),
        "optimizer_expected_returns": portfolio_opt.get("expected_returns"),
        "optimizer_covariance": portfolio_opt.get("covariance"),
        "optimizer_regime_covariance": portfolio_opt.get("regime_covariance"),
        "optimizer_correlation": portfolio_opt.get("correlation"),
        "optimizer_data_coverage": portfolio_opt.get("coverage"),
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
    if manual_expected_returns.empty:
        print("Manual expected_returns.csv is optional: local equity-research workbooks now feed confidence-shrunk expected returns into the optimizer when available.")
    if ledger.empty:
        print("Optional: copy transaction_ledger_template.csv to transaction_ledger.csv for true TWR/MWR and point-in-time attribution.")
    if brinson.empty:
        print("Optional: add benchmark_sector_history.csv for point-in-time Brinson allocation/selection/interaction attribution.")
    print("Next:")
    print("  python -m streamlit run dashboard.py")
    print("  Open the Portfolio Optimization page for efficient-frontier and ML allocation charts.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
