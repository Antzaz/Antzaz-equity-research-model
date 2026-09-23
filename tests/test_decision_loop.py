from __future__ import annotations

from pathlib import Path
import sys

import numpy as np
import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
IR = ROOT / "institutional_research"
if str(IR) not in sys.path:
    sys.path.insert(0, str(IR))

from src.decision_loop import (
    build_realized_portfolio,
    custom_thesis_scenarios,
    expected_return_decomposition,
    position_sizing_ranges,
    thesis_budget,
    transaction_cost_rebalance,
)


def test_realized_portfolio_reconstructs_point_in_time_nav_and_twr():
    dates = pd.to_datetime(["2026-01-02", "2026-01-05", "2026-01-06"])
    prices = pd.DataFrame({"AAA": [100.0, 110.0, 99.0]}, index=dates)
    benchmark = pd.Series([100.0, 105.0, 103.0], index=dates)
    ledger = pd.DataFrame([
        {
            "Date": pd.Timestamp("2026-01-02"), "Ticker": "", "Action": "DEPOSIT",
            "Shares": np.nan, "Price": np.nan, "Fees": 0.0, "CashFlow": 10000.0,
            "Currency": "USD", "DecisionID": "FUND", "ReferencePrice": np.nan,
            "SplitRatio": np.nan, "Notes": "",
        },
        {
            "Date": pd.Timestamp("2026-01-02"), "Ticker": "AAA", "Action": "BUY",
            "Shares": 10.0, "Price": 100.0, "Fees": 0.0, "CashFlow": np.nan,
            "Currency": "USD", "DecisionID": "D1", "ReferencePrice": 99.0,
            "SplitRatio": np.nan, "Notes": "",
        },
    ])

    out = build_realized_portfolio(ledger, prices, adjusted_benchmark=benchmark, corporate_actions={})
    assert out["summary"]["status"] == "PASS"
    ts = out["timeseries"]
    assert list(ts["NAV"].round(2)) == [10000.0, 10100.0, 9990.0]
    # Day 2: 90% cash + 10% stock, stock +10% => portfolio +1%.
    assert ts.loc[1, "PortfolioReturn"] == pytest.approx(0.01)
    assert out["summary"]["twr_total"] == pytest.approx((1.01 * (9990.0 / 10100.0)) - 1.0)
    latest = out["weights"][out["weights"]["Date"] == dates[-1]]
    aaa = latest[latest["Ticker"] == "AAA"].iloc[0]
    assert aaa["Shares"] == pytest.approx(10.0)
    assert aaa["MarketValue"] == pytest.approx(990.0)
    assert out["summary"]["implementation_shortfall"] == pytest.approx(10.0)


def _portfolio_inputs():
    holdings = pd.DataFrame([
        {"Ticker": "AAA", "Weight": 0.60, "RiskContributionPct": 0.70, "MarketValue": 60000.0},
        {"Ticker": "BBB", "Weight": 0.40, "RiskContributionPct": 0.30, "MarketValue": 40000.0},
    ])
    bridge = pd.DataFrame([
        {
            "Ticker": "AAA", "CurrentPrice": 100.0, "BearValue": 75.0,
            "ExpectedAlpha": 0.10, "Confidence": 0.80,
            "ConfidenceAdjustedExpectedReturn": 0.14,
            "BaseGrowthProxy": 0.08, "DividendYield": 0.01, "NetBuybackYield": 0.02,
        },
        {
            "Ticker": "BBB", "CurrentPrice": 50.0, "BearValue": 45.0,
            "ExpectedAlpha": 0.03, "Confidence": 0.60,
            "ConfidenceAdjustedExpectedReturn": 0.09,
            "BaseGrowthProxy": 0.04, "DividendYield": 0.02, "NetBuybackYield": 0.01,
        },
    ])
    liquidity = pd.DataFrame([
        {"Ticker": "AAA", "EstimatedDaysToLiquidate": 1.0, "AverageDailyDollarVolume": 5_000_000.0},
        {"Ticker": "BBB", "EstimatedDaysToLiquidate": 2.0, "AverageDailyDollarVolume": 1_000_000.0},
    ])
    return holdings, bridge, liquidity


def test_research_to_sizing_and_expected_return_budget_is_coherent():
    holdings, bridge, liquidity = _portfolio_inputs()
    sizing = position_sizing_ranges(
        holdings, bridge, max_position=0.75, half_width=0.025,
        liquidity=liquidity, max_days_to_liquidate=5.0,
    )
    assert set(sizing["Ticker"]) == {"AAA", "BBB"}
    assert ((sizing["SuggestedMin"] >= 0) & (sizing["SuggestedMax"] <= 0.75)).all()
    assert (sizing["SuggestedMin"] <= sizing["SuggestedMidpoint"]).all()
    assert (sizing["SuggestedMidpoint"] <= sizing["SuggestedMax"]).all()

    budget = thesis_budget(holdings, bridge)
    assert budget["ExpectedAlphaBudget"].sum() == pytest.approx(1.0)
    assert budget.loc[budget["Ticker"] == "AAA", "ExpectedAlphaBudget"].iloc[0] >            budget.loc[budget["Ticker"] == "BBB", "ExpectedAlphaBudget"].iloc[0]

    detail, summary = expected_return_decomposition(holdings, bridge)
    aaa = detail[detail["Ticker"] == "AAA"].iloc[0]
    assert aaa["ValuationNormalizationResidual"] == pytest.approx(0.03)
    total = summary.loc[
        summary["Component"] == "Total confidence-adjusted expected return",
        "PortfolioContribution",
    ].iloc[0]
    assert total == pytest.approx(0.60 * 0.14 + 0.40 * 0.09)


def test_custom_fundamental_scenarios_and_rebalance_gate(tmp_path):
    holdings, bridge, liquidity = _portfolio_inputs()
    scenario_path = tmp_path / "fundamental_scenarios.csv"
    scenario_path.write_text(
        "Scenario,Ticker,Shock,Notes\n"
        "Recession,AAA,-0.20,test\n"
        "Recession,BBB,-0.10,test\n",
        encoding="utf-8",
    )
    out = custom_thesis_scenarios(scenario_path, holdings)
    portfolio = out[out["Ticker"].isna()].iloc[0]
    assert portfolio["PortfolioShock"] == pytest.approx(-0.16)
    assert portfolio["CoveredWeight"] == pytest.approx(1.0)

    optimizer = pd.DataFrame([
        {"Portfolio": "Regime-Aware ML Maximum Sharpe", "Ticker": "AAA", "CurrentWeight": 0.60, "TargetWeight": 0.68, "WeightChange": 0.08},
        {"Portfolio": "Regime-Aware ML Maximum Sharpe", "Ticker": "BBB", "CurrentWeight": 0.40, "TargetWeight": 0.32, "WeightChange": -0.08},
    ])
    gate = transaction_cost_rebalance(
        optimizer, holdings, liquidity, bridge,
        spread_bps=5.0, impact_bps_at_10pct_adv=20.0,
        rebalance_threshold=0.03, min_benefit_cost_ratio=1.0,
    )
    assert set(gate["Ticker"]) == {"AAA", "BBB"}
    assert set(gate["Decision"]).issubset({"REBALANCE", "HOLD / NO TRADE"})
    assert (gate["EstimatedCostBps"] > 0).all()
