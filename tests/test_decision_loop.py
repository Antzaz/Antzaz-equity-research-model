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
    assert aaa["ValuationConvergence"] == pytest.approx(0.08)
    assert aaa["DividendYield"] == pytest.approx(0.01)
    assert aaa["NetBuybackYield"] == pytest.approx(0.02)
    assert aaa["ConfidenceShrinkage"] == pytest.approx(0.03)
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


def test_decision_journal_learning_by_sector_and_category(tmp_path):
    from src.decision_loop import decision_journal_analytics, high_level_decision_learning

    dates=pd.date_range("2025-01-02","2026-02-10",freq="B")
    aaa=pd.Series(100.0*(1.0008**np.arange(len(dates))),index=dates)
    spy=pd.Series(100.0*(1.0003**np.arange(len(dates))),index=dates)
    prices=pd.DataFrame({"AAA":aaa,"SPY":spy})
    path=tmp_path/"journal.csv"
    path.write_text(
        "DecisionID,Date,Ticker,Decision,OldWeight,NewWeight,ExpectedReturn,Conviction,Sector,ThesisCategory,Catalyst,PrimaryReason,KeyRisk,WhatWouldChangeMyMind,ReviewDate,OutcomeNotes\n"
        "D1,2025-01-02,AAA,BUY,0.00,0.10,0.12,5,Technology,AI growth,Earnings,thesis,risk,break,2025-07-01,\n",
        encoding="utf-8",
    )
    detail,summary=decision_journal_analytics(path,prices,"SPY",sector_map={"AAA":"Technology"})
    assert not detail.empty
    assert not summary.empty
    matured=detail[detail["Matured"]==True]
    assert matured["Correct"].all()
    learning=high_level_decision_learning(detail)
    assert {"Decision","Conviction","Sector","Thesis Category"}.issubset(set(learning["Dimension"]))
    sector=learning[(learning["Dimension"]=="Sector") & (learning["Group"]=="Technology")]
    assert not sector.empty
    assert sector["AverageDecisionAlpha"].iloc[0]>0


def test_brinson_loader_accepts_commented_template(tmp_path):
    from src.decision_loop import brinson_sector_attribution
    p=tmp_path/"benchmark_sector_history.csv"
    p.write_text(
        "# Copy to benchmark_sector_history.csv only when you have point-in-time benchmark sector data.\n"
        "# Return is the sector return for that date/period.\n"
        "Date,Sector,Weight,Return\n"
        "2026-01-05,Technology,0.60,0.02\n"
        "2026-01-05,Other,0.40,0.01\n",
        encoding="utf-8",
    )
    weights=pd.DataFrame([
        {"Date":pd.Timestamp("2026-01-02"),"Ticker":"AAA","Weight":0.60},
        {"Date":pd.Timestamp("2026-01-02"),"Ticker":"BBB","Weight":0.40},
        {"Date":pd.Timestamp("2026-01-05"),"Ticker":"AAA","Weight":0.60},
        {"Date":pd.Timestamp("2026-01-05"),"Ticker":"BBB","Weight":0.40},
    ])
    prices=pd.DataFrame(
        {"AAA":[100,102],"BBB":[100,101]},
        index=pd.to_datetime(["2026-01-02","2026-01-05"]),
    )
    out=brinson_sector_attribution(weights,prices,{"AAA":"Technology","BBB":"Other"},p)
    assert not out.empty
    assert {"Allocation","Selection","Interaction","TotalActiveContribution"}.issubset(out.columns)


def test_brinson_loader_malformed_optional_csv_does_not_crash(tmp_path):
    from src.decision_loop import brinson_sector_attribution
    p=tmp_path/"benchmark_sector_history.csv"
    p.write_text("# instruction line\nDate,Sector,Weight,Return\nbad,row,with,too,many,columns\n",encoding="utf-8")
    out=brinson_sector_attribution(pd.DataFrame(),pd.DataFrame(),{},p)
    assert out.empty
