from __future__ import annotations

from pathlib import Path
import sys

import numpy as np
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/"institutional_research"
if str(SRC) not in sys.path: sys.path.insert(0,str(SRC))

from src.decision_loop import (
    build_realized_portfolio, position_sizing_ranges, thesis_budget,
    custom_thesis_scenarios, expected_return_decomposition,
    transaction_cost_rebalance, decision_journal_analytics,
    high_level_decision_learning,
)

def test_realized_portfolio_reconstructs_twr_and_mwr():
    ledger=pd.DataFrame([
        {"Date":"2026-01-02","Ticker":"","Action":"DEPOSIT","Shares":np.nan,"Price":np.nan,"Fees":0,"CashFlow":1000,"Currency":"USD","DecisionID":"","ReferencePrice":np.nan,"SplitRatio":np.nan,"Notes":""},
        {"Date":"2026-01-02","Ticker":"AAA","Action":"BUY","Shares":10,"Price":100,"Fees":0,"CashFlow":np.nan,"Currency":"USD","DecisionID":"D1","ReferencePrice":100,"SplitRatio":np.nan,"Notes":""},
    ])
    ledger["Date"]=pd.to_datetime(ledger["Date"])
    prices=pd.DataFrame({"AAA":[100,105,110]},index=pd.to_datetime(["2026-01-02","2026-01-05","2026-01-06"]))
    bench=pd.Series([100,101,102],index=prices.index)
    out=build_realized_portfolio(ledger,prices,bench,{})
    s=out["summary"]
    assert s["status"]=="PASS"
    assert abs(s["twr_total"]-.10)<1e-9
    assert s["mwr_xirr"] is not None
    assert len(out["weights"])==6
    assert abs(out["timeseries"].iloc[-1]["NAV"]-1100)<1e-9

def _holdings():
    return pd.DataFrame([
        {"Ticker":"AAA","Weight":.60,"RiskContributionPct":.70,"MarketValue":60000},
        {"Ticker":"BBB","Weight":.40,"RiskContributionPct":.30,"MarketValue":40000},
    ])

def _bridge():
    return pd.DataFrame([
        {"Ticker":"AAA","ExpectedAlpha":.08,"Confidence":.80,"ConfidenceAdjustedExpectedReturn":.16,"CurrentPrice":100,"BearValue":75,"ValuationConvergenceReturn":.10,"DividendYield":.01,"NetBuybackYield":.02,"RawExpectedReturn":.13},
        {"Ticker":"BBB","ExpectedAlpha":.03,"Confidence":.60,"ConfidenceAdjustedExpectedReturn":.11,"CurrentPrice":100,"BearValue":85,"ValuationConvergenceReturn":.05,"DividendYield":.02,"NetBuybackYield":.00,"RawExpectedReturn":.07},
    ])

def test_sizing_thesis_budget_and_expected_return_decomposition():
    holdings=_holdings(); bridge=_bridge()
    sizing=position_sizing_ranges(holdings,bridge,max_position=.25)
    assert set(sizing["Ticker"])=={"AAA","BBB"}
    assert (sizing["SuggestedMax"]<=.25+1e-12).all()
    budget=thesis_budget(holdings,bridge)
    assert abs(budget["ExpectedAlphaBudget"].sum()-1)<1e-9
    detail,summary=expected_return_decomposition(holdings,bridge)
    assert len(detail)==2
    total=float(summary.loc[summary["Component"]=="Total confidence-adjusted expected return","PortfolioContribution"].iloc[0])
    assert abs(total-(.6*.16+.4*.11))<1e-9

def test_custom_fundamental_scenario_and_rebalance_cost_gate(tmp_path: Path):
    holdings=_holdings(); bridge=_bridge()
    p=tmp_path/"fundamental_scenarios.csv"
    pd.DataFrame([
        {"Scenario":"AI boom","Ticker":"AAA","Shock":.20,"Notes":"demo"},
        {"Scenario":"AI boom","Ticker":"BBB","Shock":.05,"Notes":"demo"},
    ]).to_csv(p,index=False)
    out=custom_thesis_scenarios(p,holdings)
    top=out[out["Ticker"].isna()].iloc[0]
    assert abs(float(top["PortfolioShock"])-(.6*.20+.4*.05))<1e-9
    optimizer=pd.DataFrame([
        {"Portfolio":"Regime-Aware ML Maximum Sharpe","Ticker":"AAA","CurrentWeight":.60,"TargetWeight":.50,"WeightChange":-.10},
        {"Portfolio":"Regime-Aware ML Maximum Sharpe","Ticker":"BBB","CurrentWeight":.40,"TargetWeight":.50,"WeightChange":.10},
    ])
    liquidity=pd.DataFrame([
        {"Ticker":"AAA","AverageDailyDollarVolume":10_000_000},
        {"Ticker":"BBB","AverageDailyDollarVolume":10_000_000},
    ])
    reb=transaction_cost_rebalance(optimizer,holdings,liquidity,bridge,rebalance_threshold=.03)
    assert len(reb)==2 and (reb["EstimatedCostBps"]>0).all()

def test_decision_journal_outcomes_and_learning(tmp_path: Path):
    dates=pd.date_range("2025-01-02","2026-02-01",freq="B")
    prices=pd.DataFrame({"AAA":np.linspace(100,140,len(dates)),"SPY":np.linspace(100,115,len(dates))},index=dates)
    p=tmp_path/"journal.csv"
    pd.DataFrame([{"DecisionID":"D1","Date":"2025-01-02","Ticker":"AAA","Decision":"BUY","OldWeight":0,"NewWeight":.1,"ExpectedReturn":.10,"Conviction":5,"Sector":"Tech","ThesisCategory":"Quality","Catalyst":"Earnings","PrimaryReason":"test","KeyRisk":"test","WhatWouldChangeMyMind":"test","ReviewDate":"2026-01-02","OutcomeNotes":""}]).to_csv(p,index=False)
    detail,summary=decision_journal_analytics(p,prices,"SPY")
    assert not detail.empty and not summary.empty
    learning=high_level_decision_learning(detail)
    assert not learning.empty and "CorrectRate" in learning.columns
