from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import re

import pandas as pd
import yfinance as yf

from machine_learning import (
    ExpectedReturnModel,EarningsSurpriseModel,FinancialAnomalyModel,MarketRegimeModel,
    AIImpactMLModel,PortfolioPositionSizingModel,
)
from machine_learning.common import write_json
from machine_learning.data import (
    DEFAULT_UNIVERSE, build_expected_return_dataset, current_market_features, earnings_model_frame,
    historical_financial_anomaly_frame, latest_workbook, load_ai_kpi_snapshots, load_portfolio,
    peer_tickers_from_workbook, regime_feature_frame, workbook_current_snapshot,
)
from machine_learning.workbook import write_ml_sheet

BASE=Path(__file__).resolve().parent
ML_RUNS=BASE/"ml_runs"


def ticker_type(raw:str)->str:
    t=raw.upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}",t): raise argparse.ArgumentTypeError("Enter a ticker such as GOOGL, TSM, or MSFT")
    return t


def parser():
    p=argparse.ArgumentParser(description="Run six-model machine-learning research layer")
    p.add_argument("ticker",type=ticker_type)
    p.add_argument("--workbook",help="Optional workbook path; defaults to newest generated model")
    p.add_argument("--benchmark",default="SPY")
    p.add_argument("--max-universe",type=int,default=25,help="Maximum symbols used for expected-return training")
    p.add_argument("--max-position",type=float,default=.25)
    p.add_argument("--risk-aversion",type=float,default=4.0)
    p.add_argument("--no-workbook-write",action="store_true")
    return p


def _portfolio_returns(tickers:list[str])->pd.DataFrame:
    if len(tickers)<2: return pd.DataFrame()
    try:
        raw=yf.download(tickers,period="3y",auto_adjust=True,progress=False,group_by="column")
        close=raw["Close"] if isinstance(raw.columns,pd.MultiIndex) else raw
        if isinstance(close,pd.Series): close=close.to_frame(tickers[0])
        return close.pct_change().dropna(how="all")
    except Exception:
        return pd.DataFrame()


def main()->int:
    args=parser().parse_args(); ticker=args.ticker
    workbook=Path(args.workbook).resolve() if args.workbook else latest_workbook(BASE,ticker)
    peers=peer_tickers_from_workbook(workbook,ticker)
    universe=list(dict.fromkeys(peers+DEFAULT_UNIVERSE))[:max(8,args.max_universe)]
    print(f"[ml] training universe: {len(universe)} symbols")

    current=workbook_current_snapshot(workbook,ticker); current.update(current_market_features(ticker))
    expected_frame=build_expected_return_dataset(universe,args.benchmark)
    expected=ExpectedReturnModel().fit_predict(expected_frame,current)
    print(f"[ml] {expected.name}: {expected.status}")

    earnings=EarningsSurpriseModel().fit_predict(earnings_model_frame(ticker))
    print(f"[ml] {earnings.name}: {earnings.status}")

    anomaly=FinancialAnomalyModel().fit_predict(historical_financial_anomaly_frame(workbook))
    print(f"[ml] {anomaly.name}: {anomaly.status}")

    regime=MarketRegimeModel().fit_predict(regime_feature_frame())
    print(f"[ml] {regime.name}: {regime.status}")

    ai=AIImpactMLModel().fit_predict(load_ai_kpi_snapshots(BASE,ticker))
    print(f"[ml] {ai.name}: {ai.status}")

    portfolio_df=load_portfolio(BASE)
    if portfolio_df.empty:
        portfolio=PortfolioPositionSizingModel().optimize(pd.DataFrame(),None)
        portfolio.summary="No local institutional_research/portfolio.csv was found, so portfolio sizing is ready but intentionally not run."
    else:
        tickers=portfolio_df["ticker"].tolist(); returns=_portfolio_returns(tickers)
        exp_inputs={ticker:expected.prediction} if isinstance(expected.prediction,float) else {}
        weights=dict(zip(portfolio_df["ticker"],portfolio_df["current_weight"].fillna(0.0)))
        portfolio=PortfolioPositionSizingModel().optimize(returns,exp_inputs,weights,max_weight=args.max_position,risk_aversion=args.risk_aversion)
    print(f"[ml] {portfolio.name}: {portfolio.status}")

    results=[expected,earnings,anomaly,regime,ai,portfolio]
    stamp=datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    run_dir=ML_RUNS/ticker/stamp; run_dir.mkdir(parents=True,exist_ok=True)
    payload={"ticker":ticker,"generated_at":datetime.now().astimezone().isoformat(timespec="seconds"),"workbook":str(workbook),"models":[r.to_dict() for r in results]}
    write_json(run_dir/"ml_results.json",payload)
    if not args.no_workbook_write:
        write_ml_sheet(workbook,ticker,results)
    print(f"[done] ML results: {run_dir/'ml_results.json'}")
    print(f"[done] Workbook updated: {workbook}" if not args.no_workbook_write else "[done] Workbook write skipped")
    return 0

if __name__=="__main__": raise SystemExit(main())
