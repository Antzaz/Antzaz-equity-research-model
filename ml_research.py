from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import json
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
from machine_learning.history_store import HistoryStore, utc_now
from machine_learning.quality import gate_results
from machine_learning.workbook import write_ml_sheet

BASE=Path(__file__).resolve().parent
ML_RUNS=BASE/"ml_runs"
HISTORY_DB=BASE/"ml_data"/"ml_history.sqlite"
VALUATION_FEATURES=["price_to_sales","earnings_yield","fcf_yield","book_to_market","ev_to_sales"]


def ticker_type(raw:str)->str:
    t=raw.upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}",t): raise argparse.ArgumentTypeError("Enter a ticker such as GOOGL, TSM, MSFT, or SIE.DE")
    return t


def parser():
    p=argparse.ArgumentParser(description="Run six-model machine-learning research layer")
    p.add_argument("ticker",type=ticker_type)
    p.add_argument("--workbook",help="Optional workbook path; defaults to newest generated model")
    p.add_argument("--benchmark",default="SPY")
    p.add_argument("--max-universe",type=int,default=25,help="Maximum symbols used by the legacy live-data fallback when no persistent history is ready")
    p.add_argument("--max-position",type=float,default=.25)
    p.add_argument("--risk-aversion",type=float,default=4.0)
    p.add_argument("--history-db",default=str(HISTORY_DB),help="Persistent SQLite training database")
    p.add_argument("--no-history-db",action="store_true",help="Ignore the persistent database and use legacy live-data construction")
    p.add_argument("--no-workbook-write",action="store_true")
    return p


def _portfolio_returns_live(tickers:list[str])->pd.DataFrame:
    if len(tickers)<2: return pd.DataFrame()
    try:
        raw=yf.download(tickers,period="3y",auto_adjust=True,progress=False,group_by="column")
        close=raw["Close"] if isinstance(raw.columns,pd.MultiIndex) else raw
        if isinstance(close,pd.Series): close=close.to_frame(tickers[0])
        return close.pct_change().dropna(how="all")
    except Exception:
        return pd.DataFrame()


def _portfolio_returns_db(store:HistoryStore,tickers:list[str])->pd.DataFrame:
    series=[]
    for ticker in tickers:
        pf=store.price_frame(ticker)
        if pf.empty: continue
        px=pf["adj_close"].fillna(pf["close"]).dropna().tail(756)
        if len(px)>=126: series.append(px.rename(ticker))
    if len(series)<2: return pd.DataFrame()
    return pd.concat(series,axis=1).pct_change().dropna(how="all")


def _latest_db_valuation_features(store:HistoryStore,ticker:str)->dict[str,float|None]:
    try:
        with store.connect() as con:
            row=con.execute(
                "SELECT revenue,net_income,fcf,cash,debt,equity,shares FROM fundamentals WHERE symbol=? AND period='annual' "
                "ORDER BY available_date DESC, CASE source WHEN 'Alpha Vantage' THEN 0 WHEN 'SEC Company Facts' THEN 1 ELSE 2 END LIMIT 1",
                (ticker.upper(),),
            ).fetchone()
        pf=store.price_frame(ticker)
        if row is None or pf.empty: return {}
        rev,ni,fcf,cash,debt,equity,shares=row; px=float(pf["adj_close"].fillna(pf["close"]).dropna().iloc[-1])
        mc=px*float(shares) if shares not in (None,0) else None
        ev=(mc+(debt or 0)-(cash or 0)) if mc else None
        return {
            "price_to_sales":mc/rev if mc and rev else None,
            "earnings_yield":ni/mc if mc and ni is not None else None,
            "fcf_yield":fcf/mc if mc and fcf is not None else None,
            "book_to_market":equity/mc if mc and equity is not None else None,
            "ev_to_sales":ev/rev if ev is not None and rev else None,
        }
    except Exception:
        return {}


def _journal(store:HistoryStore|None,ticker:str,stamp:str,results):
    if store is None: return
    rows=[]
    for r in results:
        pred=r.prediction
        if pred is None: continue
        rows.append((stamp,ticker,r.name,utc_now(),365 if "12M" in r.name else None,json.dumps(pred,default=str),r.confidence,
                     json.dumps({**(r.details or {}),"status":r.status},default=str),"ml-layer-v3-quality-gated",utc_now()))
    if not rows: return
    with store.connect() as con:
        con.executemany(
            "INSERT INTO predictions(run_id,symbol,model,as_of,horizon_days,prediction,confidence,features_json,model_version,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
            rows,
        )


def main()->int:
    args=parser().parse_args(); ticker=args.ticker
    workbook=Path(args.workbook).resolve() if args.workbook else latest_workbook(BASE,ticker)
    peers=peer_tickers_from_workbook(workbook,ticker)
    universe=list(dict.fromkeys(peers+DEFAULT_UNIVERSE))[:max(8,args.max_universe)]

    store=None
    db_path=Path(args.history_db)
    if not args.no_history_db and db_path.exists():
        try:
            store=HistoryStore(db_path); print(f"[ml] persistent history: {db_path}")
        except Exception as exc:
            print(f"[ml] persistent history unavailable ({exc}); falling back to live construction")

    current=workbook_current_snapshot(workbook,ticker); current.update(current_market_features(ticker))
    expected_frame=pd.DataFrame(); expected_source="live fallback"
    if store is not None:
        expected_frame=store.expected_return_frame(min_rows=30)
        if not expected_frame.empty:
            current.update(_latest_db_valuation_features(store,ticker)); expected_source="persistent point-in-time database"
    if expected_frame.empty:
        print(f"[ml] training universe: {len(universe)} symbols (legacy live fallback)")
        expected_frame=build_expected_return_dataset(universe,args.benchmark)
    else:
        print(f"[ml] expected-return training rows from database: {len(expected_frame):,}")
    expected_model=ExpectedReturnModel()
    if expected_source.startswith("persistent"):
        expected_model.feature_cols=list(dict.fromkeys(expected_model.feature_cols+VALUATION_FEATURES))
    expected=expected_model.fit_predict(expected_frame,current)
    expected.details=dict(expected.details or {}); expected.details["training_source"]=expected_source
    expected.details["feature_columns"]=expected_model.feature_cols

    earnings_frame=pd.DataFrame(); earnings_source="live fallback"
    if store is not None:
        earnings_frame=store.earnings_model_frame(ticker)
        if len(earnings_frame)>=8: earnings_source="persistent earnings history"
        else: earnings_frame=pd.DataFrame()
    if earnings_frame.empty: earnings_frame=earnings_model_frame(ticker)
    earnings=EarningsSurpriseModel().fit_predict(earnings_frame)
    earnings.details=dict(earnings.details or {}); earnings.details["training_source"]=earnings_source

    anomaly=FinancialAnomalyModel().fit_predict(historical_financial_anomaly_frame(workbook))
    regime=MarketRegimeModel().fit_predict(regime_feature_frame(period="20y"))
    ai=AIImpactMLModel().fit_predict(load_ai_kpi_snapshots(BASE,ticker))

    portfolio_df=load_portfolio(BASE)
    if portfolio_df.empty:
        portfolio=PortfolioPositionSizingModel().optimize(pd.DataFrame(),None)
        portfolio.summary="No local institutional_research/portfolio.csv was found, so portfolio sizing is ready but intentionally not run."
    else:
        tickers=portfolio_df["ticker"].tolist()
        returns=_portfolio_returns_db(store,tickers) if store is not None else pd.DataFrame()
        if returns.empty: returns=_portfolio_returns_live(tickers)
        # Weak expected-return evidence is intentionally excluded from portfolio optimization.
        exp_inputs={ticker:expected.prediction} if expected.status=="PASS" and isinstance(expected.prediction,float) else {}
        weights=dict(zip(portfolio_df["ticker"],portfolio_df["current_weight"].fillna(0.0)))
        portfolio=PortfolioPositionSizingModel().optimize(returns,exp_inputs,weights,max_weight=args.max_position,risk_aversion=args.risk_aversion)

    results=gate_results([expected,earnings,anomaly,regime,ai,portfolio])
    for r in results: print(f"[ml] {r.name}: {r.status} / {r.confidence}")

    stamp=datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    run_dir=ML_RUNS/ticker/stamp; run_dir.mkdir(parents=True,exist_ok=True)
    payload={"ticker":ticker,"generated_at":datetime.now().astimezone().isoformat(timespec="seconds"),"workbook":str(workbook),
             "persistent_history_db":str(db_path) if store else None,"models":[r.to_dict() for r in results]}
    write_json(run_dir/"ml_results.json",payload)
    _journal(store,ticker,stamp,results)
    if not args.no_workbook_write: write_ml_sheet(workbook,ticker,results)
    print(f"[done] ML results: {run_dir/'ml_results.json'}")
    print(f"[done] Workbook updated: {workbook}" if not args.no_workbook_write else "[done] Workbook write skipped")
    return 0

if __name__=="__main__": raise SystemExit(main())
