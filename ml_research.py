from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
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
from machine_learning.history_sources import current_sp500_universe, yahoo_fundamental_rows, yahoo_price_rows
from machine_learning.history_store import HistoryStore, utc_now
from machine_learning.quality import gate_results
from machine_learning.workbook import write_ml_sheet
from runtime_data_guards import provider_symbol

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
    p.add_argument("--max-universe",type=int,default=25,help="Maximum symbols used by the legacy live-data fallback when persistent history is not ready")
    p.add_argument("--max-position",type=float,default=.25)
    p.add_argument("--risk-aversion",type=float,default=4.0)
    p.add_argument("--history-db",default=str(HISTORY_DB),help="Persistent SQLite training database")
    p.add_argument("--no-history-db",action="store_true",help="Ignore the persistent database and use legacy live-data construction")
    p.add_argument("--no-auto-history",action="store_true",help="Do not automatically seed a sparse/missing persistent history database")
    p.add_argument("--auto-history-limit",type=int,default=80,help="Maximum names in the bounded one-time automatic history seed")
    p.add_argument("--auto-history-years",type=int,default=12,help="Price-history years used by the bounded automatic history seed")
    p.add_argument("--no-workbook-write",action="store_true")
    return p


def _provider_ticker(raw:str)->str:
    return provider_symbol(str(raw or "").upper().strip())


def _portfolio_returns_live(tickers:list[str])->pd.DataFrame:
    if len(tickers)<2: return pd.DataFrame()
    provider=[_provider_ticker(t) for t in tickers]
    reverse=dict(zip(provider,tickers))
    try:
        raw=yf.download(provider,period="3y",auto_adjust=True,progress=False,group_by="column")
        close=raw["Close"] if isinstance(raw.columns,pd.MultiIndex) else raw
        if isinstance(close,pd.Series): close=close.to_frame(provider[0])
        close=close.rename(columns=reverse)
        return close.pct_change().dropna(how="all")
    except Exception:
        return pd.DataFrame()


def _portfolio_returns_db(store:HistoryStore,tickers:list[str])->pd.DataFrame:
    series=[]
    for ticker in tickers:
        pf=store.price_frame(_provider_ticker(ticker))
        if pf.empty: continue
        px=pf["adj_close"].fillna(pf["close"]).dropna().tail(756)
        if len(px)>=126: series.append(px.rename(ticker))
    if len(series)<2: return pd.DataFrame()
    return pd.concat(series,axis=1).pct_change().dropna(how="all")


def _latest_db_valuation_features(store:HistoryStore,ticker:str)->dict[str,float|None]:
    symbol=_provider_ticker(ticker)
    try:
        with store.connect() as con:
            row=con.execute(
                "SELECT revenue,net_income,fcf,cash,debt,equity,shares FROM fundamentals WHERE symbol=? AND period='annual' "
                "ORDER BY available_date DESC, CASE source WHEN 'Alpha Vantage' THEN 0 WHEN 'SEC Company Facts' THEN 1 ELSE 2 END LIMIT 1",
                (symbol,),
            ).fetchone()
        pf=store.price_frame(symbol)
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


def _db_count(store:HistoryStore,table:str)->int:
    if table not in {"symbols","prices","fundamentals","features"}: return 0
    with store.connect() as con:
        return int(con.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])


def _symbol_count(store:HistoryStore,table:str,symbol:str)->int:
    symbol=_provider_ticker(symbol)
    with store.connect() as con:
        if table=="prices":
            return int(con.execute("SELECT COUNT(DISTINCT date) FROM prices WHERE symbol=?",(symbol,)).fetchone()[0])
        if table=="fundamentals":
            return int(con.execute("SELECT COUNT(DISTINCT fiscal_year) FROM fundamentals WHERE symbol=? AND period='annual'",(symbol,)).fetchone()[0])
    return 0


def _starter_universe(ticker:str,peers:list[str],benchmark:str,limit:int)->list[dict]:
    requested=[]
    for raw in [ticker,*peers,benchmark]:
        symbol=_provider_ticker(raw)
        if symbol and symbol not in requested: requested.append(symbol)
    broad=current_sp500_universe(max(20,int(limit)))
    rows=[]; seen=set()
    for symbol in requested:
        rows.append({"symbol":symbol,"universe_source":"Automatic ML research seed"}); seen.add(symbol)
    for row in broad:
        symbol=_provider_ticker(row.get("symbol"))
        if not symbol or symbol in seen: continue
        item=dict(row); item["symbol"]=symbol; rows.append(item); seen.add(symbol)
        if len(rows)>=max(limit,len(requested)): break
    for symbol in DEFAULT_UNIVERSE:
        symbol=_provider_ticker(symbol)
        if symbol and symbol not in seen and len(rows)<max(limit,len(requested)):
            rows.append({"symbol":symbol,"universe_source":"Built-in ML fallback universe"}); seen.add(symbol)
    b=_provider_ticker(benchmark)
    if b and b not in seen: rows.append({"symbol":b,"universe_source":"ML benchmark"})
    return rows


def _auto_seed_history(store:HistoryStore,ticker:str,peers:list[str],benchmark:str,limit:int=80,years:int=12)->dict:
    """Bounded, resumable seed so --ml uses persistent point-in-time data by default.

    The full ml_history.py workflow remains the preferred path for a deep 500-name / 20-year
    research database. This helper only prevents a normal --ml run from silently reverting to a
    tiny ephemeral dataset when the local SQLite store has not yet been prepared.
    """
    status={
        "attempted":False,"ready":False,"symbols":_db_count(store,"symbols"),
        "prices":_db_count(store,"prices"),"fundamentals":_db_count(store,"fundamentals"),
        "features":_db_count(store,"features"),"added_prices":0,"added_fundamentals":0,"built_features":0,
    }
    target_ready=_symbol_count(store,"prices",ticker)>=126 and _symbol_count(store,"fundamentals",ticker)>=2
    if status["features"]>=150 and target_ready:
        status["ready"]=True; status["note"]="Persistent history already meets the automatic readiness threshold."
        return status

    status["attempted"]=True
    rows=_starter_universe(ticker,peers,benchmark,max(20,min(int(limit),150)))
    store.upsert_symbols(rows); symbols=[r["symbol"] for r in rows]

    needed=[s for s in symbols if _symbol_count(store,"prices",s)<126]
    buf=[]
    for row in yahoo_price_rows(needed,years=max(3,min(int(years),20)),batch_size=30):
        buf.append(row)
        if len(buf)>=10000:
            status["added_prices"]+=store.upsert_prices(buf); buf=[]
    if buf: status["added_prices"]+=store.upsert_prices(buf)

    benchmark_symbol=_provider_ticker(benchmark)
    needed_f=[s for s in symbols if s!=benchmark_symbol and _symbol_count(store,"fundamentals",s)<2]
    if needed_f:
        with ThreadPoolExecutor(max_workers=min(8,max(1,len(needed_f)))) as pool:
            futures={pool.submit(yahoo_fundamental_rows,s):s for s in needed_f}
            for future in as_completed(futures):
                try: rows_out=future.result() or []
                except Exception: rows_out=[]
                if rows_out: status["added_fundamentals"]+=store.upsert_fundamentals(rows_out)

    try: status["built_features"]=store.build_features(benchmark_symbol)
    except Exception as exc: status["feature_error"]=repr(exc)
    status.update({
        "symbols":_db_count(store,"symbols"),"prices":_db_count(store,"prices"),
        "fundamentals":_db_count(store,"fundamentals"),"features":_db_count(store,"features"),
    })
    status["ready"]=status["features"]>=30
    status["note"]=(
        f"Automatic persistent history {'ready' if status['ready'] else 'sparse'}: {status['features']} realized "
        f"point-in-time feature rows across {status['symbols']} tracked symbols. For deeper training use "
        "`python ml_history.py bootstrap --universe sp500 --limit 500 --years 20` and optional provider enrichments."
    )
    return status


def _journal(store:HistoryStore|None,ticker:str,stamp:str,results):
    if store is None: return
    rows=[]
    for r in results:
        pred=r.prediction
        if pred is None: continue
        rows.append((stamp,ticker,r.name,utc_now(),365 if "12M" in r.name else None,json.dumps(pred,default=str),r.confidence,
                     json.dumps({**(r.details or {}),"status":r.status},default=str),"ml-layer-v4-purged-pit",utc_now()))
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

    store=None; auto_history={}
    db_path=Path(args.history_db)
    if not args.no_history_db:
        try:
            store=HistoryStore(db_path); print(f"[ml] persistent history: {db_path}")
            if not args.no_auto_history:
                auto_history=_auto_seed_history(store,ticker,peers,args.benchmark,args.auto_history_limit,args.auto_history_years)
                print(f"[ml] history readiness: {auto_history.get('note','')}")
        except Exception as exc:
            store=None; auto_history={"attempted":True,"ready":False,"error":repr(exc)}
            print(f"[ml] persistent history unavailable ({exc}); falling back to live construction")

    current=workbook_current_snapshot(workbook,ticker); current.update(current_market_features(_provider_ticker(ticker)))
    expected_frame=pd.DataFrame(); expected_source="live fallback"
    if store is not None:
        expected_frame=store.expected_return_frame(min_rows=30)
        if not expected_frame.empty:
            current.update(_latest_db_valuation_features(store,ticker)); expected_source="persistent point-in-time database"
    if expected_frame.empty:
        print(f"[ml] training universe: {len(universe)} symbols (legacy live fallback)")
        expected_frame=build_expected_return_dataset([_provider_ticker(x) for x in universe],_provider_ticker(args.benchmark))
    else:
        print(f"[ml] expected-return training rows from database: {len(expected_frame):,}")
    expected_model=ExpectedReturnModel()
    if expected_source.startswith("persistent"):
        expected_model.feature_cols=list(dict.fromkeys(expected_model.feature_cols+VALUATION_FEATURES))
    expected=expected_model.fit_predict(expected_frame,current)
    expected.details=dict(expected.details or {}); expected.details["training_source"]=expected_source
    expected.details["feature_columns"]=expected_model.feature_cols
    if auto_history: expected.details["automatic_history"]=auto_history

    earnings_frame=pd.DataFrame(); earnings_source="live fallback"
    if store is not None:
        earnings_frame=store.earnings_model_frame(_provider_ticker(ticker))
        if len(earnings_frame)>=8: earnings_source="persistent earnings history"
        else: earnings_frame=pd.DataFrame()
    if earnings_frame.empty: earnings_frame=earnings_model_frame(_provider_ticker(ticker))
    earnings=EarningsSurpriseModel().fit_predict(earnings_frame)
    earnings.details=dict(earnings.details or {}); earnings.details["training_source"]=earnings_source

    anomaly=FinancialAnomalyModel().fit_predict(historical_financial_anomaly_frame(workbook))
    regime=MarketRegimeModel().fit_predict(regime_feature_frame(period="20y"))
    ai=AIImpactMLModel().fit_predict(load_ai_kpi_snapshots(BASE,ticker))

    # Apply validation gates before any downstream model consumes another model's prediction.
    expected,earnings,anomaly,regime,ai=gate_results([expected,earnings,anomaly,regime,ai])

    portfolio_df=load_portfolio(BASE)
    if portfolio_df.empty:
        portfolio=PortfolioPositionSizingModel().optimize(pd.DataFrame(),None)
        portfolio.summary="No local institutional_research/portfolio.csv was found, so portfolio sizing is ready but intentionally not run. Add actual tickers/current weights rather than inventing a portfolio."
    else:
        tickers=portfolio_df["ticker"].tolist()
        returns=_portfolio_returns_db(store,tickers) if store is not None else pd.DataFrame()
        if returns.empty: returns=_portfolio_returns_live(tickers)
        exp_inputs={ticker:expected.prediction} if expected.status=="PASS" and isinstance(expected.prediction,float) else {}
        weights=dict(zip(portfolio_df["ticker"],portfolio_df["current_weight"].fillna(0.0)))
        portfolio=PortfolioPositionSizingModel().optimize(returns,exp_inputs,weights,max_weight=args.max_position,risk_aversion=args.risk_aversion)
    portfolio=gate_results([portfolio])[0]

    results=[expected,earnings,anomaly,regime,ai,portfolio]
    for r in results: print(f"[ml] {r.name}: {r.status} / {r.confidence}")

    stamp=datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    run_dir=ML_RUNS/ticker/stamp; run_dir.mkdir(parents=True,exist_ok=True)
    payload={"ticker":ticker,"generated_at":datetime.now().astimezone().isoformat(timespec="seconds"),"workbook":str(workbook),
             "persistent_history_db":str(db_path) if store else None,"automatic_history":auto_history,"models":[r.to_dict() for r in results]}
    write_json(run_dir/"ml_results.json",payload)
    _journal(store,ticker,stamp,results)
    if not args.no_workbook_write: write_ml_sheet(workbook,ticker,results)
    print(f"[done] ML results: {run_dir/'ml_results.json'}")
    print(f"[done] Workbook updated: {workbook}" if not args.no_workbook_write else "[done] Workbook write skipped")
    return 0

if __name__=="__main__": raise SystemExit(main())
