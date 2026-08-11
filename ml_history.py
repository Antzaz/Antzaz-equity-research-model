from __future__ import annotations

"""Build and maintain the persistent historical training database.

Examples:
    python ml_history.py status
    python ml_history.py bootstrap --universe sp500 --limit 500 --years 20
    python ml_history.py enrich-alpha --universe sp500 --limit 500 --call-budget 20
    python ml_history.py enrich-fmp --universe sp500 --limit 500 --call-budget 200
    python ml_history.py build-features

The database is local under ml_data/ml_history.sqlite and is intentionally gitignored.
All network steps are resumable/idempotent through SQLite upserts and provider state.
"""

import argparse
from pathlib import Path
import os
import shutil

from machine_learning.data import DEFAULT_UNIVERSE
from machine_learning.history_store import HistoryStore
from machine_learning.history_sources import (
    AlphaVantageClient,FMPClient,current_sp500_universe,fred_macro_rows,sec_cik_map,
    sec_fundamental_rows,yahoo_fundamental_rows,yahoo_price_rows,
)

BASE=Path(__file__).resolve().parent
DEFAULT_DB=BASE/"ml_data"/"ml_history.sqlite"


def parser():
    p=argparse.ArgumentParser(description="Persistent point-in-time ML history database")
    p.add_argument("--db",default=str(DEFAULT_DB),help="SQLite database path")
    sub=p.add_subparsers(dest="command",required=True)
    sub.add_parser("status")
    for name in ("bootstrap","backfill-prices","backfill-fundamentals","enrich-alpha","enrich-fmp"):
        s=sub.add_parser(name)
        s.add_argument("--universe",choices=["sp500","default"],default="sp500")
        s.add_argument("--limit",type=int,default=500)
        if name in {"bootstrap","backfill-prices"}: s.add_argument("--years",type=int,default=20)
        if name in {"bootstrap","enrich-alpha"}: s.add_argument("--call-budget",type=int,default=20,help="Maximum Alpha Vantage calls this run; free keys are currently limited to 25/day")
        if name=="enrich-fmp": s.add_argument("--call-budget",type=int,default=200)
        if name=="bootstrap":
            s.add_argument("--skip-sec",action="store_true")
            s.add_argument("--skip-alpha",action="store_true")
            s.add_argument("--skip-macro",action="store_true")
    sub.add_parser("backfill-macro")
    f=sub.add_parser("build-features"); f.add_argument("--benchmark",default="SPY")
    return p


def _disk_preflight(path:Path,min_gb=2.0):
    path.parent.mkdir(parents=True,exist_ok=True)
    free=shutil.disk_usage(path.parent).free/(1024**3)
    if free<min_gb:
        raise OSError(f"Only {free:.2f} GB free. Free at least {min_gb:.1f} GB before a large historical backfill.")
    print(f"[history] disk free: {free:.2f} GB")


def _universe(store,kind,limit):
    if kind=="sp500": rows=current_sp500_universe(limit)
    else: rows=[{"symbol":s,"universe_source":"Built-in research universe"} for s in DEFAULT_UNIVERSE[:limit]]
    # Always include the benchmark because targets are benchmark-relative.
    if not any(r.get("symbol")=="SPY" for r in rows): rows.append({"symbol":"SPY","name":"SPDR S&P 500 ETF Trust","universe_source":"ML benchmark"})
    store.upsert_symbols(rows)
    return rows,[r["symbol"] for r in rows]


def _price_complete(store,symbol,years):
    minimum=max(250,min(int(years)*180,2500))
    with store.connect() as con:
        n=con.execute("SELECT COUNT(DISTINCT date) FROM prices WHERE symbol=? AND source='Yahoo Finance'",(symbol,)).fetchone()[0]
    return n>=minimum


def backfill_prices(store,symbols,years):
    needed=[s for s in symbols if not _price_complete(store,s,years)]
    print(f"[prices] {len(needed)} symbol(s) need Yahoo history; {len(symbols)-len(needed)} already sufficient")
    buf=[]; total=0
    for row in yahoo_price_rows(needed,years=years):
        buf.append(row)
        if len(buf)>=10000:
            total+=store.upsert_prices(buf); buf=[]; print(f"[prices] upserted {total:,} rows")
    if buf: total+=store.upsert_prices(buf)
    print(f"[prices] complete: {total:,} row upserts")
    return total


def _fundamental_count(store,symbol):
    with store.connect() as con:
        return con.execute("SELECT COUNT(DISTINCT fiscal_year) FROM fundamentals WHERE symbol=? AND period='annual'",(symbol,)).fetchone()[0]


def backfill_fundamentals(store,rows,use_sec=True):
    ua=(os.getenv("SEC_USER_AGENT") or "").strip(); cmap={}
    if use_sec and ua:
        try:
            cmap=sec_cik_map(ua); print(f"[fundamentals] SEC ticker map loaded: {len(cmap):,} symbols")
        except Exception as exc: print(f"[fundamentals] SEC map unavailable: {exc}")
    elif use_sec:
        print("[fundamentals] SEC_USER_AGENT is not configured; using Yahoo fallback. Set a real contact User-Agent for deeper free SEC history.")
    inserted=0
    for i,r in enumerate(rows,1):
        symbol=r["symbol"]
        if symbol=="SPY" or _fundamental_count(store,symbol)>=8: continue
        frows=[]
        cik=str(r.get("cik") or cmap.get(symbol) or "").zfill(10) if (r.get("cik") or cmap.get(symbol)) else None
        if use_sec and ua and cik:
            frows=sec_fundamental_rows(symbol,cik,ua)
        if not frows: frows=yahoo_fundamental_rows(symbol)
        inserted+=store.upsert_fundamentals(frows)
        if i%25==0: print(f"[fundamentals] processed {i}/{len(rows)}; row upserts={inserted:,}")
    print(f"[fundamentals] complete: {inserted:,} row upserts")
    return inserted


def enrich_alpha(store,symbols,call_budget=20):
    client=AlphaVantageClient(max_calls=call_budget)
    if not client.key:
        print("[alpha] ALPHAVANTAGE_API_KEY is not configured; skipped")
        return 0
    enriched=0
    for symbol in symbols:
        if symbol=="SPY" or client.max_calls-client.calls<4: break
        if store.get_state("Alpha Vantage",f"enriched:{symbol}"): continue
        fundamentals,earnings=client.enrich_symbol(symbol)
        if fundamentals or earnings:
            store.upsert_fundamentals(fundamentals); store.upsert_earnings(earnings)
            store.set_state("Alpha Vantage",f"enriched:{symbol}",{"calls_at_completion":client.calls})
            enriched+=1
            print(f"[alpha] {symbol}: fundamentals={len(fundamentals)}, earnings={len(earnings)}, calls={client.calls}/{client.max_calls}")
        else:
            print(f"[alpha] {symbol}: no usable response; leaving pending for a later retry")
    print(f"[alpha] enriched {enriched} symbol(s), used {client.calls} call(s)")
    return enriched


def enrich_fmp(store,symbols,call_budget=200):
    client=FMPClient(max_calls=call_budget)
    if not client.key:
        print("[fmp] FMP_API_KEY is not configured; adapter is ready but skipped")
        return 0
    enriched=0
    for symbol in symbols:
        if symbol=="SPY" or client.max_calls-client.calls<2: break
        consensus,earnings=client.enrich_symbol(symbol)
        if consensus: store.upsert_consensus(consensus)
        if earnings: store.upsert_earnings(earnings)
        if consensus or earnings:
            enriched+=1; print(f"[fmp] {symbol}: consensus={len(consensus)}, earnings={len(earnings)}, calls={client.calls}/{client.max_calls}")
    print(f"[fmp] enriched {enriched} symbol(s), used {client.calls} call(s)")
    return enriched


def backfill_macro(store):
    rows=[]; total=0
    for row in fred_macro_rows():
        rows.append(row)
        if len(rows)>=5000: total+=store.upsert_macro(rows); rows=[]
    if rows: total+=store.upsert_macro(rows)
    print(f"[macro] FRED row upserts: {total:,}")
    return total


def print_status(store):
    counts=store.table_counts(); size=store.path.stat().st_size/(1024**2) if store.path.exists() else 0
    print(f"Database: {store.path}")
    print(f"Size: {size:.1f} MB")
    for k,v in counts.items(): print(f"  {k:20s} {v:,}")
    print(f"ALPHAVANTAGE_API_KEY: {'configured' if os.getenv('ALPHAVANTAGE_API_KEY') else 'not configured'}")
    print(f"FMP_API_KEY: {'configured' if os.getenv('FMP_API_KEY') else 'not configured'}")
    print(f"SEC_USER_AGENT: {'configured' if os.getenv('SEC_USER_AGENT') else 'not configured'}")


def main():
    args=parser().parse_args(); store=HistoryStore(Path(args.db))
    if args.command=="status": print_status(store); return 0
    _disk_preflight(store.path)
    if args.command=="build-features":
        n=store.build_features(args.benchmark); print(f"[features] materialized {n:,} point-in-time rows"); print_status(store); return 0
    if args.command=="backfill-macro": backfill_macro(store); print_status(store); return 0
    rows,symbols=_universe(store,args.universe,args.limit)
    if args.command=="backfill-prices": backfill_prices(store,symbols,args.years)
    elif args.command=="backfill-fundamentals": backfill_fundamentals(store,rows,True)
    elif args.command=="enrich-alpha": enrich_alpha(store,symbols,args.call_budget)
    elif args.command=="enrich-fmp": enrich_fmp(store,symbols,args.call_budget)
    elif args.command=="bootstrap":
        backfill_prices(store,symbols,args.years)
        backfill_fundamentals(store,rows,not args.skip_sec)
        if not args.skip_alpha: enrich_alpha(store,symbols,args.call_budget)
        if not args.skip_macro: backfill_macro(store)
        n=store.build_features("SPY"); print(f"[features] materialized {n:,} point-in-time rows")
    print_status(store); return 0


if __name__=="__main__": raise SystemExit(main())
