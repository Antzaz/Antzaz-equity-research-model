from __future__ import annotations

"""Persistent point-in-time history store for the machine-learning research layer.

SQLite is used deliberately: it is built into Python, transactional, resumable, compact,
and easy to query without introducing another database service. The file lives under
ml_data/ and is gitignored; it is local research data, not source code.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import json
import math
import sqlite3

import numpy as np
import pandas as pd


SCHEMA_VERSION = 1


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _finite(v):
    try:
        x=float(v)
        return x if math.isfinite(x) else None
    except Exception:
        return None


@dataclass
class HistoryStore:
    path: Path

    def __post_init__(self):
        self.path=Path(self.path)
        self.path.parent.mkdir(parents=True,exist_ok=True)
        self.init_schema()

    def connect(self):
        con=sqlite3.connect(self.path,timeout=60)
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA synchronous=NORMAL")
        con.execute("PRAGMA foreign_keys=ON")
        return con

    def init_schema(self):
        ddl="""
        CREATE TABLE IF NOT EXISTS metadata(
            key TEXT PRIMARY KEY,
            value TEXT,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS symbols(
            symbol TEXT PRIMARY KEY,
            name TEXT, sector TEXT, industry TEXT, exchange TEXT, cik TEXT, currency TEXT,
            universe_source TEXT, first_seen TEXT, last_seen TEXT
        );
        CREATE TABLE IF NOT EXISTS prices(
            symbol TEXT NOT NULL, date TEXT NOT NULL, close REAL, adj_close REAL, volume REAL,
            source TEXT NOT NULL, fetched_at TEXT NOT NULL,
            PRIMARY KEY(symbol,date,source)
        );
        CREATE INDEX IF NOT EXISTS idx_prices_symbol_date ON prices(symbol,date);
        CREATE TABLE IF NOT EXISTS fundamentals(
            symbol TEXT NOT NULL, fiscal_year INTEGER NOT NULL, fiscal_date TEXT,
            available_date TEXT NOT NULL, period TEXT NOT NULL DEFAULT 'annual',
            revenue REAL, operating_income REAL, net_income REAL, eps REAL, shares REAL,
            ocf REAL, capex REAL, fcf REAL, cash REAL, debt REAL, equity REAL, assets REAL,
            liabilities REAL, rd REAL, sbc REAL, currency TEXT, source TEXT NOT NULL,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY(symbol,fiscal_year,period,source)
        );
        CREATE INDEX IF NOT EXISTS idx_fundamentals_symbol_available ON fundamentals(symbol,available_date);
        CREATE TABLE IF NOT EXISTS earnings(
            symbol TEXT NOT NULL, reported_date TEXT NOT NULL, fiscal_date TEXT,
            eps_actual REAL, eps_estimate REAL, surprise_pct REAL,
            revenue_actual REAL, revenue_estimate REAL, source TEXT NOT NULL,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY(symbol,reported_date,source)
        );
        CREATE INDEX IF NOT EXISTS idx_earnings_symbol_date ON earnings(symbol,reported_date);
        CREATE TABLE IF NOT EXISTS macro(
            series TEXT NOT NULL, date TEXT NOT NULL, value REAL, source TEXT NOT NULL,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY(series,date,source)
        );
        CREATE TABLE IF NOT EXISTS consensus_snapshots(
            symbol TEXT NOT NULL, observed_at TEXT NOT NULL, fiscal_date TEXT,
            metric TEXT NOT NULL, mean REAL, low REAL, high REAL, analyst_count REAL,
            provider TEXT NOT NULL, note TEXT, fetched_at TEXT NOT NULL,
            PRIMARY KEY(symbol,observed_at,fiscal_date,metric,provider)
        );
        CREATE TABLE IF NOT EXISTS features(
            symbol TEXT NOT NULL, as_of TEXT NOT NULL, target_date TEXT,
            revenue_growth REAL, operating_margin REAL, net_margin REAL, fcf_margin REAL,
            capex_to_revenue REAL, rd_to_revenue REAL, roe REAL, net_debt_to_revenue REAL,
            price_to_sales REAL, earnings_yield REAL, fcf_yield REAL, book_to_market REAL,
            ev_to_sales REAL, momentum_12m REAL, momentum_6m REAL, volatility_6m REAL,
            drawdown_12m REAL, target_excess_return_12m REAL,
            source TEXT NOT NULL, built_at TEXT NOT NULL,
            PRIMARY KEY(symbol,as_of,source)
        );
        CREATE INDEX IF NOT EXISTS idx_features_asof ON features(as_of);
        CREATE TABLE IF NOT EXISTS provider_state(
            provider TEXT NOT NULL, state_key TEXT NOT NULL, state_value TEXT,
            updated_at TEXT NOT NULL,
            PRIMARY KEY(provider,state_key)
        );
        CREATE TABLE IF NOT EXISTS predictions(
            id INTEGER PRIMARY KEY AUTOINCREMENT, run_id TEXT, symbol TEXT NOT NULL,
            model TEXT NOT NULL, as_of TEXT NOT NULL, horizon_days INTEGER,
            prediction TEXT, confidence TEXT, features_json TEXT, model_version TEXT,
            realized_at TEXT, realized_value REAL, error REAL, created_at TEXT NOT NULL
        );
        """
        with self.connect() as con:
            con.executescript(ddl)
            con.execute(
                "INSERT INTO metadata(key,value,updated_at) VALUES('schema_version',?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at",
                (str(SCHEMA_VERSION),utc_now()),
            )

    def set_state(self,provider,key,value):
        with self.connect() as con:
            con.execute(
                "INSERT INTO provider_state(provider,state_key,state_value,updated_at) VALUES(?,?,?,?) "
                "ON CONFLICT(provider,state_key) DO UPDATE SET state_value=excluded.state_value,updated_at=excluded.updated_at",
                (provider,key,json.dumps(value,default=str),utc_now()),
            )

    def get_state(self,provider,key,default=None):
        with self.connect() as con:
            row=con.execute("SELECT state_value FROM provider_state WHERE provider=? AND state_key=?",(provider,key)).fetchone()
        if not row: return default
        try: return json.loads(row[0])
        except Exception: return row[0]

    def upsert_symbols(self,rows):
        now=utc_now(); payload=[]
        for r in rows:
            payload.append((
                str(r.get("symbol") or "").upper(),r.get("name"),r.get("sector"),r.get("industry"),
                r.get("exchange"),str(r.get("cik") or "") or None,r.get("currency"),r.get("universe_source"),now,now,
            ))
        payload=[x for x in payload if x[0]]
        if not payload: return 0
        with self.connect() as con:
            con.executemany(
                "INSERT INTO symbols(symbol,name,sector,industry,exchange,cik,currency,universe_source,first_seen,last_seen) "
                "VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(symbol) DO UPDATE SET "
                "name=COALESCE(excluded.name,symbols.name),sector=COALESCE(excluded.sector,symbols.sector),"
                "industry=COALESCE(excluded.industry,symbols.industry),exchange=COALESCE(excluded.exchange,symbols.exchange),"
                "cik=COALESCE(excluded.cik,symbols.cik),currency=COALESCE(excluded.currency,symbols.currency),"
                "universe_source=COALESCE(excluded.universe_source,symbols.universe_source),last_seen=excluded.last_seen",
                payload,
            )
        return len(payload)

    def upsert_prices(self,rows):
        now=utc_now(); payload=[]
        for r in rows:
            symbol=str(r.get("symbol") or "").upper(); date=str(r.get("date") or "")[:10]; source=str(r.get("source") or "")
            if not symbol or not date or not source: continue
            payload.append((symbol,date,_finite(r.get("close")),_finite(r.get("adj_close")),_finite(r.get("volume")),source,now))
        if not payload: return 0
        with self.connect() as con:
            con.executemany(
                "INSERT INTO prices(symbol,date,close,adj_close,volume,source,fetched_at) VALUES(?,?,?,?,?,?,?) "
                "ON CONFLICT(symbol,date,source) DO UPDATE SET close=excluded.close,adj_close=excluded.adj_close,"
                "volume=excluded.volume,fetched_at=excluded.fetched_at",payload,
            )
        return len(payload)

    def upsert_fundamentals(self,rows):
        now=utc_now(); cols=[
            "symbol","fiscal_year","fiscal_date","available_date","period","revenue","operating_income","net_income",
            "eps","shares","ocf","capex","fcf","cash","debt","equity","assets","liabilities","rd","sbc","currency","source"
        ]; payload=[]
        numeric=set(cols[5:20])
        for r in rows:
            if not r.get("symbol") or r.get("fiscal_year") is None or not r.get("available_date") or not r.get("source"): continue
            vals=[]
            for c in cols:
                v=r.get(c)
                if c=="symbol": v=str(v).upper()
                if c in numeric: v=_finite(v)
                vals.append(v)
            payload.append(tuple(vals+[now]))
        if not payload: return 0
        q="INSERT INTO fundamentals("+",".join(cols)+",fetched_at) VALUES("+",".join(["?"]*(len(cols)+1))+ ") "
        q+="ON CONFLICT(symbol,fiscal_year,period,source) DO UPDATE SET "+",".join(
            f"{c}=excluded.{c}" for c in cols[2:]
        )+",fetched_at=excluded.fetched_at"
        with self.connect() as con: con.executemany(q,payload)
        return len(payload)

    def upsert_earnings(self,rows):
        now=utc_now(); payload=[]
        for r in rows:
            symbol=str(r.get("symbol") or "").upper(); reported=str(r.get("reported_date") or "")[:10]; source=str(r.get("source") or "")
            if not symbol or not reported or not source: continue
            payload.append((symbol,reported,str(r.get("fiscal_date") or "")[:10] or None,_finite(r.get("eps_actual")),
                            _finite(r.get("eps_estimate")),_finite(r.get("surprise_pct")),_finite(r.get("revenue_actual")),
                            _finite(r.get("revenue_estimate")),source,now))
        if not payload: return 0
        with self.connect() as con:
            con.executemany(
                "INSERT INTO earnings(symbol,reported_date,fiscal_date,eps_actual,eps_estimate,surprise_pct,revenue_actual,revenue_estimate,source,fetched_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(symbol,reported_date,source) DO UPDATE SET "
                "fiscal_date=excluded.fiscal_date,eps_actual=excluded.eps_actual,eps_estimate=excluded.eps_estimate,"
                "surprise_pct=excluded.surprise_pct,revenue_actual=excluded.revenue_actual,revenue_estimate=excluded.revenue_estimate,fetched_at=excluded.fetched_at",
                payload,
            )
        return len(payload)

    def upsert_macro(self,rows):
        now=utc_now(); payload=[]
        for r in rows:
            if not r.get("series") or not r.get("date") or not r.get("source"): continue
            payload.append((r["series"],str(r["date"])[:10],_finite(r.get("value")),r["source"],now))
        if not payload: return 0
        with self.connect() as con:
            con.executemany(
                "INSERT INTO macro(series,date,value,source,fetched_at) VALUES(?,?,?,?,?) "
                "ON CONFLICT(series,date,source) DO UPDATE SET value=excluded.value,fetched_at=excluded.fetched_at",payload,
            )
        return len(payload)

    def upsert_consensus(self,rows):
        now=utc_now(); payload=[]
        for r in rows:
            if not r.get("symbol") or not r.get("observed_at") or not r.get("metric") or not r.get("provider"): continue
            payload.append((str(r["symbol"]).upper(),str(r["observed_at"]),str(r.get("fiscal_date") or "")[:10] or None,r["metric"],
                            _finite(r.get("mean")),_finite(r.get("low")),_finite(r.get("high")),_finite(r.get("analyst_count")),
                            r["provider"],r.get("note"),now))
        if not payload: return 0
        with self.connect() as con:
            con.executemany(
                "INSERT INTO consensus_snapshots(symbol,observed_at,fiscal_date,metric,mean,low,high,analyst_count,provider,note,fetched_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(symbol,observed_at,fiscal_date,metric,provider) DO UPDATE SET "
                "mean=excluded.mean,low=excluded.low,high=excluded.high,analyst_count=excluded.analyst_count,note=excluded.note,fetched_at=excluded.fetched_at",payload,
            )
        return len(payload)

    def table_counts(self):
        names=["symbols","prices","fundamentals","earnings","macro","consensus_snapshots","features","predictions"]
        out={}
        with self.connect() as con:
            for n in names: out[n]=con.execute(f"SELECT COUNT(*) FROM {n}").fetchone()[0]
        return out

    def symbols(self,limit=None):
        q="SELECT symbol FROM symbols ORDER BY symbol"+(" LIMIT ?" if limit else "")
        with self.connect() as con:
            rows=con.execute(q,(limit,) if limit else ()).fetchall()
        return [r[0] for r in rows]

    def price_frame(self,symbol,source_preference=("Yahoo Finance","Alpha Vantage")):
        with self.connect() as con:
            df=pd.read_sql_query("SELECT date,close,adj_close,volume,source FROM prices WHERE symbol=? ORDER BY date",con,params=(symbol.upper(),))
        if df.empty: return df
        df["date"]=pd.to_datetime(df["date"])
        rank={s:i for i,s in enumerate(source_preference)}
        df["_rank"]=df["source"].map(rank).fillna(999)
        df=df.sort_values(["date","_rank"]).drop_duplicates("date",keep="first").drop(columns="_rank").set_index("date")
        return df

    def expected_return_frame(self,min_rows=30):
        with self.connect() as con:
            df=pd.read_sql_query("SELECT * FROM features WHERE target_excess_return_12m IS NOT NULL ORDER BY as_of,symbol",con)
        if len(df)<min_rows: return pd.DataFrame()
        df["as_of"]=pd.to_datetime(df["as_of"]); df["target_date"]=pd.to_datetime(df["target_date"])
        return df

    def earnings_model_frame(self,symbol):
        with self.connect() as con:
            e=pd.read_sql_query(
                "SELECT reported_date,eps_actual,eps_estimate,surprise_pct FROM earnings WHERE symbol=? AND surprise_pct IS NOT NULL ORDER BY reported_date",
                con,params=(symbol.upper(),),
            )
        if e.empty: return e
        e["reported_date"]=pd.to_datetime(e["reported_date"])
        prices=self.price_frame(symbol)
        rows=[]
        for i,row in e.iterrows():
            if i<4: continue
            dt=row["reported_date"]
            pre=prices.loc[:dt].tail(63) if not prices.empty else pd.DataFrame()
            px=pre["adj_close"].fillna(pre["close"]).dropna() if not pre.empty else pd.Series(dtype=float)
            prior=e.loc[:i-1,"surprise_pct"]
            rows.append({
                "as_of":dt-pd.Timedelta(days=1),"target_date":dt,
                "prior_surprise_1q":prior.iloc[-1],"prior_surprise_2q_avg":prior.tail(2).mean(),
                "prior_surprise_4q_avg":prior.tail(4).mean(),"surprise_vol_4q":prior.tail(4).std(),
                "price_momentum_3m":float(px.iloc[-1]/px.iloc[0]-1) if len(px)>=40 else None,
                "price_vol_3m":float(px.pct_change().std()*np.sqrt(252)) if len(px)>=40 else None,
                "eps_estimate":row["eps_estimate"],"target_surprise":row["surprise_pct"],
            })
        return pd.DataFrame(rows)

    def build_features(self,benchmark="SPY"):
        """Materialize point-in-time company-year features and 12m excess-return targets."""
        with self.connect() as con:
            fundamentals=pd.read_sql_query(
                "SELECT * FROM fundamentals WHERE period='annual' ORDER BY symbol,available_date,"
                "CASE source WHEN 'Alpha Vantage' THEN 0 WHEN 'SEC Company Facts' THEN 1 WHEN 'Yahoo Finance' THEN 2 ELSE 9 END",
                con,
            )
        if fundamentals.empty: return 0
        fundamentals=fundamentals.drop_duplicates(["symbol","fiscal_year"],keep="first")
        bench=self.price_frame(benchmark)
        bench_px=bench["adj_close"].fillna(bench["close"]).dropna() if not bench.empty else pd.Series(dtype=float)
        rows=[]
        for symbol,g in fundamentals.groupby("symbol"):
            g=g.sort_values("available_date").reset_index(drop=True); pf=self.price_frame(symbol)
            if pf.empty: continue
            px=pf["adj_close"].fillna(pf["close"]).dropna()
            for i,r in g.iterrows():
                as_of=pd.Timestamp(r["available_date"]); target=as_of+pd.DateOffset(years=1)
                before=px.loc[:as_of]
                after=px.loc[px.index>=as_of]
                future=px.loc[px.index>=target]
                if before.empty or after.empty or future.empty: continue
                p0=float(after.iloc[0]); p1=float(future.iloc[0]); stock_ret=p1/p0-1
                bench_ret=0.0
                if not bench_px.empty:
                    b0=bench_px.loc[bench_px.index>=as_of]; b1=bench_px.loc[bench_px.index>=target]
                    if not b0.empty and not b1.empty: bench_ret=float(b1.iloc[0]/b0.iloc[0]-1)
                rev=_finite(r["revenue"]); op=_finite(r["operating_income"]); ni=_finite(r["net_income"])
                ocf=_finite(r["ocf"]); cap=_finite(r["capex"]); fcf=_finite(r["fcf"])
                if fcf is None and ocf is not None and cap is not None: fcf=ocf-abs(cap)
                equity=_finite(r["equity"]); debt=_finite(r["debt"]) or 0; cash=_finite(r["cash"]) or 0
                rd=_finite(r["rd"]); shares=_finite(r["shares"])
                prev_rev=_finite(g.iloc[i-1]["revenue"]) if i>0 else None
                market_cap=p0*shares if shares and shares>1e5 else None
                ev=(market_cap+debt-cash) if market_cap is not None else None
                p12=before.tail(252); p6=before.tail(126)
                dd=None
                if len(p12)>=20:
                    running=p12.cummax(); dd=float((p12/running-1).min())
                rows.append({
                    "symbol":symbol,"as_of":as_of.date().isoformat(),"target_date":target.date().isoformat(),
                    "revenue_growth":rev/prev_rev-1 if rev and prev_rev else None,
                    "operating_margin":op/rev if rev and op is not None else None,
                    "net_margin":ni/rev if rev and ni is not None else None,
                    "fcf_margin":fcf/rev if rev and fcf is not None else None,
                    "capex_to_revenue":abs(cap)/rev if rev and cap is not None else None,
                    "rd_to_revenue":rd/rev if rev and rd is not None else None,
                    "roe":ni/equity if ni is not None and equity not in (None,0) else None,
                    "net_debt_to_revenue":(debt-cash)/rev if rev else None,
                    "price_to_sales":market_cap/rev if market_cap is not None and rev else None,
                    "earnings_yield":ni/market_cap if market_cap and ni is not None else None,
                    "fcf_yield":fcf/market_cap if market_cap and fcf is not None else None,
                    "book_to_market":equity/market_cap if market_cap and equity is not None else None,
                    "ev_to_sales":ev/rev if ev is not None and rev else None,
                    "momentum_12m":float(p12.iloc[-1]/p12.iloc[0]-1) if len(p12)>=126 else None,
                    "momentum_6m":float(p6.iloc[-1]/p6.iloc[0]-1) if len(p6)>=63 else None,
                    "volatility_6m":float(p6.pct_change().std()*np.sqrt(252)) if len(p6)>=63 else None,
                    "drawdown_12m":dd,"target_excess_return_12m":stock_ret-bench_ret,
                    "source":"Persistent point-in-time history v1","built_at":utc_now(),
                })
        if not rows: return 0
        cols=list(rows[0]); placeholders=",".join(["?"]*len(cols))
        q="INSERT INTO features("+",".join(cols)+") VALUES("+placeholders+") ON CONFLICT(symbol,as_of,source) DO UPDATE SET "+",".join(
            f"{c}=excluded.{c}" for c in cols if c not in {"symbol","as_of","source"}
        )
        with self.connect() as con:
            con.executemany(q,[tuple(_finite(r[c]) if c not in {"symbol","as_of","target_date","source","built_at"} else r[c] for c in cols) for r in rows])
        return len(rows)
