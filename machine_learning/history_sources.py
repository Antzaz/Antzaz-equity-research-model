from __future__ import annotations

"""Network adapters for the persistent ML history database.

Bulk/default policy:
- Yahoo Finance: broad price history and a no-key annual-statement fallback.
- SEC Company Facts: deep US annual fundamentals when a compliant SEC_USER_AGENT is configured.
- Alpha Vantage: quota-aware normalized statements + earnings history enrichment.
- FRED graph CSV: official macro history without a FRED API key.
- FMP: optional analyst/earnings enrichment when FMP_API_KEY is configured.

Provider responses are never silently treated as equivalent. Every stored row carries a source.
"""

from datetime import datetime, timezone
import io
import os
import time

import numpy as np
import pandas as pd
import requests
import yfinance as yf

from issuer_source_engine import sec_crossborder_history
from .data import DEFAULT_UNIVERSE


AV_URL="https://www.alphavantage.co/query"
FMP_URL="https://financialmodelingprep.com/stable"
FRED_CSV="https://fred.stlouisfed.org/graph/fredgraph.csv"

FRED_SERIES={
    "DGS10":"10Y Treasury yield",
    "DGS2":"2Y Treasury yield",
    "T10Y2":"10Y-2Y Treasury spread",
    "FEDFUNDS":"Effective federal funds rate",
    "CPIAUCSL":"Consumer Price Index",
    "UNRATE":"Unemployment rate",
    "INDPRO":"Industrial production index",
    "BAMLH0A0HYM2":"US high-yield option-adjusted spread",
}


def _num(v):
    try:
        if v in (None,"","None","null","-"): return None
        x=float(v)
        return x if np.isfinite(x) else None
    except Exception: return None


def current_sp500_universe(limit=500):
    """Current S&P 500 constituents. This is a current-universe convenience, not survivorship-free history."""
    try:
        tables=pd.read_html("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies")
        df=tables[0]
        rows=[]
        for _,r in df.iterrows():
            symbol=str(r.get("Symbol") or "").replace(".","-").strip().upper()
            if not symbol: continue
            rows.append({
                "symbol":symbol,"name":r.get("Security"),"sector":r.get("GICS Sector"),
                "industry":r.get("GICS Sub-Industry"),"cik":str(r.get("CIK") or "").split(".")[0] or None,
                "universe_source":"Current S&P 500 constituent snapshot",
            })
        return rows[:limit]
    except Exception:
        return [{"symbol":s,"universe_source":"Built-in fallback universe"} for s in DEFAULT_UNIVERSE[:limit]]


def yahoo_price_rows(symbols,years=20,batch_size=40):
    symbols=list(dict.fromkeys(str(s).upper() for s in symbols if s))
    period=f"{int(years)}y"
    for start in range(0,len(symbols),batch_size):
        batch=symbols[start:start+batch_size]
        try:
            raw=yf.download(batch,period=period,auto_adjust=False,actions=False,progress=False,threads=True,group_by="ticker")
        except Exception:
            raw=pd.DataFrame()
        if raw is None or raw.empty:
            continue
        for symbol in batch:
            try:
                if len(batch)==1 and not isinstance(raw.columns,pd.MultiIndex):
                    df=raw.copy()
                elif isinstance(raw.columns,pd.MultiIndex) and symbol in raw.columns.get_level_values(0):
                    df=raw[symbol].copy()
                elif isinstance(raw.columns,pd.MultiIndex) and symbol in raw.columns.get_level_values(-1):
                    df=raw.xs(symbol,axis=1,level=-1).copy()
                else:
                    continue
                if "Close" not in df: continue
                adj=df["Adj Close"] if "Adj Close" in df else df["Close"]
                vol=df["Volume"] if "Volume" in df else pd.Series(index=df.index,dtype=float)
                for dt,close in pd.to_numeric(df["Close"],errors="coerce").items():
                    if pd.isna(close): continue
                    av=_num(adj.loc[dt]) if dt in adj.index else None; vv=_num(vol.loc[dt]) if dt in vol.index else None
                    yield {"symbol":symbol,"date":pd.Timestamp(dt).date().isoformat(),"close":float(close),"adj_close":av,"volume":vv,"source":"Yahoo Finance"}
            except Exception:
                continue


def _series_value(df,labels,col):
    if df is None or getattr(df,"empty",True) or col not in df.columns: return None
    for label in labels:
        if label in df.index: return _num(df.at[label,col])
    return None


def yahoo_fundamental_rows(symbol):
    try:
        t=yf.Ticker(symbol); inc=t.income_stmt; cf=t.cashflow; bs=t.balance_sheet; info=t.info or {}
    except Exception:
        return []
    if inc is None or inc.empty: return []
    out=[]; cols=sorted(pd.to_datetime(inc.columns))
    for col in cols:
        fiscal=pd.Timestamp(col).tz_localize(None); year=int(fiscal.year)
        rev=_series_value(inc,["Total Revenue","Operating Revenue"],col)
        ni=_series_value(inc,["Net Income","Net Income Common Stockholders"],col)
        eps=_series_value(inc,["Diluted EPS"],col)
        op=_series_value(inc,["Operating Income"],col)
        ocf=_series_value(cf,["Operating Cash Flow","Total Cash From Operating Activities"],col)
        cap=_series_value(cf,["Capital Expenditure","Capital Expenditures"],col); cap=abs(cap) if cap is not None else None
        rd=_series_value(inc,["Research And Development"],col); sbc=_series_value(cf,["Stock Based Compensation"],col)
        cash=_series_value(bs,["Cash Cash Equivalents And Short Term Investments","Cash And Cash Equivalents"],col)
        debt=_series_value(bs,["Total Debt"],col); equity=_series_value(bs,["Stockholders Equity","Total Equity Gross Minority Interest"],col)
        assets=_series_value(bs,["Total Assets"],col); liabilities=_series_value(bs,["Total Liabilities Net Minority Interest","Total Liabilities"],col)
        shares=ni/eps if ni is not None and eps not in (None,0) else None
        if rev is None and ni is None: continue
        out.append({
            "symbol":symbol,"fiscal_year":year,"fiscal_date":fiscal.date().isoformat(),
            "available_date":(fiscal+pd.Timedelta(days=120)).date().isoformat(),"period":"annual",
            "revenue":rev,"operating_income":op,"net_income":ni,"eps":eps,"shares":shares,
            "ocf":ocf,"capex":cap,"fcf":ocf-cap if ocf is not None and cap is not None else None,
            "cash":cash,"debt":debt,"equity":equity,"assets":assets,"liabilities":liabilities,"rd":rd,"sbc":sbc,
            "currency":info.get("financialCurrency"),"source":"Yahoo Finance",
        })
    return out


def sec_cik_map(user_agent=None):
    ua=(user_agent or os.getenv("SEC_USER_AGENT") or "").strip()
    if not ua: return {}
    r=requests.get("https://www.sec.gov/files/company_tickers.json",headers={"User-Agent":ua},timeout=30)
    r.raise_for_status(); data=r.json(); out={}
    for item in data.values():
        s=str(item.get("ticker") or "").upper(); cik=item.get("cik_str")
        if s and cik is not None: out[s]=str(cik).zfill(10)
    return out


def sec_fundamental_rows(symbol,cik,user_agent=None):
    ua=(user_agent or os.getenv("SEC_USER_AGENT") or "").strip()
    if not ua or not cik: return []
    try:
        r=requests.get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{str(cik).zfill(10)}.json",headers={"User-Agent":ua},timeout=35)
        r.raise_for_status(); hist,_=sec_crossborder_history(r.json())
    except Exception:
        return []
    out=[]
    for year,d in sorted(hist.items()):
        # The existing SEC parser validates annual duration. Dec-31 +120d is deliberately conservative
        # when exact filing availability is not retained by the normalized history object.
        fiscal=pd.Timestamp(year=year,month=12,day=31); eps=_num(d.get("eps")); ni=_num(d.get("ni")); shares=_num(d.get("shares"))
        if shares is None and ni is not None and eps not in (None,0): shares=ni/eps
        out.append({
            "symbol":symbol,"fiscal_year":year,"fiscal_date":fiscal.date().isoformat(),
            "available_date":(fiscal+pd.Timedelta(days=120)).date().isoformat(),"period":"annual",
            "revenue":d.get("revenue"),"operating_income":d.get("op"),"net_income":ni,"eps":eps,"shares":shares,
            "ocf":d.get("ocf"),"capex":d.get("capex"),"fcf":d.get("fcf"),"cash":d.get("cash"),"debt":d.get("debt"),
            "equity":d.get("equity"),"assets":d.get("assets"),"liabilities":d.get("liabilities"),"rd":d.get("rd"),"sbc":d.get("sbc"),
            "currency":"USD","source":"SEC Company Facts",
        })
    time.sleep(0.12)
    return out


class AlphaVantageClient:
    def __init__(self,key=None,max_calls=20):
        self.key=(key or os.getenv("ALPHAVANTAGE_API_KEY") or "").strip(); self.max_calls=max(0,int(max_calls)); self.calls=0

    @property
    def available(self): return bool(self.key) and self.calls<self.max_calls

    def get(self,function,symbol):
        if not self.available: return None
        self.calls+=1
        try:
            r=requests.get(AV_URL,params={"function":function,"symbol":symbol,"apikey":self.key},timeout=35)
            r.raise_for_status(); data=r.json()
            if not isinstance(data,dict) or data.get("Information") or data.get("Note") or data.get("Error Message"): return None
            return data
        except Exception: return None

    def enrich_symbol(self,symbol):
        if self.max_calls-self.calls<4: return [],[]
        inc=self.get("INCOME_STATEMENT",symbol) or {}; bs=self.get("BALANCE_SHEET",symbol) or {}
        cf=self.get("CASH_FLOW",symbol) or {}; earnings=self.get("EARNINGS",symbol) or {}
        by_date={}
        for row in inc.get("annualReports",[]) or []: by_date.setdefault(row.get("fiscalDateEnding"),{})["inc"]=row
        for row in bs.get("annualReports",[]) or []: by_date.setdefault(row.get("fiscalDateEnding"),{})["bs"]=row
        for row in cf.get("annualReports",[]) or []: by_date.setdefault(row.get("fiscalDateEnding"),{})["cf"]=row
        annual_eps={str(r.get("fiscalDateEnding")):r for r in earnings.get("annualEarnings",[]) or []}
        fundamentals=[]
        for fd,parts in sorted(by_date.items()):
            if not fd: continue
            fiscal=pd.Timestamp(fd); a=parts.get("inc",{}); b=parts.get("bs",{}); c=parts.get("cf",{}); er=annual_eps.get(fd,{})
            rev=_num(a.get("totalRevenue")); ni=_num(a.get("netIncome")); eps=_num(er.get("reportedEPS"))
            shares=_num(b.get("commonStockSharesOutstanding"));
            if shares is None and ni is not None and eps not in (None,0): shares=ni/eps
            ocf=_num(c.get("operatingCashflow")); cap=_num(c.get("capitalExpenditures")); cap=abs(cap) if cap is not None else None
            cash=_num(b.get("cashAndShortTermInvestments")) or _num(b.get("cashAndCashEquivalentsAtCarryingValue"))
            debt=_num(b.get("shortLongTermDebtTotal"));
            if debt is None: debt=(_num(b.get("longTermDebt")) or 0)+(_num(b.get("shortTermDebt")) or 0)
            fundamentals.append({
                "symbol":symbol,"fiscal_year":int(fiscal.year),"fiscal_date":fiscal.date().isoformat(),
                "available_date":(fiscal+pd.Timedelta(days=120)).date().isoformat(),"period":"annual",
                "revenue":rev,"operating_income":_num(a.get("operatingIncome")),"net_income":ni,"eps":eps,"shares":shares,
                "ocf":ocf,"capex":cap,"fcf":ocf-cap if ocf is not None and cap is not None else None,
                "cash":cash,"debt":debt,"equity":_num(b.get("totalShareholderEquity")),"assets":_num(b.get("totalAssets")),
                "liabilities":_num(b.get("totalLiabilities")),"rd":_num(a.get("researchAndDevelopment")),
                "sbc":_num(c.get("stockBasedCompensation")),"currency":a.get("reportedCurrency") or b.get("reportedCurrency"),"source":"Alpha Vantage",
            })
        earnings_rows=[]
        for r in earnings.get("quarterlyEarnings",[]) or []:
            reported=r.get("reportedDate"); fiscal=r.get("fiscalDateEnding")
            if not reported: continue
            est=_num(r.get("estimatedEPS")); actual=_num(r.get("reportedEPS")); surprise=_num(r.get("surprisePercentage"))
            if surprise is not None: surprise/=100.0
            earnings_rows.append({"symbol":symbol,"reported_date":reported,"fiscal_date":fiscal,"eps_actual":actual,"eps_estimate":est,"surprise_pct":surprise,"source":"Alpha Vantage"})
        return fundamentals,earnings_rows


def fred_macro_rows(series_map=None):
    series_map=series_map or FRED_SERIES
    for series in series_map:
        try:
            r=requests.get(FRED_CSV,params={"id":series},timeout=30); r.raise_for_status()
            df=pd.read_csv(io.StringIO(r.text)); date_col=df.columns[0]; value_col=df.columns[-1]
            values=pd.to_numeric(df[value_col],errors="coerce")
            for dt,val in zip(pd.to_datetime(df[date_col],errors="coerce"),values):
                if pd.isna(dt) or pd.isna(val): continue
                yield {"series":series,"date":dt.date().isoformat(),"value":float(val),"source":"Federal Reserve FRED"}
        except Exception:
            continue


class FMPClient:
    def __init__(self,key=None,max_calls=200):
        self.key=(key or os.getenv("FMP_API_KEY") or "").strip(); self.max_calls=max(0,int(max_calls)); self.calls=0
    @property
    def available(self): return bool(self.key) and self.calls<self.max_calls
    def get(self,path,params=None):
        if not self.available: return None
        self.calls+=1; params=dict(params or {}); params["apikey"]=self.key
        try:
            r=requests.get(FMP_URL+"/"+path.lstrip("/"),params=params,timeout=35); r.raise_for_status(); return r.json()
        except Exception: return None
    def enrich_symbol(self,symbol,observed_at=None):
        observed_at=observed_at or datetime.now(timezone.utc).isoformat(timespec="seconds")
        estimates=self.get("analyst-estimates",{"symbol":symbol,"period":"annual","page":0,"limit":20})
        rows=[]
        if isinstance(estimates,list):
            for r in estimates:
                fd=str(r.get("date") or "")[:10] or None
                for metric,stem,count_key in (("Revenue","revenue","numAnalystsRevenue"),("EPS","eps","numAnalystsEps")):
                    mean=_num(r.get(stem+"Avg")); low=_num(r.get(stem+"Low")); high=_num(r.get(stem+"High")); n=_num(r.get(count_key))
                    if mean is None: continue
                    rows.append({"symbol":symbol,"observed_at":observed_at,"fiscal_date":fd,"metric":metric,"mean":mean,"low":low,"high":high,
                                 "analyst_count":n,"provider":"FMP","note":"Stored as a provider snapshot. Revisions become point-in-time only as repeated snapshots accumulate."})
        earnings=self.get("earnings",{"symbol":symbol})
        erows=[]
        if isinstance(earnings,list):
            for r in earnings:
                dt=str(r.get("date") or "")[:10]
                if not dt: continue
                actual=_num(r.get("epsActual")); est=_num(r.get("epsEstimated")); surprise=(actual/est-1) if actual is not None and est not in (None,0) else None
                erows.append({"symbol":symbol,"reported_date":dt,"eps_actual":actual,"eps_estimate":est,"surprise_pct":surprise,
                              "revenue_actual":_num(r.get("revenueActual")),"revenue_estimate":_num(r.get("revenueEstimated")),"source":"FMP"})
        return rows,erows
