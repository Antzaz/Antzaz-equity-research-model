"""
Clean Technology Equity Research Updater v7
Usage:
    python update_model.py GOOGL

Place this script and GOOGL_Equity_Research_CLEAN_v7.xlsx in the same folder.
Install once:
    python -m pip install requests yfinance openpyxl

Model architecture:
SEC annual fundamentals -> Historical Financials -> Three-Case Scenarios
-> DCF -> Dashboard. Comparative Analysis pulls directly from Peer Comps.
"""

import sys, subprocess, importlib, re
from pathlib import Path
from datetime import datetime

deps={"requests":"requests","yfinance":"yfinance","openpyxl":"openpyxl"}
for module,package in deps.items():
    try: importlib.import_module(module)
    except ImportError:
        subprocess.check_call([sys.executable,"-m","pip","install",package])

import requests, yfinance as yf
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

BASE=Path(__file__).resolve().parent
TEMPLATE=BASE/"GOOGL_Equity_Research_CLEAN_v7.xlsx"
OUTDIR=BASE/"updated_models"; OUTDIR.mkdir(exist_ok=True)
SEC_HEADERS={"User-Agent":"Personal equity research model research@example.com"}

PEER_GROUPS={
    "GOOGL":["MSFT","META","AMZN","AAPL","NFLX"],
    "GOOG":["MSFT","META","AMZN","AAPL","NFLX"],
    "NVDA":["AMD","AVGO","TSM","INTC","QCOM"],
    "MSFT":["ORCL","CRM","ADBE","NOW","GOOGL"],
    "META":["GOOGL","AMZN","NFLX","MSFT","PINS"],
}
DEFAULT_PEERS=["MSFT","META","AMZN","AAPL","NFLX"]

def sec_json(url):
    r=requests.get(url,headers=SEC_HEADERS,timeout=30); r.raise_for_status(); return r.json()

def cik_for(ticker):
    for item in sec_json("https://www.sec.gov/files/company_tickers.json").values():
        if item["ticker"].upper()==ticker.upper():
            return str(item["cik_str"]).zfill(10)
    return None

def company_facts(ticker):
    cik=cik_for(ticker)
    return sec_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json") if cik else None

def annual_series(facts,tags,preferred_unit=None):
    if not facts: return {}
    gaap=facts.get("facts",{}).get("us-gaap",{})
    for tag in tags:
        if tag not in gaap: continue
        units=gaap[tag].get("units",{})
        unit=preferred_unit if preferred_unit in units else (next(iter(units)) if units else None)
        if not unit: continue
        out={}
        for x in units[unit]:
            if x.get("form") not in ("10-K","10-K/A") or x.get("fp")!="FY": continue
            if x.get("fy") is None or x.get("val") is None: continue
            out[int(x["fy"])]=float(x["val"])
        if out: return out
    return {}

def build_history(ticker):
    facts=company_facts(ticker)
    revenue=annual_series(facts,["RevenueFromContractWithCustomerExcludingAssessedTax","SalesRevenueNet","Revenues"])
    cost=annual_series(facts,["CostOfRevenue","CostOfGoodsAndServicesSold"])
    gross=annual_series(facts,["GrossProfit"])
    op=annual_series(facts,["OperatingIncomeLoss"])
    ni=annual_series(facts,["NetIncomeLoss"])
    eps=annual_series(facts,["EarningsPerShareDiluted"],"USD/shares")
    ocf=annual_series(facts,["NetCashProvidedByUsedInOperatingActivities"])
    capex=annual_series(facts,["PaymentsToAcquirePropertyPlantAndEquipment","PaymentsToAcquireProductiveAssets"])
    depr=annual_series(facts,["DepreciationDepletionAndAmortizationPropertyPlantAndEquipment","DepreciationDepletionAndAmortization"])
    rd=annual_series(facts,["ResearchAndDevelopmentExpense"])
    sbc=annual_series(facts,["ShareBasedCompensation"])
    years=sorted(revenue)[-6:] if revenue else []
    out={}
    for y in years:
        rev=revenue.get(y); c=cost.get(y); gp=gross.get(y)
        if gp is None and rev is not None and c is not None: gp=rev-c
        cx=abs(capex[y]) if y in capex else None
        fcf=ocf.get(y)-cx if ocf.get(y) is not None and cx is not None else None
        shares=ni[y]/eps[y] if y in ni and y in eps and eps[y] else None
        out[y]={"revenue":rev,"cost":c,"gross":gp,"op":op.get(y),"ni":ni.get(y),"eps":eps.get(y),
                "shares":shares,"ocf":ocf.get(y),"capex":cx,"fcf":fcf,"depr":depr.get(y),"rd":rd.get(y),"sbc":sbc.get(y)}
    return out

def yf_info(ticker): return yf.Ticker(ticker).info or {}

def put_company(wb,ticker,info):
    ws=wb["Company Data"]
    p=info.get("currentPrice") or info.get("regularMarketPrice")
    vals={"B4":ticker,"B5":info.get("longName") or ticker,"B6":info.get("sector"),"B7":info.get("industry"),
          "B8":p,"B10":info.get("marketCap")/1e9 if info.get("marketCap") else None,
          "B11":info.get("enterpriseValue")/1e9 if info.get("enterpriseValue") else None,
          "B12":info.get("totalCash")/1e9 if info.get("totalCash") else None,
          "B13":info.get("totalDebt")/1e9 if info.get("totalDebt") else None,
          "B15":info.get("forwardPE")}
    for c,v in vals.items(): ws[c]=v

def put_history(wb,hist):
    if not hist:
        print("Warning: SEC annual history unavailable; keeping template historical data.")
        return
    ws = wb["Historical Financials"]
    years = sorted(hist)[-6:]
    for i, y in enumerate(years, 2):
        d = hist[y]
        scale = lambda x: x/1e9 if x is not None else None
        ws.cell(3,i).value = y
        ws.cell(4,i).value = scale(d.get("revenue"))
        ws.cell(6,i).value = scale(d.get("cost"))
        ws.cell(9,i).value = scale(d.get("op"))
        ws.cell(11,i).value = scale(d.get("ni"))
        ws.cell(12,i).value = d.get("eps")
        ws.cell(14,i).value = scale(d.get("ocf"))
        ws.cell(15,i).value = scale(d.get("capex"))
        ws.cell(18,i).value = scale(d.get("depr"))
        ws.cell(19,i).value = scale(d.get("rd"))
        ws.cell(21,i).value = scale(d.get("sbc"))

def fade(start,end): return [start+(end-start)*i/9 for i in range(10)]

def _template_history_fallback(wb):
    ws = wb["Historical Financials"]
    out = {}
    for c in range(2, 8):
        year = ws.cell(3, c).value
        try: year = int(year)
        except Exception: continue
        def bn_to_raw(v): return float(v) * 1e9 if isinstance(v, (int, float)) else None
        out[year] = {
            "revenue": bn_to_raw(ws.cell(4, c).value),
            "op": bn_to_raw(ws.cell(9, c).value),
            "capex": bn_to_raw(ws.cell(15, c).value),
            "depr": bn_to_raw(ws.cell(18, c).value),
            "ocf": bn_to_raw(ws.cell(14, c).value),
            "fcf": bn_to_raw(ws.cell(16, c).value),
            "ni": bn_to_raw(ws.cell(11, c).value),
            "eps": ws.cell(12, c).value if isinstance(ws.cell(12, c).value, (int, float)) else None,
            "cost": bn_to_raw(ws.cell(6, c).value),
            "gross": bn_to_raw(ws.cell(7, c).value),
            "rd": bn_to_raw(ws.cell(19, c).value),
            "sbc": bn_to_raw(ws.cell(21, c).value),
            "shares": None,
        }
    return out

def update_scenarios(wb,hist,info):
    ws = wb["Three-Case Scenarios"]
    fallback = _template_history_fallback(wb)
    merged = dict(fallback); merged.update(hist or {})
    years = sorted(y for y, d in merged.items() if d and d.get("revenue"))
    hc = info.get("revenueGrowth") or 0.10
    opm = info.get("operatingMargins") or 0.30
    cap = 0.20; dep = 0.05
    if len(years) >= 2:
        last = years[-1]; first = years[max(0, len(years)-6)]
        r0 = merged[first].get("revenue"); r1 = merged[last].get("revenue")
        n = max(1, last-first)
        if r0 and r1 and r0 > 0 and r1 > 0: hc = (r1/r0)**(1/n)-1
        latest = merged[last]; rev = latest.get("revenue")
        if rev:
            if latest.get("op") is not None: opm = latest["op"]/rev
            if latest.get("capex") is not None: cap = abs(latest["capex"])/rev
            if latest.get("depr") is not None: dep = abs(latest["depr"])/rev
    hc = max(0.02, min(0.20, float(hc)))
    opm = max(0.10, min(0.45, float(opm)))
    cap = max(0.05, min(0.35, float(cap)))
    dep = max(0.01, min(0.10, float(dep)))
    base = max(.08, min(.18, hc))
    growths = [fade(max(.03, base-.05), .04), fade(base, .06), fade(min(.22, base+.03), .08)]
    blocks = [range(2,12), range(14,24), range(26,36)]
    for cols, gs in zip(blocks, growths):
        for c, v in zip(cols, gs): ws.cell(12, c).value = v
    margins = [
        fade(max(.15, opm-.03), max(.15, opm-.04)),
        fade(opm, min(.40, opm+.01)),
        fade(min(.45, opm+.01), min(.45, opm+.04)),
    ]
    for cols, ms in zip(blocks, margins):
        for c, v in zip(cols, ms): ws.cell(14, c).value = v
    capex_sets = [fade(min(.35, cap+.05), .16), fade(min(.35, cap+.01), .13), fade(max(.12, cap-.01), .12)]
    for cols, cs in zip(blocks, capex_sets):
        for c, v in zip(cols, cs): ws.cell(20, c).value = v
    for cols in blocks:
        for c in cols: ws.cell(18, c).value = dep

def update_peers(wb,ticker):
    peers=PEER_GROUPS.get(ticker,DEFAULT_PEERS); rows=[ticker]+peers
    ws=wb["Peer Comps"]
    for r in range(4,10):
        for c in range(1,9): ws.cell(r,c).value=None
    for r,sym in enumerate(rows[:6],4):
        info=yf_info(sym)
        vals=[info.get("longName") or sym,sym,info.get("forwardPE"),info.get("enterpriseToRevenue"),
              info.get("enterpriseToEbitda"),info.get("revenueGrowth"),info.get("operatingMargins"),info.get("returnOnEquity")]
        for c,v in enumerate(vals,1): ws.cell(r,c).value=v

def update_filings(wb,ticker):
    cik=cik_for(ticker)
    if not cik: return
    recent=sec_json(f"https://data.sec.gov/submissions/CIK{cik}.json").get("filings",{}).get("recent",{})
    ws=wb["Filings"]
    for r in range(4,20):
        for c in range(1,6): ws.cell(r,c).value=None
    row=4
    for form,period,filed,acc,doc in zip(recent.get("form",[]),recent.get("reportDate",[]),recent.get("filingDate",[]),recent.get("accessionNumber",[]),recent.get("primaryDocument",[])):
        if form not in {"10-K","10-Q","8-K","DEF 14A"}: continue
        url=f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-','')}/{doc}"
        for c,v in enumerate([form,period,filed,url,"SEC filing"],1): ws.cell(row,c).value=v
        row+=1
        if row>15: break

def get_ticker():
    raw = sys.argv[1] if len(sys.argv) > 1 else input("Ticker (e.g. GOOGL): ")
    raw = raw.strip()
    if " " in raw or ".py" in raw.lower():
        parts = raw.replace('"', '').replace("'", "").split()
        candidates = [p.upper() for p in parts
                      if re.fullmatch(r"[A-Za-z0-9.\-]{1,10}", p)
                      and not p.lower().endswith(".py")
                      and p.lower() not in {"python", "py"}]
        if candidates: raw = candidates[-1]
    ticker = raw.upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}", ticker):
        raise ValueError(f"Invalid ticker: {ticker!r}. Enter only a ticker, e.g. GOOGL, MSFT, or NVDA.")
    return ticker

def main():
    ticker = get_ticker()
    print(f"Ticker: {ticker}")
    if not TEMPLATE.exists(): raise FileNotFoundError(TEMPLATE)
    info=yf_info(ticker); hist=build_history(ticker)
    wb=load_workbook(TEMPLATE,data_only=False)
    put_company(wb,ticker,info); put_history(wb,hist); update_scenarios(wb,hist,info); update_peers(wb,ticker); update_filings(wb,ticker)
    wb["Dashboard"]["A1"]=f"{ticker} Long-Term Value Investing Dashboard"
    try:
        if getattr(wb, "calculation", None) is None:
            wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
        else:
            wb.calculation.calcMode = "auto"
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
    except Exception as exc:
        print(f"Warning: could not set workbook recalculation flags: {exc}")
    out=OUTDIR/f"{ticker}_Equity_Research_{datetime.now():%Y%m%d_%H%M%S}.xlsx"; wb.save(out)
    print("Saved:",out)

if __name__=="__main__": main()
