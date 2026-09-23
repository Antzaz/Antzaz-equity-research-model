from __future__ import annotations

"""Offline-only equity accountability and portfolio-link research extensions.

This module deliberately writes only to the private/offline workbook and local research_data
history. It does not publish anything to the recruiter/public showcase.

Implemented layers:
- forecast snapshots + forecast-vs-actual accountability;
- earnings / estimate revision deltas across model runs;
- capital-allocation history and incremental-return diagnostics;
- valuation snapshot history and historical percentiles;
- gated segment SOTP diagnostic;
- catalyst / thesis-change timeline.

All calculations are evidence-gated. Missing source rows remain REVIEW/N/M rather than being
silently estimated.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import csv
import math
import statistics

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREEN="E2F0D9"; GOLD="FFF2CC"; RED="FCE4D6"; GREY="666666"; LIGHT="F5F9FC"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'; FMT_MULT='0.0x;[Red](0.0x);-'
ROOT=Path(__file__).resolve().parent


def _fill(c): return PatternFill("solid", fgColor=c)
def _num(v, default=None):
    try:
        if isinstance(v, bool) or v in (None, ""): return default
        x=float(v)
        return x if math.isfinite(x) else default
    except Exception: return default


def _find(ws, label, cols=(1,)):
    needle=str(label).strip().lower()
    for c in cols:
        for r in range(1, ws.max_row+1):
            if str(ws.cell(r,c).value or "").strip().lower()==needle:
                return r,c
    return None


def _label_value(wb, sheet, label, cols=(1,), offset=1):
    if sheet not in wb.sheetnames: return None
    hit=_find(wb[sheet], label, cols)
    if not hit: return None
    r,c=hit
    return wb[sheet].cell(r,c+offset).value


def _year_cols(ws, header_row, max_col=40):
    out={}
    for c in range(2,min(ws.max_column,max_col)+1):
        v=ws.cell(header_row,c).value
        try:
            y=int(v)
            if 1990<=y<=2100: out[y]=c
        except Exception:
            pass
    return out


def _section(ws,row,title,end=9):
    for c in range(1,end+1):
        ws.cell(row,c).fill=_fill(NAVY); ws.cell(row,c).font=Font(bold=True,color=WHITE)
    ws.cell(row,1,title)


def _header(ws,row,vals):
    for c,v in enumerate(vals,1):
        x=ws.cell(row,c,v); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE)
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)


def _prep_sheet(wb,name,title,subtitle):
    if name in wb.sheetnames: wb.remove(wb[name])
    ws=wb.create_sheet(name)
    ws.sheet_view.showGridLines=False
    for c in range(1,10):
        ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=title; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws["A3"]=subtitle; ws.merge_cells("A3:I3"); ws["A3"].font=Font(italic=True,color=GREY)
    ws["A3"].alignment=Alignment(wrap_text=True)
    return ws


def _research_dir(root: str|Path, ticker: str):
    p=Path(root)/"research_data"/str(ticker).upper()
    p.mkdir(parents=True,exist_ok=True)
    return p


def _append_snapshot(path: Path, row: dict, unique_keys: tuple[str,...]):
    existing=pd.read_csv(path) if path.exists() else pd.DataFrame()
    new=pd.DataFrame([row])
    out=pd.concat([existing,new],ignore_index=True) if not existing.empty else new
    if all(k in out.columns for k in unique_keys):
        out=out.drop_duplicates(list(unique_keys),keep="last")
    out.to_csv(path,index=False)
    return out


def _historical_actuals(wb):
    """Best-effort annual actuals from the canonical Historical Financials sheet."""
    if "Historical Financials" not in wb.sheetnames: return {}
    ws=wb["Historical Financials"]
    header=3
    years=_year_cols(ws,header,12)
    if not years: return {}
    labels={}
    for r in range(1,ws.max_row+1):
        lab=str(ws.cell(r,1).value or "").strip().lower()
        if lab: labels[lab]=r
    candidates={
        "Revenue":["revenue","total revenue","total net revenue"],
        "Operating Income":["operating income","operating income / (loss)"],
        "Net Income":["net income","net income attributable to common","net income / (loss)"],
        "EPS":["diluted eps","eps diluted","diluted earnings per share"],
        "Operating Cash Flow":["operating cash flow","cash from operations","net cash provided by operating activities"],
        "Capex":["capital expenditures","capex","purchases of property plant and equipment"],
        "FCF":["free cash flow"],
        "Diluted Shares":["diluted shares","diluted weighted average shares"],
    }
    out={}
    for y,c in years.items():
        rec={}
        for metric,names in candidates.items():
            r=next((labels[n] for n in names if n in labels),None)
            if r: rec[metric]=_num(ws.cell(r,c).value)
        if rec.get("FCF") is None and rec.get("Operating Cash Flow") is not None and rec.get("Capex") is not None:
            cap=abs(rec["Capex"]); rec["FCF"]=rec["Operating Cash Flow"]-cap
        out[y]=rec
    return out


def _scenario_forecasts(wb):
    """Extract annual base-case revenue and FCF forecasts from Three-Case Scenarios."""
    if "Three-Case Scenarios" not in wb.sheetnames: return []
    ws=wb["Three-Case Scenarios"]
    # Find rows semantically and infer year header from the same region.
    rev_hit=_find(ws,"Revenue",cols=tuple(range(1,min(ws.max_column,15)+1)))
    fcf_hit=_find(ws,"Free Cash Flow",cols=tuple(range(1,min(ws.max_column,15)+1)))
    eps_hit=_find(ws,"EPS",cols=tuple(range(1,min(ws.max_column,15)+1)))
    hits=[x for x in [rev_hit,fcf_hit,eps_hit] if x]
    if not hits: return []
    # Production sheet commonly places forecast years above the scenario block. Search rows 1..20.
    year_map={}
    for r in range(1,min(ws.max_row,25)+1):
        cols=_year_cols(ws,r,40)
        if len(cols)>=2:
            year_map=cols
            header_row=r
            break
    if not year_map: return []
    rows=[]
    for metric,hit in [("Revenue",rev_hit),("FCF",fcf_hit),("EPS",eps_hit)]:
        if not hit: continue
        rr,_=hit
        for y,c in year_map.items():
            v=_num(ws.cell(rr,c).value)
            if v is not None: rows.append({"FiscalYear":y,"Metric":metric,"Forecast":v})
    return rows


def _consensus_rows(wb):
    if "Expectations & Consensus" not in wb.sheetnames: return []
    ws=wb["Expectations & Consensus"]; rows=[]
    for r in range(1,ws.max_row+1):
        metric=str(ws.cell(r,1).value or "").strip()
        year=_num(ws.cell(r,2).value)
        if metric not in {"Revenue","EPS","EBIT","EBITDA","FCF"} or year is None: continue
        consensus=_num(ws.cell(r,3).value); model=_num(ws.cell(r,4).value)
        rows.append({"FiscalYear":int(year),"Metric":metric,"Consensus":consensus,"Model":model})
    return rows


def ensure_forecast_accountability(wb,ticker,root=ROOT):
    ticker=str(ticker).upper(); now=datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    d=_research_dir(root,ticker); forecasts=_scenario_forecasts(wb); actuals=_historical_actuals(wb)
    snap_rows=[]
    for row in forecasts:
        snap_rows.append({"SnapshotAt":now,"Ticker":ticker,**row})
    cons=_consensus_rows(wb)
    for row in cons:
        snap_rows.append({"SnapshotAt":now,"Ticker":ticker,"FiscalYear":row["FiscalYear"],"Metric":f"Consensus {row['Metric']}","Forecast":row["Consensus"]})
        snap_rows.append({"SnapshotAt":now,"Ticker":ticker,"FiscalYear":row["FiscalYear"],"Metric":f"Model {row['Metric']}","Forecast":row["Model"]})
    path=d/"forecast_snapshots.csv"
    hist=pd.read_csv(path) if path.exists() else pd.DataFrame()
    if snap_rows:
        add=pd.DataFrame(snap_rows)
        hist=pd.concat([hist,add],ignore_index=True) if not hist.empty else add
        hist=hist.drop_duplicates(["SnapshotAt","Ticker","FiscalYear","Metric"],keep="last")
        hist.to_csv(path,index=False)

    eval_rows=[]
    if not hist.empty:
        hist["FiscalYear"]=pd.to_numeric(hist["FiscalYear"],errors="coerce")
        hist["Forecast"]=pd.to_numeric(hist["Forecast"],errors="coerce")
        for _,r in hist.dropna(subset=["FiscalYear","Forecast"]).iterrows():
            y=int(r["FiscalYear"]); metric=str(r["Metric"])
            actual_metric=metric.replace("Consensus ","").replace("Model ","")
            actual=actuals.get(y,{}).get(actual_metric)
            if actual is None and actual_metric=="FCF": actual=actuals.get(y,{}).get("FCF")
            if actual is None: continue
            err=float(actual-r["Forecast"])
            pct=err/abs(r["Forecast"]) if r["Forecast"] else None
            eval_rows.append({**r.to_dict(),"Actual":actual,"Error":err,"AbsError":abs(err),"PctError":pct,"DirectionCorrect":None})
    eval_df=pd.DataFrame(eval_rows)
    if not eval_df.empty:
        eval_df.to_csv(d/"forecast_evaluations.csv",index=False)
        summary=(eval_df.groupby("Metric",as_index=False)
                 .agg(Observations=("AbsError","count"),MeanAbsoluteError=("PctError",lambda s: pd.to_numeric(s,errors="coerce").abs().mean()),
                      Bias=("PctError","mean"),MeanAbsoluteValueError=("AbsError","mean")))
        summary["MatureFiscalYears"]=eval_df.groupby("Metric")["FiscalYear"].nunique().reindex(summary["Metric"]).values
        summary.to_csv(d/"forecast_accuracy_summary.csv",index=False)
    else:
        summary=pd.DataFrame(columns=["Metric","Observations","MeanAbsoluteError","Bias","MeanAbsoluteValueError","MatureFiscalYears"])

    ws=_prep_sheet(wb,"Forecast Accountability",f"{ticker} — Forecast Accountability",
                   "Point-in-time model/consensus snapshots are stored locally and compared with later reported actuals. This is an analyst track record, not a backfilled historical forecast.")
    _section(ws,5,"Current Forecast Snapshot")
    _header(ws,6,["Fiscal Year","Metric","Current Forecast","Actual (if reported)","Error","Error %","Snapshot Time","Status"])
    current=pd.DataFrame(snap_rows)
    rr=7
    for _,r in current.head(24).iterrows():
        y=int(r["FiscalYear"]); metric=str(r["Metric"]); actual_metric=metric.replace("Consensus ","").replace("Model ","")
        actual=actuals.get(y,{}).get(actual_metric)
        err=(actual-r["Forecast"]) if actual is not None and _num(r["Forecast"]) is not None else None
        pct=err/abs(r["Forecast"]) if err is not None and r["Forecast"] else None
        vals=[y,metric,_num(r["Forecast"]),actual,err,pct,r["SnapshotAt"],"MATURED" if actual is not None else "OPEN"]
        for c,v in enumerate(vals,1): ws.cell(rr,c,v)
        ws.cell(rr,6).number_format=FMT_PCT; rr+=1

    _section(ws,rr+1,"Historical Forecast Accuracy"); sr=rr+2
    _header(ws,sr,["Metric","Observations","Mean Absolute % Error","Bias","Mean Absolute Value Error","Mature Fiscal Years"])
    for i,(_,r) in enumerate(summary.iterrows(),sr+1):
        vals=[r.get("Metric"),r.get("Observations"),r.get("MeanAbsoluteError"),r.get("Bias"),r.get("MeanAbsoluteValueError"),r.get("MatureFiscalYears")]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        ws.cell(i,3).number_format=FMT_PCT; ws.cell(i,4).number_format=FMT_PCT
    for c,w in {"A":16,"B":24,"C":20,"D":20,"E":18,"F":18,"G":27,"H":16}.items(): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"
    return {"snapshots":len(snap_rows),"evaluations":len(eval_df),"summary_rows":len(summary)}


def ensure_earnings_revisions(wb,ticker,root=ROOT):
    ticker=str(ticker).upper(); d=_research_dir(root,ticker); now=datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    current=pd.DataFrame(_consensus_rows(wb))
    path=d/"estimate_revision_history.csv"
    old=pd.read_csv(path) if path.exists() else pd.DataFrame()
    prior={}
    if not old.empty:
        old["SnapshotAt"]=pd.to_datetime(old["SnapshotAt"],errors="coerce")
        for key,g in old.sort_values("SnapshotAt").groupby(["FiscalYear","Metric"]):
            prior[key]=g.iloc[-1]
    rows=[]
    for _,r in current.iterrows():
        key=(int(r["FiscalYear"]),str(r["Metric"])); p=prior.get(key)
        prev_cons=_num(p.get("Consensus")) if p is not None else None
        prev_model=_num(p.get("Model")) if p is not None else None
        rows.append({"SnapshotAt":now,"Ticker":ticker,"FiscalYear":key[0],"Metric":key[1],
                     "Consensus":_num(r["Consensus"]),"Model":_num(r["Model"]),
                     "PriorConsensus":prev_cons,"ConsensusRevision":(_num(r["Consensus"])/prev_cons-1) if prev_cons not in (None,0) and _num(r["Consensus"]) is not None else None,
                     "PriorModel":prev_model,"ModelRevision":(_num(r["Model"])/prev_model-1) if prev_model not in (None,0) and _num(r["Model"]) is not None else None})
    if rows:
        _append_snapshot(path,rows[0],("SnapshotAt","FiscalYear","Metric"))
        out=pd.read_csv(path)
        out=out.iloc[:-1] if len(out) else out
        out=pd.concat([out,pd.DataFrame(rows)],ignore_index=True)
        out.to_csv(path,index=False)

    ws=_prep_sheet(wb,"Earnings & Revisions",f"{ticker} — Earnings & Estimate Revisions",
                   "Tracks model-vs-consensus gaps and revisions between local research runs. Actual surprise is added once the fiscal period appears in Historical Financials.")
    _section(ws,5,"Current Estimate Revision Monitor")
    _header(ws,6,["Fiscal Year","Metric","Prior Consensus","Current Consensus","Consensus Revision","Prior Model","Current Model","Model Revision","Model vs Consensus"])
    actuals=_historical_actuals(wb)
    for i,r in enumerate(rows,7):
        gap=(r["Model"]/r["Consensus"]-1) if r["Consensus"] not in (None,0) and r["Model"] is not None else None
        vals=[r["FiscalYear"],r["Metric"],r["PriorConsensus"],r["Consensus"],r["ConsensusRevision"],r["PriorModel"],r["Model"],r["ModelRevision"],gap]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        for c in (5,8,9): ws.cell(i,c).number_format=FMT_PCT
    _section(ws,max(9,7+len(rows)+2),"Latest Reported Surprise / Accountability")
    start=max(10,7+len(rows)+3); _header(ws,start,["Fiscal Year","Metric","Actual","Latest Prior Consensus","Surprise %","Status"])
    j=start+1
    for (y,metric),p in sorted(prior.items()):
        actual=actuals.get(int(y),{}).get(metric)
        cons=_num(p.get("Consensus"))
        if actual is None or cons in (None,0): continue
        vals=[int(y),metric,actual,cons,actual/cons-1,"MATURED"]
        for c,v in enumerate(vals,1): ws.cell(j,c,v)
        ws.cell(j,5).number_format=FMT_PCT; j+=1
    for c,w in {"A":15,"B":20,"C":18,"D":18,"E":18,"F":18,"G":18,"H":18,"I":18}.items(): ws.column_dimensions[c].width=w
    return {"revision_rows":len(rows)}


def _statement_series(wb, labels):
    out={}
    # Prefer canonical Financial Statements; fallback to Historical Financials.
    for sheet in ["Financial Statements","Historical Financials"]:
        if sheet not in wb.sheetnames: continue
        ws=wb[sheet]
        for r in range(1,ws.max_row+1):
            label=str(ws.cell(r,1).value or "").strip().lower()
            if label not in {str(x).lower() for x in labels}: continue
            # find closest year header above
            for hr in range(max(1,r-5),r):
                years=_year_cols(ws,hr,20)
                if years:
                    for y,c in years.items():
                        v=_num(ws.cell(r,c).value)
                        if v is not None: out[y]=v
                    if out: return out
    return out


def ensure_capital_allocation(wb,ticker,root=ROOT):
    ticker=str(ticker).upper()
    ocf=_statement_series(wb,["Operating Cash Flow","Net Cash Provided by Operating Activities","Cash From Operations"])
    capex=_statement_series(wb,["Capital Expenditures","Purchases of Property Plant and Equipment","Capex"])
    buybacks=_statement_series(wb,["Repurchase of Stock","Repurchases of Common Stock","Common Stock Repurchased"])
    dividends=_statement_series(wb,["Dividends Paid","Common Dividends Paid"])
    debt=_statement_series(wb,["Total Debt"])
    cash=_statement_series(wb,["Cash And Cash Equivalents","Cash & Equivalents"])
    shares=_statement_series(wb,["Diluted Shares","Diluted Weighted Average Shares"])
    ni=_statement_series(wb,["Net Income","Net Income Attributable to Common"])
    years=sorted(set(ocf)|set(capex)|set(buybacks)|set(dividends)|set(debt)|set(cash)|set(shares)|set(ni))
    rows=[]
    for y in years:
        fcf=ocf.get(y)-abs(capex.get(y,0)) if y in ocf and y in capex else None
        net_share=None
        if y in shares:
            py=max([z for z in shares if z<y],default=None)
            if py and shares.get(py) not in (None,0): net_share=1-shares[y]/shares[py]
        rows.append({"Year":y,"OCF":ocf.get(y),"Capex":capex.get(y),"FCF":fcf,"Buybacks":buybacks.get(y),"Dividends":dividends.get(y),
                     "Debt":debt.get(y),"Cash":cash.get(y),"NetIncome":ni.get(y),"DilutedShares":shares.get(y),"NetShareReduction":net_share})
    df=pd.DataFrame(rows)
    d=_research_dir(root,ticker)
    if not df.empty: df.to_csv(d/"capital_allocation_history.csv",index=False)
    # Incremental return on capital proxy: change in NOPAT / change in invested capital.
    incr=None
    if len(df)>=3:
        first,last=df.iloc[0],df.iloc[-1]
        op_series=_statement_series(wb,["Operating Income","Operating Income / (Loss)"])
        tax=_num(_label_value(wb,"Cost of Capital","Effective tax rate"),.20)
        if first["Year"] in op_series and last["Year"] in op_series:
            dnopat=op_series[last["Year"]]*(1-tax)-op_series[first["Year"]]*(1-tax)
            cap0=(first.get("Debt") or 0)+(0)-(first.get("Cash") or 0)
            cap1=(last.get("Debt") or 0)+(0)-(last.get("Cash") or 0)
            dcap=cap1-cap0
            if abs(dcap)>1e-9: incr=dnopat/dcap
    ws=_prep_sheet(wb,"Capital Allocation",f"{ticker} — Capital Allocation",
                   "Tracks how operating cash is reinvested or returned and surfaces incremental-return diagnostics. Missing cash-flow uses remain blank rather than estimated.")
    _section(ws,5,"Historical Capital Allocation")
    _header(ws,6,["Year","Operating CF","Capex","Free Cash Flow","Buybacks","Dividends","Debt","Cash","Diluted Shares","Net Share Reduction"])
    for i,(_,r) in enumerate(df.iterrows(),7):
        vals=[r["Year"],r["OCF"],r["Capex"],r["FCF"],r["Buybacks"],r["Dividends"],r["Debt"],r["Cash"],r["DilutedShares"],r["NetShareReduction"]]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        for c in range(2,10): ws.cell(i,c).number_format=FMT_BN
        ws.cell(i,10).number_format=FMT_PCT
    sr=max(9,7+len(df)+2); _section(ws,sr,"Allocation Diagnostics")
    diagnostics=[
        ("Latest FCF",_num(df.iloc[-1]["FCF"]) if not df.empty else None,FMT_BN),
        ("Latest net share reduction",_num(df.iloc[-1]["NetShareReduction"]) if not df.empty else None,FMT_PCT),
        ("Incremental ROIC proxy",incr,FMT_PCT),
        ("ROIC minus WACC",None,FMT_PCT),
    ]
    roic=_num(_label_value(wb,"Dashboard","ROIC"),None); wacc=_num(_label_value(wb,"Cost of Capital","Base",cols=(1,)),None)
    if wacc is None: wacc=_num(_label_value(wb,"Decision View","Calculated Base WACC"),None)
    if roic is not None and wacc is not None: diagnostics[-1]=("ROIC minus WACC",roic-wacc,FMT_PCT)
    for j,(lab,val,fmt) in enumerate(diagnostics,sr+1):
        ws.cell(j,1,lab); ws.cell(j,2,val); ws.cell(j,2).number_format=fmt
        ws.cell(j,3,"PASS" if val is not None else "REVIEW")
    return {"years":len(df),"incremental_roic":incr}


def _current_peer_metrics(wb):
    out={}
    if "Peer Comps" not in wb.sheetnames: return out
    ws=wb["Peer Comps"]; headers={str(ws.cell(3,c).value or "").strip():c for c in range(1,ws.max_column+1)}
    target=None
    for r in range(4,min(ws.max_row,50)+1):
        if str(ws.cell(r,headers.get("Peer Type",16)).value or "")=="Target classification": target=r; break
    target=target or 4
    for key in ["Forward P/E","EV/Revenue","EV/EBITDA","Revenue Growth","Operating Margin","ROE"]:
        c=headers.get(key)
        if c: out[key]=_num(ws.cell(target,c).value)
    return out


def ensure_valuation_history(wb,ticker,root=ROOT):
    ticker=str(ticker).upper(); d=_research_dir(root,ticker); now=datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    metrics=_current_peer_metrics(wb)
    price=_num(_label_value(wb,"Decision View","Current Market Price"))
    base=_num(_label_value(wb,"Decision View","Base DCF Fair Value"))
    implied=_num(_label_value(wb,"Decision View","Market-Implied 10Y FCF CAGR"))
    row={"SnapshotAt":now,"Ticker":ticker,"Price":price,"BaseValue":base,"BaseUpside":base/price-1 if price and base else None,"ImpliedFCFCAGR":implied,**metrics}
    hist=_append_snapshot(d/"valuation_history.csv",row,("SnapshotAt","Ticker"))
    ws=_prep_sheet(wb,"Valuation History",f"{ticker} — Valuation History",
                   "Local point-in-time valuation snapshots accumulated across model runs. Percentiles are based only on observations actually captured by this project.")
    _section(ws,5,"Current Valuation vs Captured History")
    _header(ws,6,["Metric","Current","Historical Median","Percentile","Observations","Status"])
    rr=7
    for metric in ["Forward P/E","EV/Revenue","EV/EBITDA","BaseUpside","ImpliedFCFCAGR"]:
        cur=_num(row.get(metric)); s=pd.to_numeric(hist.get(metric,pd.Series(dtype=float)),errors="coerce").dropna()
        median=float(s.median()) if len(s) else None
        pct=float((s<=cur).mean()) if cur is not None and len(s)>=2 else None
        vals=[metric,cur,median,pct,int(len(s)),"PASS" if len(s)>=5 else "REVIEW"]
        for c,v in enumerate(vals,1): ws.cell(rr,c,v)
        if metric in {"BaseUpside","ImpliedFCFCAGR"}:
            ws.cell(rr,2).number_format=FMT_PCT; ws.cell(rr,3).number_format=FMT_PCT
        else:
            ws.cell(rr,2).number_format=FMT_MULT; ws.cell(rr,3).number_format=FMT_MULT
        ws.cell(rr,4).number_format=FMT_PCT; rr+=1
    _section(ws,rr+1,"Captured Snapshot History")
    _header(ws,rr+2,["Snapshot","Price","Forward P/E","EV/Revenue","EV/EBITDA","Base Upside","Implied FCF CAGR"])
    for i,(_,r) in enumerate(hist.tail(30).iterrows(),rr+3):
        vals=[r.get("SnapshotAt"),r.get("Price"),r.get("Forward P/E"),r.get("EV/Revenue"),r.get("EV/EBITDA"),r.get("BaseUpside"),r.get("ImpliedFCFCAGR")]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        ws.cell(i,2).number_format=FMT_PRICE
        for c in (3,4,5): ws.cell(i,c).number_format=FMT_MULT
        for c in (6,7): ws.cell(i,c).number_format=FMT_PCT
    return {"observations":len(hist)}


def _segment_numeric_rows(wb):
    if "Segment Analysis" not in wb.sheetnames: return []
    ws=wb["Segment Analysis"]; rows=[]
    # Use latest numeric value among early annual columns; never infer missing segment numbers.
    for r in range(1,ws.max_row+1):
        name=str(ws.cell(r,1).value or "").strip()
        if not name or name.lower() in {"segment","business line / revenue group","business / segment"}: continue
        vals=[]
        for c in range(2,min(ws.max_column,10)+1):
            v=_num(ws.cell(r,c).value)
            if v is not None: vals.append(v)
        if vals and max(abs(x) for x in vals)>0: rows.append((name,vals[-1]))
    # De-duplicate obvious totals/header rows.
    clean=[]; seen=set()
    for name,v in rows:
        key=name.lower()
        if key in seen or any(x in key for x in ["total","growth","margin","mix","share"]): continue
        seen.add(key); clean.append((name,v))
    return clean[:12]


def ensure_sotp(wb,ticker):
    ticker=str(ticker).upper(); seg=_segment_numeric_rows(wb); pm=_current_peer_metrics(wb)
    multiple=_num(pm.get("EV/Revenue"))
    price=_num(_label_value(wb,"Decision View","Current Market Price")); shares=_num(_label_value(wb,"Company Data","Shares Outstanding"))
    cash=_num(_label_value(wb,"Company Data","Cash")); debt=_num(_label_value(wb,"Company Data","Debt"))
    ws=_prep_sheet(wb,"SOTP",f"{ticker} — Sum-of-the-Parts Diagnostic",
                   "Gated diagnostic only. It is populated only when multiple numeric segment revenues exist; a single company-level EV/Revenue multiple is shown as a placeholder reference, not as a production segment valuation.")
    _section(ws,5,"Segment Valuation Diagnostic")
    _header(ws,6,["Segment","Latest Numeric Revenue","Reference EV/Revenue","Implied EV","Status","Caveat"])
    status="REVIEW"
    total=0.0
    if len(seg)>=2 and multiple is not None:
        for i,(name,rev) in enumerate(seg,7):
            ev=rev*multiple; total+=ev
            vals=[name,rev,multiple,ev,"REVIEW","Uses company-level target EV/Revenue as a transparent placeholder; replace with segment-specific peer multiple before decision use."]
            for c,v in enumerate(vals,1): ws.cell(i,c,v)
            ws.cell(i,2).number_format=FMT_BN; ws.cell(i,3).number_format=FMT_MULT; ws.cell(i,4).number_format=FMT_BN
        status="REVIEW — SEGMENT-SPECIFIC MULTIPLES REQUIRED"
    else:
        ws["A7"]="N/M"; ws["B7"]="Insufficient reliable numeric segment revenue or valuation multiple."; ws.merge_cells("B7:F7")
    sr=max(10,8+len(seg)); _section(ws,sr,"SOTP Bridge")
    equity=total+(cash or 0)-(debt or 0) if total else None
    value=equity/shares if equity is not None and shares not in (None,0) else None
    rows=[("Segment enterprise value",total if total else None,FMT_BN),("Net cash / (debt)",(cash or 0)-(debt or 0) if cash is not None or debt is not None else None,FMT_BN),("Indicative equity value",equity,FMT_BN),("Indicative value / share",value,FMT_PRICE),("Current price",price,FMT_PRICE),("Status",status,"General")]
    for i,(lab,val,fmt) in enumerate(rows,sr+1): ws.cell(i,1,lab); ws.cell(i,2,val); ws.cell(i,2).number_format=fmt
    return {"segments":len(seg),"status":status,"value_per_share":value}


def ensure_catalyst_timeline(wb,ticker,info=None,root=ROOT):
    ticker=str(ticker).upper(); info=info or {}; d=_research_dir(root,ticker); now=datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    events=[]
    earnings=info.get("earningsTimestamp") or info.get("earningsTimestampStart")
    if earnings:
        try:
            dt=pd.to_datetime(float(earnings),unit="s",utc=True).isoformat()
            events.append({"EventDate":dt,"EventType":"Earnings","Description":"Next/indicated earnings event from current market-data snapshot","ExpectedImpact":"Review thesis vs reported KPIs","Source":"market-data snapshot","Status":"OPEN"})
        except Exception: pass
    # Add current model-state checkpoints.
    view=_label_value(wb,"Decision View","MODEL VIEW")
    base_up=_num(_label_value(wb,"Decision View","Base DCF Upside"))
    dq_fail=0
    if "Data Quality" in wb.sheetnames:
        for r in range(1,wb["Data Quality"].max_row+1):
            dq_fail += str(wb["Data Quality"].cell(r,2).value or "").upper().strip()=="FAIL"
    events.append({"EventDate":now,"EventType":"Model Run","Description":str(view or "Model refreshed"),"ExpectedImpact":f"Base DCF upside {base_up:.1%}" if base_up is not None else "Review current valuation evidence","Source":"Decision View","Status":"FAIL" if dq_fail else "PASS"})
    path=d/"catalyst_timeline.csv"
    hist=pd.read_csv(path) if path.exists() else pd.DataFrame()
    if events:
        add=pd.DataFrame(events)
        hist=pd.concat([hist,add],ignore_index=True) if not hist.empty else add
        hist=hist.drop_duplicates(["EventDate","EventType","Description"],keep="last")
        hist.to_csv(path,index=False)
    ws=_prep_sheet(wb,"Catalyst Timeline",f"{ticker} — Catalyst & Thesis-Change Timeline",
                   "Chronological monitoring surface. Automatic events are evidence-linked; discretionary thesis events can be appended to research_data/<TICKER>/catalyst_timeline.csv.")
    _section(ws,5,"Catalysts & Thesis Changes")
    _header(ws,6,["Event Date","Type","Description","Expected / Thesis Impact","Source","Status"])
    for i,(_,r) in enumerate(hist.sort_values("EventDate",ascending=False).head(40).iterrows(),7):
        vals=[r.get("EventDate"),r.get("EventType"),r.get("Description"),r.get("ExpectedImpact"),r.get("Source"),r.get("Status")]
        for c,v in enumerate(vals,1): ws.cell(i,c,v); ws.cell(i,c).alignment=Alignment(wrap_text=True,vertical="top")
    return {"events":len(hist)}


def apply_offline_research_extensions(wb,ticker,info=None,root=ROOT):
    """Apply all private/offline accountability layers. Never called by public showcase code."""
    results={}
    steps=[
        ("forecast_accountability",lambda:ensure_forecast_accountability(wb,ticker,root)),
        ("earnings_revisions",lambda:ensure_earnings_revisions(wb,ticker,root)),
        ("capital_allocation",lambda:ensure_capital_allocation(wb,ticker,root)),
        ("valuation_history",lambda:ensure_valuation_history(wb,ticker,root)),
        ("sotp",lambda:ensure_sotp(wb,ticker)),
        ("catalyst_timeline",lambda:ensure_catalyst_timeline(wb,ticker,info,root)),
    ]
    for name,fn in steps:
        try: results[name]={"status":"PASS",**(fn() or {})}
        except Exception as exc: results[name]={"status":"REVIEW","error":str(exc)}
    return results
