from __future__ import annotations

"""Private/offline analyst-accountability and deeper fundamental research extensions."""

from datetime import datetime, timezone
from pathlib import Path
import csv
import json
import math
import re
import statistics

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from business_model_registry import workbook_policy

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREY="666666"
GREEN="E2F0D9"; GOLD="FFF2CC"; RED="FCE4D6"; INPUT="FFF2CC"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_MULT='0.0x;[Red](0.0x);-'
THIN=Side(style="thin",color="D9E1F2")


def _fill(color): return PatternFill("solid",fgColor=color)


def _num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        x=float(v); return x if math.isfinite(x) else default
    except Exception: return default


def _find(ws,label,col=1):
    needle=str(label).strip().lower()
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,col).value or "").strip().lower()==needle: return r
    return None


def _find_contains(ws,needles,col=1):
    needles=[str(x).lower() for x in needles]
    for r in range(1,ws.max_row+1):
        text=str(ws.cell(r,col).value or "").strip().lower()
        if text and any(n in text for n in needles): return r
    return None


def _new_sheet(wb,name):
    if name in wb.sheetnames: wb.remove(wb[name])
    ws=wb.create_sheet(name); ws.sheet_view.showGridLines=False
    try: ws.sheet_view.zoomScale=90
    except Exception: pass
    return ws


def _title(ws,text,note=None,end=10):
    for c in range(1,end+1):
        ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=text; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    if note:
        ws["A3"]=note; ws["A3"].font=Font(italic=True,color=GREY,size=9)
        ws["A3"].alignment=Alignment(wrap_text=True,vertical="top")
        ws.merge_cells(start_row=3,start_column=1,end_row=3,end_column=end)


def _section(ws,row,text,end=10):
    for c in range(1,end+1):
        ws.cell(row,c).fill=_fill(NAVY); ws.cell(row,c).font=Font(bold=True,color=WHITE)
    ws.cell(row,1,text)


def _header(ws,row,headers):
    for c,v in enumerate(headers,1):
        cell=ws.cell(row,c,v); cell.fill=_fill(BLUE); cell.font=Font(bold=True,color=WHITE)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.border=Border(bottom=THIN)


def _status_fill(status): return _fill(GREEN if status=="PASS" else RED if status=="FAIL" else GOLD)


def _utc_iso(): return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _historical_actuals(wb):
    if "Historical Financials" not in wb.sheetnames: return {}
    ws=wb["Historical Financials"]; out={}
    for c in range(2,8):
        year=_num(ws.cell(3,c).value)
        if year is None: continue
        year=int(year); revenue=_num(ws.cell(4,c).value); op=_num(ws.cell(9,c).value)
        ni=_num(ws.cell(11,c).value); eps=_num(ws.cell(12,c).value)
        ocf=_num(ws.cell(14,c).value); capex=_num(ws.cell(15,c).value)
        depr=_num(ws.cell(18,c).value); sbc=_num(ws.cell(21,c).value)
        fcf=ocf-abs(capex) if ocf is not None and capex is not None else None
        out[year]={
            "Revenue":revenue,"EBIT Margin":op/revenue if op is not None and revenue not in (None,0) else None,
            "EPS":eps,"FCF":fcf,"OperatingIncome":op,"NetIncome":ni,"OCF":ocf,
            "Capex":abs(capex) if capex is not None else None,
            "Depreciation":abs(depr) if depr is not None else None,
            "SBC":abs(sbc) if sbc is not None else None,
            "Shares":ni/eps if ni is not None and eps not in (None,0) else None,
        }
    return out


def _consensus_rows(wb):
    if "Expectations & Consensus" not in wb.sheetnames: return []
    ws=wb["Expectations & Consensus"]; rows=[]
    for r in range(1,min(ws.max_row,120)+1):
        metric=str(ws.cell(r,1).value or "").strip(); year=_num(ws.cell(r,2).value)
        if metric not in {"Revenue","EPS","EBIT Margin","FCF","Capex / Revenue"} or year is None: continue
        rows.append({
            "Metric":metric,"FiscalYear":int(year),"Consensus":_num(ws.cell(r,3).value),
            "Model":_num(ws.cell(r,4).value),"Revision30D":_num(ws.cell(r,7).value),
            "Revision90D":_num(ws.cell(r,8).value),"Dispersion":_num(ws.cell(r,9).value),
            "ProviderCount":_num(ws.cell(r,12).value),"Providers":ws.cell(r,13).value,
        })
    return rows


def _earnings_surprises(wb):
    if "Advanced Analytics" not in wb.sheetnames: return []
    ws=wb["Advanced Analytics"]; rows=[]
    for r in range(1,min(ws.max_row,35)+1):
        reported=ws.cell(r,9).value; est=_num(ws.cell(r,10).value)
        actual=_num(ws.cell(r,11).value); surprise=_num(ws.cell(r,12).value)
        if reported and str(reported).strip().lower()!="reported" and any(x is not None for x in (est,actual,surprise)):
            rows.append({"Reported":reported,"Estimate":est,"Actual":actual,"Surprise":surprise})
    return rows[-12:]


def _snapshot(wb,ticker,captured_at=None):
    actuals=_historical_actuals(wb)
    return {
        "ticker":str(ticker).upper(),"captured_at":captured_at or _utc_iso(),
        "latest_actual_year":max(actuals) if actuals else None,
        "estimates":_consensus_rows(wb),"earnings_surprises":_earnings_surprises(wb),
    }


def _history_path(root,ticker): return Path(root)/str(ticker).upper()/"forecast_history.json"


def _load_history(root,ticker):
    path=_history_path(root,ticker)
    if not path.exists(): return {"ticker":str(ticker).upper(),"snapshots":[]}
    try: data=json.loads(path.read_text(encoding="utf-8"))
    except Exception: data={"ticker":str(ticker).upper(),"snapshots":[]}
    if not isinstance(data,dict): data={"ticker":str(ticker).upper(),"snapshots":[]}
    data.setdefault("ticker",str(ticker).upper()); data.setdefault("snapshots",[])
    return data


def _snapshot_key(s):
    return sorted((r.get("Metric"),r.get("FiscalYear"),r.get("Consensus"),r.get("Model")) for r in s.get("estimates",[]))


def persist_forecast_snapshot(root,ticker,payload):
    data=_load_history(root,ticker); snaps=data["snapshots"]
    changed=not snaps or _snapshot_key(snaps[-1])!=_snapshot_key(payload)
    path=_history_path(root,ticker); path.parent.mkdir(parents=True,exist_ok=True)
    if changed:
        snaps.append(payload)
        path.write_text(json.dumps(data,indent=2,ensure_ascii=False,default=str)+"\n",encoding="utf-8")
    return path,changed


def _prior_estimate(history,metric,year,current_capture):
    rows=[]
    for snap in history.get("snapshots",[]):
        if snap.get("captured_at")==current_capture: continue
        for item in snap.get("estimates",[]):
            if item.get("Metric")==metric and int(item.get("FiscalYear") or 0)==int(year):
                rows.append((snap.get("captured_at") or "",item))
    return sorted(rows,key=lambda x:x[0])[-1][1] if rows else None


def forecast_accuracy_records(history,actuals):
    rows=[]
    for snap in history.get("snapshots",[]):
        for est in snap.get("estimates",[]):
            metric=est.get("Metric"); year=int(est.get("FiscalYear") or 0)
            if metric not in {"Revenue","EPS","FCF","EBIT Margin"} or year not in actuals: continue
            actual=_num(actuals[year].get(metric)); previous=_num(actuals.get(year-1,{}).get(metric))
            if actual is None: continue
            for kind,key in (("Model","Model"),("Consensus","Consensus")):
                forecast=_num(est.get(key))
                if forecast is None: continue
                error=(forecast-actual) if metric=="EBIT Margin" else ((forecast/actual-1) if actual not in (None,0) else None)
                if error is None: continue
                fchg=forecast-previous if previous is not None else None
                achg=actual-previous if previous is not None else None
                direction=(fchg>0)==(achg>0) if fchg not in (None,0) and achg not in (None,0) else None
                rows.append({
                    "CapturedAt":snap.get("captured_at"),"FiscalYear":year,"Metric":metric,
                    "ForecastType":kind,"Forecast":forecast,"Actual":actual,"Error":error,
                    "AbsoluteError":abs(error),"DirectionCorrect":direction,
                })
    return rows


def forecast_accuracy_summary(records):
    groups={}
    for r in records: groups.setdefault((r["Metric"],r["ForecastType"]),[]).append(r)
    out=[]
    for (metric,kind),rows in groups.items():
        errs=[x["AbsoluteError"] for x in rows]; signed=[x["Error"] for x in rows]
        dirs=[x["DirectionCorrect"] for x in rows if isinstance(x["DirectionCorrect"],bool)]
        if not errs: continue
        out.append({
            "Metric":metric,"ForecastType":kind,"Observations":len(errs),
            "MeanAbsoluteError":sum(errs)/len(errs),"MedianAbsoluteError":statistics.median(errs),
            "Bias":sum(signed)/len(signed),"DirectionAccuracy":sum(dirs)/len(dirs) if dirs else None,
        })
    return out


def write_accuracy_csv(root,ticker,summary):
    path=Path(root)/str(ticker).upper()/"forecast_accuracy_summary.csv"; path.parent.mkdir(parents=True,exist_ok=True)
    fields=["Metric","ForecastType","Observations","MeanAbsoluteError","MedianAbsoluteError","Bias","DirectionAccuracy"]
    with path.open("w",newline="",encoding="utf-8") as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader()
        for row in summary: w.writerow({k:row.get(k) for k in fields})
    return path


def ensure_forecast_accountability(wb,ticker,history,actuals,current):
    ws=_new_sheet(wb,"Forecast Accountability")
    _title(ws,f"{ticker} — Forecast Accountability",
           "Stored forecasts are graded only after the corresponding fiscal-year actual exists. Model and consensus errors remain separate so analyst/model calibration can be compared.",9)
    _section(ws,5,"Current Forward Forecast Snapshot",9)
    _header(ws,6,["Fiscal Year","Metric","Model Forecast","Consensus","Model vs Consensus","30D Revision","90D Revision","Dispersion","Providers"])
    rr=7
    for row in current.get("estimates",[]):
        model=_num(row.get("Model")); cons=_num(row.get("Consensus")); margin=row["Metric"] in {"EBIT Margin","Capex / Revenue"}
        gap=(model-cons) if margin and model is not None and cons is not None else ((model/cons-1) if model is not None and cons not in (None,0) else None)
        vals=[row["FiscalYear"],row["Metric"],model,cons,gap,row.get("Revision30D"),row.get("Revision90D"),row.get("Dispersion"),row.get("Providers")]
        for c,v in enumerate(vals,1): ws.cell(rr,c,v)
        fmt=FMT_PCT if margin else FMT_PRICE if row["Metric"]=="EPS" else FMT_BN
        ws.cell(rr,3).number_format=fmt; ws.cell(rr,4).number_format=fmt
        for c in (5,6,7,8): ws.cell(rr,c).number_format=FMT_PCT
        rr+=1
    records=forecast_accuracy_records(history,actuals); summary=forecast_accuracy_summary(records)
    start=max(rr+2,18); _section(ws,start,"Matured Forecast Scorecard",9)
    _header(ws,start+1,["Metric","Forecast Type","Observations","Mean Absolute Error","Median Absolute Error","Bias","Direction Accuracy","Status","Interpretation"])
    r=start+2
    for row in summary:
        status="PASS" if row["Observations"]>=3 else "REVIEW"
        vals=[row["Metric"],row["ForecastType"],row["Observations"],row["MeanAbsoluteError"],row["MedianAbsoluteError"],row["Bias"],row["DirectionAccuracy"],status,"More observations improve calibration confidence."]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        for c in (4,5,6,7): ws.cell(r,c).number_format=FMT_PCT
        ws.cell(r,8).fill=_status_fill(status); ws.cell(r,8).font=Font(bold=True); r+=1
    if not summary:
        ws.cell(r,1,"REVIEW"); ws.cell(r,2,"No matured stored forecast yet. This becomes meaningful after a forecasted fiscal year reports."); ws.cell(r,1).fill=_fill(GOLD)
    d=r+3; _section(ws,d,"Forecast-vs-Actual Detail",9)
    _header(ws,d+1,["Captured At","Fiscal Year","Metric","Forecast Type","Forecast","Actual","Error","Absolute Error","Direction Correct"])
    for i,row in enumerate(records[-100:],d+2):
        vals=[row["CapturedAt"],row["FiscalYear"],row["Metric"],row["ForecastType"],row["Forecast"],row["Actual"],row["Error"],row["AbsoluteError"],row["DirectionCorrect"]]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        ws.cell(i,7).number_format=FMT_PCT; ws.cell(i,8).number_format=FMT_PCT
    for c,w in {"A":24,"B":13,"C":20,"D":16,"E":16,"F":16,"G":16,"H":18,"I":45}.items(): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"; return {"records":records,"summary":summary}


def ensure_earnings_revisions(wb,ticker,history,current):
    ws=_new_sheet(wb,"Earnings & Revisions")
    _title(ws,f"{ticker} — Earnings, Estimate Revisions & Thesis Read-Through",
           "Compares the current stored consensus/model snapshot with the prior local snapshot. Revision signals are evidence only and never change DCF assumptions automatically.",10)
    _section(ws,5,"Estimate Revision Monitor",10)
    _header(ws,6,["Fiscal Year","Metric","Previous Consensus","Current Consensus","Consensus Revision","Previous Model","Current Model","Model Revision","Dispersion","Evidence Signal"])
    r=7; capture=current.get("captured_at")
    for row in current.get("estimates",[]):
        prev=_prior_estimate(history,row["Metric"],row["FiscalYear"],capture) or {}
        pc=_num(prev.get("Consensus")); cc=_num(row.get("Consensus")); pm=_num(prev.get("Model")); cm=_num(row.get("Model"))
        margin=row["Metric"] in {"EBIT Margin","Capex / Revenue"}
        cr=(cc-pc) if margin and cc is not None and pc is not None else ((cc/pc-1) if cc is not None and pc not in (None,0) else None)
        mr=(cm-pm) if margin and cm is not None and pm is not None else ((cm/pm-1) if cm is not None and pm not in (None,0) else None)
        signal="POSITIVE" if cr is not None and cr>.02 else "NEGATIVE" if cr is not None and cr<-.02 else "MIXED / UNCHANGED" if cr is not None else "NO PRIOR SNAPSHOT"
        vals=[row["FiscalYear"],row["Metric"],pc,cc,cr,pm,cm,mr,row.get("Dispersion"),signal]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        fmt=FMT_PCT if margin else FMT_PRICE if row["Metric"]=="EPS" else FMT_BN
        for c in (3,4,6,7): ws.cell(r,c).number_format=fmt
        for c in (5,8,9): ws.cell(r,c).number_format=FMT_PCT
        r+=1
    start=max(r+2,18); _section(ws,start,"Recent EPS Surprise Context",10)
    _header(ws,start+1,["Reported","EPS Estimate","Actual EPS","EPS Surprise","Signal","Use","Caveat"])
    er=current.get("earnings_surprises") or []
    if er:
        for i,row in enumerate(er,start+2):
            surprise=_num(row.get("Surprise")); signal="POSITIVE" if surprise is not None and surprise>.03 else "NEGATIVE" if surprise is not None and surprise<-.03 else "MIXED"
            vals=[row.get("Reported"),row.get("Estimate"),row.get("Actual"),surprise,signal,"Compare with revisions and thesis KPIs","One quarter does not prove a long-term thesis."]
            for c,v in enumerate(vals,1): ws.cell(i,c,v)
            ws.cell(i,2).number_format=FMT_PRICE; ws.cell(i,3).number_format=FMT_PRICE; ws.cell(i,4).number_format=FMT_PCT
    else:
        ws.cell(start+2,1,"REVIEW"); ws.cell(start+2,2,"No provider earnings-surprise history available.")
    for c,w in {"A":18,"B":16,"C":18,"D":18,"E":18,"F":18,"G":18,"H":18,"I":18,"J":32}.items(): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"


def _year_headers(ws):
    out=[]
    for r in range(1,min(ws.max_row,60)+1):
        m={}
        for c in range(2,min(ws.max_column,16)+1):
            v=_num(ws.cell(r,c).value)
            if v is not None and 1900<=int(v)<=2100 and float(v)==int(v): m[int(v)]=c
        if len(m)>=2: out.append((r,m))
    return out


def _fs_value(wb,needles,year):
    if "Financial Statements" not in wb.sheetnames: return None
    ws=wb["Financial Statements"]; row=_find_contains(ws,needles,1)
    if not row:return None
    candidates=[x for x in _year_headers(ws) if x[0]<row and year in x[1]]
    if not candidates:return None
    _,m=max(candidates,key=lambda x:x[0]); return _num(ws.cell(row,m[year]).value)


def _tax_rate(wb):
    if "Cost of Capital" in wb.sheetnames:
        ws=wb["Cost of Capital"]; r=_find_contains(ws,["effective tax rate"])
        if r:
            v=_num(ws.cell(r,2).value)
            if v is not None and 0<=v<.60:return v
    return .21


def ensure_capital_allocation(wb,ticker,info=None):
    actuals=_historical_actuals(wb); ws=_new_sheet(wb,"Capital Allocation")
    _title(ws,f"{ticker} — Capital Allocation & Incremental Returns",
           "Tracks internally generated cash deployment and estimates ROIC/incremental ROIC only when comparable statement rows exist. Missing buyback/M&A/debt data stays blank.",12)
    _section(ws,5,"Historical Cash Deployment",12)
    _header(ws,6,["Year","OCF","Capex","FCF","SBC","Buybacks","Dividends","Acquisitions","Debt Repayment","Share Reduction","NOPAT","Invested Capital"])
    tax=_tax_rate(wb); rows=[]
    for year in sorted(actuals):
        a=actuals[year]; buy=_fs_value(wb,["repurchase","common stock repurchased","payments for repurchase"],year)
        div=_fs_value(wb,["dividends paid","payments of dividends","common dividends"],year)
        acq=_fs_value(wb,["acquisition","business acquisitions"],year); debtpay=_fs_value(wb,["repayments of debt","debt repayments"],year)
        debt=_fs_value(wb,["total debt"],year); equity=_fs_value(wb,["total equity","stockholders' equity"],year)
        cash=_fs_value(wb,["cash and cash equivalents","cash & cash equivalents"],year)
        invested=debt+equity-cash if None not in (debt,equity,cash) else None
        nopat=a["OperatingIncome"]*(1-tax) if a["OperatingIncome"] is not None else None
        prev=actuals.get(year-1,{}).get("Shares"); shares=a.get("Shares")
        reduction=prev/shares-1 if prev not in (None,0) and shares not in (None,0) else None
        rows.append({"Year":year,"OCF":a["OCF"],"Capex":a["Capex"],"FCF":a["FCF"],"SBC":a["SBC"],
            "Buybacks":abs(buy) if buy is not None else None,"Dividends":abs(div) if div is not None else None,
            "Acquisitions":abs(acq) if acq is not None else None,"DebtRepayment":abs(debtpay) if debtpay is not None else None,
            "ShareReduction":reduction,"NOPAT":nopat,"InvestedCapital":invested})
    for r,row in enumerate(rows,7):
        vals=[row["Year"],row["OCF"],row["Capex"],row["FCF"],row["SBC"],row["Buybacks"],row["Dividends"],row["Acquisitions"],row["DebtRepayment"],row["ShareReduction"],row["NOPAT"],row["InvestedCapital"]]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        for c in (2,3,4,5,6,7,8,9,11,12): ws.cell(r,c).number_format=FMT_BN
        ws.cell(r,10).number_format=FMT_PCT
    start=max(7+len(rows)+2,16); _section(ws,start,"Capital-Efficiency Diagnostics",12)
    _header(ws,start+1,["Metric","Value","Status","Interpretation","Method / Caveat"])
    valid=[x for x in rows if x.get("InvestedCapital") is not None and x.get("NOPAT") is not None]
    first=valid[0] if valid else None; last=valid[-1] if valid else None
    roic=last["NOPAT"]/last["InvestedCapital"] if last and last["InvestedCapital"] not in (None,0) else None
    inc=None
    if first and last and last["Year"]>first["Year"]:
        dcap=last["InvestedCapital"]-first["InvestedCapital"]; dn=last["NOPAT"]-first["NOPAT"]
        if abs(dcap)>1e-9:inc=dn/dcap
    wacc=None
    if "Cost of Capital" in wb.sheetnames:
        cws=wb["Cost of Capital"]
        r=_find_contains(cws,["base wacc"])
        if r:
            wacc=_num(cws.cell(r,2).value)
        if wacc is None:
            for rr in range(1,cws.max_row+1):
                if str(cws.cell(rr,1).value or "").strip().lower()=="base":
                    candidate=_num(cws.cell(rr,2).value)
                    if candidate is not None and 0<candidate<.50:
                        wacc=candidate
                        break
    spread=roic-wacc if roic is not None and wacc is not None else None
    latest=rows[-1] if rows else {}; market_cap=_num(wb["Company Data"]["B10"].value) if "Company Data" in wb.sheetnames else None
    buyback_yield=latest.get("Buybacks")/market_cap if latest.get("Buybacks") is not None and market_cap not in (None,0) else None
    net_reduction=latest.get("ShareReduction")
    metrics=[
        ("Latest ROIC",roic,"Return on operating capital","NOPAT / (Debt + Equity - Cash)"),
        ("ROIC - WACC spread",spread,"Value-creation spread","Positive spread indicates returns above current calculated cost of capital"),
        ("Incremental ROIC",inc,"Return on incremental capital","ΔNOPAT / ΔInvested Capital across comparable annual endpoints"),
        ("Latest gross buyback yield",buyback_yield,"Repurchase intensity","Annual repurchases / current market cap; not net of SBC"),
        ("Latest net share reduction",net_reduction,"Net dilution / shrinkage","Derived from NI/EPS implied diluted shares; positive = fewer shares"),
    ]
    for i,(label,value,interp,method) in enumerate(metrics,start+2):
        status="PASS" if value is not None else "REVIEW"; ws.cell(i,1,label); ws.cell(i,2,value); ws.cell(i,2).number_format=FMT_PCT
        ws.cell(i,3,status); ws.cell(i,3).fill=_status_fill(status); ws.cell(i,3).font=Font(bold=True)
        ws.cell(i,4,interp); ws.cell(i,5,method); ws.cell(i,5).alignment=Alignment(wrap_text=True)
    for c,w in {"A":28,"B":16,"C":13,"D":34,"E":75,"F":15,"G":15,"H":15,"I":15,"J":16,"K":16,"L":18}.items(): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"
    return {"roic":roic,"incremental_roic":inc,"roic_wacc_spread":spread,"net_share_reduction":net_reduction,"buyback_yield":buyback_yield}


def _historical_pe_rows(wb):
    if "Advanced Analytics" not in wb.sheetnames:return []
    ws=wb["Advanced Analytics"]; out=[]
    for r in range(7,min(ws.max_row,25)+1):
        year=_num(ws.cell(r,1).value)
        if year is not None:
            out.append({"Year":int(year),"Price":_num(ws.cell(r,2).value),"EPS":_num(ws.cell(r,3).value),"PE":_num(ws.cell(r,4).value)})
    return out


def ensure_valuation_history(wb,ticker):
    actuals=_historical_actuals(wb); rows=_historical_pe_rows(wb); ws=_new_sheet(wb,"Valuation History")
    _title(ws,f"{ticker} — Historical Valuation Regime",
           "Places current valuation against the company's own available history. Historical multiples provide context only and do not imply automatic mean reversion.",10)
    _section(ws,5,"Annual Valuation History",10)
    _header(ws,6,["Year","Year-End Price","EPS","P/E","FCF / Share","Price / FCF","FCF Yield","P/E Percentile","FCF Yield Percentile","Status"])
    pe=[x["PE"] for x in rows if x.get("PE") is not None and x["PE"]>0]; fy=[]; prepared=[]
    for x in rows:
        a=actuals.get(x["Year"],{}); shares=a.get("Shares"); fcf=a.get("FCF")
        fcfps=fcf/shares if fcf is not None and shares not in (None,0) else None
        pfcf=x["Price"]/fcfps if x.get("Price") not in (None,0) and fcfps not in (None,0) and fcfps>0 else None
        yield_=1/pfcf if pfcf not in (None,0) else None
        if yield_ is not None:fy.append(yield_)
        prepared.append({**x,"FCFShare":fcfps,"PFCF":pfcf,"FCFYield":yield_})
    def pctile(values,v):
        vals=sorted(x for x in values if x is not None)
        return sum(x<=v for x in vals)/len(vals) if vals and v is not None else None
    for r,x in enumerate(prepared,7):
        vals=[x["Year"],x["Price"],x["EPS"],x["PE"],x["FCFShare"],x["PFCF"],x["FCFYield"],pctile(pe,x["PE"]),pctile(fy,x["FCFYield"]),"PASS" if x["PE"] is not None else "REVIEW"]
        for c,v in enumerate(vals,1):ws.cell(r,c,v)
        ws.cell(r,2).number_format=FMT_PRICE; ws.cell(r,3).number_format=FMT_PRICE; ws.cell(r,4).number_format=FMT_MULT
        ws.cell(r,5).number_format=FMT_PRICE; ws.cell(r,6).number_format=FMT_MULT
        for c in (7,8,9):ws.cell(r,c).number_format=FMT_PCT
    start=max(7+len(prepared)+2,16); _section(ws,start,"Current Regime",10)
    _header(ws,start+1,["Metric","Current","Historical Median","Historical Percentile","Interpretation","Source / Caveat"])
    current_pe=_num(wb["Company Data"]["B15"].value) if "Company Data" in wb.sheetnames else None
    market_cap=_num(wb["Company Data"]["B10"].value) if "Company Data" in wb.sheetnames else None
    latest=actuals[max(actuals)] if actuals else {}
    current_fy=latest.get("FCF")/market_cap if latest.get("FCF") is not None and market_cap not in (None,0) else None
    stats=[
        ("Forward P/E",current_pe,statistics.median(pe) if pe else None,pctile(pe,current_pe),"Own-history multiple context","Current forward P/E vs trailing historical P/E is directional, not like-for-like."),
        ("Latest FCF yield",current_fy,statistics.median(fy) if fy else None,pctile(fy,current_fy),"Cash-flow valuation context","Latest reported FCF / current market cap."),
    ]
    for i,row in enumerate(stats,start+2):
        for c,v in enumerate(row,1):ws.cell(i,c,v)
        ws.cell(i,2).number_format=FMT_MULT if row[0]=="Forward P/E" else FMT_PCT
        ws.cell(i,3).number_format=FMT_MULT if row[0]=="Forward P/E" else FMT_PCT; ws.cell(i,4).number_format=FMT_PCT
        ws.cell(i,6).alignment=Alignment(wrap_text=True)
    for c,w in {"A":26,"B":18,"C":18,"D":20,"E":38,"F":70,"G":16,"H":18,"I":18,"J":14}.items():ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"
    return {"current_pe_percentile":pctile(pe,current_pe),"current_fcf_yield_percentile":pctile(fy,current_fy)}


def _segment_rows(wb):
    """Extract the latest issuer-reported segment-revenue column without estimating values."""
    if "Segment Analysis" not in wb.sheetnames:
        return []
    ws=wb["Segment Analysis"]; header=None; year_cols={}
    for r in range(1,min(ws.max_row,50)+1):
        first=str(ws.cell(r,1).value or "").strip().lower()
        years={}
        for col in range(2,min(ws.max_column,18)+1):
            raw=ws.cell(r,col).value
            value=_num(raw)
            if value is not None and 2000<=int(value)<=2100 and float(value)==int(value):
                years[int(value)]=col
                continue
            text=str(raw or "").strip()
            match=re.search(r"\b(20\d{2})\b",text)
            if match and ("revenue" in text.lower() or first in {"segment","business line / revenue group"}):
                years[int(match.group(1))]=col
        # Prefer the reportable-segment table. Business-line tables are only a fallback.
        if years and ("segment" in first or first=="business line / revenue group"):
            header=r; year_cols=years
            if "segment" in first:
                break
    if header is None or not year_cols:
        return []
    latest=max(year_cols); col=year_cols[latest]; out=[]; blanks=0
    for rr in range(header+1,min(ws.max_row,header+40)+1):
        name=str(ws.cell(rr,1).value or "").strip()
        if not name:
            blanks+=1
            if blanks>=3: break
            continue
        blanks=0
        low=name.lower()
        if any(x in low for x in ("source & data quality","revenue by business","business line / revenue group","total","consolidated")):
            if "revenue by business" in low or "source & data quality" in low:
                break
            continue
        revenue=_num(ws.cell(rr,col).value)
        if revenue is not None and revenue>0:
            out.append({"Segment":name,"FiscalYear":latest,"Revenue":revenue})
    return out


def ensure_sotp_framework(wb,ticker):
    ws=_new_sheet(wb,"SOTP Framework")
    policy=workbook_policy(wb,ticker)
    direct_value_mode=policy.key in {"bank","insurance","capital_markets","reit","insurance_conglomerate"}
    input_label="Analyst Segment Value" if direct_value_mode else "Analyst EV / Revenue"
    note=(f"{policy.label}: use a sourced direct segment value and document the sector-appropriate basis."
          if direct_value_mode else
          "Operating-company cross-check: reliable segment revenue can be paired with a sourced analyst EV/Revenue assumption.")
    _title(ws,f"{ticker} — Segment / Sum-of-the-Parts Framework",
           note+" SOTP never overwrites the primary valuation automatically.",10)
    segs=_segment_rows(wb)
    _section(ws,5,"Segment Valuation",10)
    _header(ws,6,["Segment","Fiscal Year","Reported Segment Revenue / Scale Metric",input_label,
                  "Implied Segment Value","Notes / Rationale","Input Status","Source Status","Weight","Decision Use"])
    for i,seg in enumerate(segs,7):
        ws.cell(i,1,seg["Segment"]); ws.cell(i,2,seg["FiscalYear"]); ws.cell(i,3,seg["Revenue"]); ws.cell(i,3).number_format=FMT_BN
        ws.cell(i,4,None); ws.cell(i,4).fill=_fill(INPUT); ws.cell(i,4).number_format=FMT_BN if direct_value_mode else FMT_MULT
        ws.cell(i,5,f'=IFERROR(D{i},"")' if direct_value_mode else f'=IFERROR(C{i}*D{i},"")'); ws.cell(i,5).number_format=FMT_BN
        ws.cell(i,6,("Enter sourced segment value and state the valuation basis (e.g. P/TBV, NAV, normalized earnings)."
                     if direct_value_mode else
                     "Enter a defensible segment EV/Revenue multiple and source/rationale."))
        ws.cell(i,7,"REVIEW — analyst input required"); ws.cell(i,7).fill=_fill(GOLD)
        ws.cell(i,8,"PASS — issuer segment scale metric")
        endrow=6+len(segs)
        ws.cell(i,9,f'=IFERROR(E{i}/SUM(E7:E{endrow}),"")'); ws.cell(i,9).number_format=FMT_PCT
        ws.cell(i,10,"Cross-check only; primary valuation remains authoritative.")
    start=max(8+len(segs),15)
    _section(ws,start,"SOTP Bridge",10)
    _header(ws,start+1,["Metric","Value","Status","Interpretation","Caveat"])
    net_cash=shares=price=None
    if "Company Data" in wb.sheetnames:
        d=wb["Company Data"]; cash=_num(d["B12"].value); debt=_num(d["B13"].value)
        net_cash=cash-debt if cash is not None and debt is not None else None
        shares=_num(d["B9"].value); price=_num(d["B8"].value)
    total_formula=f'=SUM(E7:E{6+len(segs)})' if segs else None
    equity_formula=(f'=IFERROR(B{start+2},"")' if direct_value_mode else f'=IFERROR(B{start+2}+B{start+3},"")') if segs else None
    bridge=[
        ("Segment Value Sum",total_formula,"REVIEW","Sum of analyst-supported segment values","Requires explicit analyst inputs"),
        ("Net Cash / (Debt)",net_cash,"PASS" if net_cash is not None else "REVIEW","Balance-sheet bridge",
         "For bank/insurance/REIT direct-value mode this is context only and is not automatically added."),
        ("SOTP Equity Value",equity_formula,"REVIEW","Indicative cross-check","Direct-value mode assumes analyst segment inputs are already equity-value appropriate."),
        ("SOTP Value / Share",f'=IFERROR(B{start+4}/{shares},"")' if segs and shares not in (None,0) else None,
         "REVIEW","Indicative value / diluted shares","Not used in primary valuation"),
        ("Current Price",price,"PASS" if price is not None else "REVIEW","Market reference","Current snapshot"),
        ("Business Model",policy.label,"PASS","Valuation routing",f"Primary valuation: {policy.primary_valuation}"),
    ]
    for idx,(label,value,status,interp,caveat) in enumerate(bridge,start+2):
        ws.cell(idx,1,label); ws.cell(idx,2,value)
        if isinstance(value,(int,float)) or (isinstance(value,str) and value.startswith("=")):
            ws.cell(idx,2).number_format=FMT_PRICE if "Share" in label or label=="Current Price" else FMT_BN
        ws.cell(idx,3,status); ws.cell(idx,3).fill=_status_fill(status); ws.cell(idx,3).font=Font(bold=True)
        ws.cell(idx,4,interp); ws.cell(idx,5,caveat); ws.cell(idx,5).alignment=Alignment(wrap_text=True)
    if not segs:
        ws["A7"]="REVIEW"
        ws["B7"]="No reliable numeric segment scale table was detected. SOTP remains disabled rather than fabricating segment economics."
        ws["A7"].fill=_fill(GOLD); ws.merge_cells("B7:J8"); ws["B7"].alignment=Alignment(wrap_text=True,vertical="top")
    for col,width in {"A":34,"B":14,"C":23,"D":22,"E":20,"F":65,"G":25,"H":25,"I":14,"J":38}.items():
        ws.column_dimensions[col].width=width
    ws.freeze_panes="A7"
    return {"segments":len(segs),"status":"REVIEW — analyst inputs required" if segs else "REVIEW",
            "policy":policy.key,"direct_value_mode":direct_value_mode}

def _timeline_events(history,current):
    events=[];snaps=history.get("snapshots",[])
    for idx,snap in enumerate(snaps[-12:]):
        est=snap.get("estimates") or []
        if not est:continue
        pos=neg=0
        if idx>0:
            prev=snaps[-12:][idx-1];pmap={(x.get("Metric"),x.get("FiscalYear")):_num(x.get("Consensus")) for x in prev.get("estimates",[])}
            for row in est:
                cur=_num(row.get("Consensus"));old=pmap.get((row.get("Metric"),row.get("FiscalYear")))
                if cur is not None and old not in (None,0):
                    change=cur/old-1;pos+=change>.02;neg+=change<-.02
        signal="POSITIVE" if pos>neg else "NEGATIVE" if neg>pos else "MIXED / UNCHANGED"
        events.append({"Date":snap.get("captured_at"),"Event":"Forecast / consensus snapshot","Signal":signal,"Evidence":f"{len(est)} estimate row(s); positive revisions={pos}, negative revisions={neg}","Source":"Local forecast history","Action":"Review assumptions; no automatic DCF change."})
    for e in current.get("earnings_surprises") or []:
        surprise=_num(e.get("Surprise"));signal="POSITIVE" if surprise is not None and surprise>.03 else "NEGATIVE" if surprise is not None and surprise<-.03 else "MIXED"
        events.append({"Date":str(e.get("Reported")),"Event":"Earnings result","Signal":signal,"Evidence":f"EPS surprise {surprise:.1%}" if surprise is not None else "EPS surprise unavailable","Source":"Advanced Analytics provider history","Action":"Compare with revision and KPI evidence."})
    return events


def ensure_thesis_timeline(wb,ticker,history,current,research_root=None):
    ws=_new_sheet(wb,"Thesis Timeline")
    _title(ws,f"{ticker} — Catalyst & Thesis-Change Timeline",
           "Chronological evidence log from stored estimates, earnings surprises and KPI snapshots. Signals are triage labels; valuation changes remain analyst-reviewed.",8)
    _section(ws,5,"Evidence Timeline",8);_header(ws,6,["Date","Event","Signal","Evidence","Source","Analyst Action","Thesis Effect","Notes"])
    events=_timeline_events(history,current)
    if research_root:
        kp=Path(research_root)/str(ticker).upper()/"kpi_history.json"
        if kp.exists():
            try:
                data=json.loads(kp.read_text(encoding="utf-8"))
                for snap in data.get("snapshots",[])[-12:]:
                    events.append({"Date":snap.get("captured_at"),"Event":"KPI evidence snapshot","Signal":"REVIEW","Evidence":f"{len(snap.get('kpis') or [])} KPI/evidence row(s) captured","Source":"KPI / Earnings Agent history","Action":"Review changes versus prior KPI snapshot."})
            except Exception:pass
    events=sorted(events,key=lambda x:str(x.get("Date") or ""))
    for r,e in enumerate(events,7):
        vals=[e.get("Date"),e.get("Event"),e.get("Signal"),e.get("Evidence"),e.get("Source"),e.get("Action"),e.get("Signal"),None]
        for c,v in enumerate(vals,1):ws.cell(r,c,v)
        ws.cell(r,4).alignment=Alignment(wrap_text=True);ws.cell(r,6).alignment=Alignment(wrap_text=True)
    if not events:
        ws["A7"]="REVIEW";ws["B7"]="No stored revision/KPI/earnings events yet. The timeline fills as repeated offline research runs accumulate evidence."
    for c,w in {"A":24,"B":28,"C":18,"D":62,"E":34,"F":52,"G":18,"H":38}.items():ws.column_dimensions[c].width=w
    ws.freeze_panes="A7";return {"events":len(events)}


def _quality_row(wb,label,status,detail,rule):
    if "Data Quality" not in wb.sheetnames:return
    ws=wb["Data Quality"];r=None
    for rr in range(1,ws.max_row+1):
        if str(ws.cell(rr,1).value or "").strip()==label:r=rr;break
    r=r or ws.max_row+1
    ws.cell(r,1,label);ws.cell(r,2,status);ws.cell(r,3,detail);ws.cell(r,4,rule)
    ws.cell(r,2).fill=_status_fill(status);ws.cell(r,2).font=Font(bold=True)
    for c in range(1,5):ws.cell(r,c).alignment=Alignment(wrap_text=True,vertical="top")


def ensure_offline_equity_extensions(wb,ticker,info=None,research_root=None,persist_history=True,captured_at=None):
    ticker=str(ticker).upper().strip();root=Path(research_root) if research_root else Path(__file__).resolve().parent/"research_data"
    payload=_snapshot(wb,ticker,captured_at);history=_load_history(root,ticker)
    if persist_history:
        _,changed=persist_forecast_snapshot(root,ticker,payload);history=_load_history(root,ticker)
    else:
        changed=False;snaps=list(history.get("snapshots",[]))
        if not snaps or _snapshot_key(snaps[-1])!=_snapshot_key(payload):snaps.append(payload)
        history={**history,"snapshots":snaps}
    actuals=_historical_actuals(wb);acc=ensure_forecast_accountability(wb,ticker,history,actuals,payload)
    if persist_history:write_accuracy_csv(root,ticker,acc["summary"])
    ensure_earnings_revisions(wb,ticker,history,payload)
    cap=ensure_capital_allocation(wb,ticker,info or {});val=ensure_valuation_history(wb,ticker)
    sotp=ensure_sotp_framework(wb,ticker);timeline=ensure_thesis_timeline(wb,ticker,history,payload,research_root=root)
    _quality_row(wb,"Forecast-accountability history","PASS" if history.get("snapshots") else "REVIEW",f"{len(history.get('snapshots') or [])} local forecast snapshot(s); {len(acc['records'])} matured forecast observation(s).","Only forecasts with a corresponding reported fiscal-year actual are graded.")
    _quality_row(wb,"Capital-allocation diagnostics","PASS" if cap.get("roic") is not None else "REVIEW",f"ROIC={cap.get('roic')}; incremental ROIC={cap.get('incremental_roic')}; net share reduction={cap.get('net_share_reduction')}.","Missing statement rows remain blank; no acquisition/buyback/capital values are fabricated.")
    _quality_row(wb,"Historical valuation context","PASS" if val.get("current_pe_percentile") is not None else "REVIEW",f"Current forward P/E historical percentile context={val.get('current_pe_percentile')}.","Forward current P/E vs trailing historical P/E is directional context.")
    _quality_row(wb,"SOTP framework","REVIEW",f"{sotp.get('segments',0)} numeric segment row(s) detected.","Segment multiples are analyst inputs; SOTP never overwrites primary DCF automatically.")
    return {"ticker":ticker,"snapshot_changed":changed,"forecast_snapshots":len(history.get("snapshots") or []),"matured_forecast_records":len(acc["records"]),"capital_allocation":cap,"valuation_history":val,"sotp":sotp,"timeline":timeline}
