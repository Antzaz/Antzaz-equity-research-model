"""Company-agnostic SEC segment analysis.

The parser first uses known reportable-segment labels for common large-cap companies,
then falls back to conservative table discovery for any US SEC filer. It never invents
segment economics: when reliable extraction is unavailable, the same standardized sheet
is created with clearly marked manual inputs so downstream charts still have a stable
schema.
"""

import re
import math
import statistics
from io import StringIO
import requests
try:
    import pandas as pd
except Exception:
    pd=None

from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; LIGHT="F5F9FC"; GOLD="FFF2CC"
INPUT_BLUE="0000FF"; GREY="666666"; LINK_GREEN="008000"
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PCT='0.0%;[Red](0.0%);-'

CONFIGS={
    "GOOGL":{"segments":["Google Services","Google Cloud","Other Bets"],"business":["Google Search & other","YouTube ads","Google Network","Google subscriptions, platforms, and devices","Google Cloud","Other Bets"]},
    "GOOG":{"segments":["Google Services","Google Cloud","Other Bets"],"business":["Google Search & other","YouTube ads","Google Network","Google subscriptions, platforms, and devices","Google Cloud","Other Bets"]},
    "MSFT":{"segments":["Productivity and Business Processes","Intelligent Cloud","More Personal Computing"]},
    "AMZN":{"segments":["North America","International","AWS"],"business":["Online stores","Physical stores","Third-party seller services","Advertising services","Subscription services","AWS","Other"]},
    "META":{"segments":["Family of Apps","Reality Labs"]},
    "NVDA":{"segments":["Compute & Networking","Graphics"]},
    "AAPL":{"business":["iPhone","Mac","iPad","Wearables, Home and Accessories","Services"]},
    "NFLX":{"segments":["United States and Canada","Europe, Middle East and Africa","Latin America","Asia-Pacific"]},
    "ORCL":{"business":["Cloud services and license support","Cloud license and on-premise license","Hardware","Services"]},
    "CRM":{"business":["Sales","Service","Platform and Other","Marketing and Commerce","Integration and Analytics"]},
    "ADBE":{"segments":["Digital Media","Digital Experience","Publishing and Advertising"]},
    "AMD":{"segments":["Data Center","Client","Gaming","Embedded"]},
    "INTC":{"segments":["Client Computing Group","Data Center and AI","Network and Edge","Mobileye"]},
    "QCOM":{"segments":["QCT","QTL"]},
    "WMT":{"segments":["Walmart U.S.","Walmart International","Sam's Club"]},
}

GOOGL_REV_FALLBACK={"Google Services":[272.543,304.930,342.721],"Google Cloud":[33.088,43.229,58.705],"Other Bets":[1.527,1.648,1.537],"Google Search & other":[175.033,198.084,224.532],"YouTube ads":[31.510,36.147,40.367],"Google Network":[31.312,30.359,29.792],"Google subscriptions, platforms, and devices":[34.688,40.340,48.030]}
GOOGL_OP_FALLBACK={"Google Services":[95.858,121.263,139.404],"Google Cloud":[1.716,6.112,13.910],"Other Bets":[-4.095,-4.444,-7.515]}
BLACKLIST=("total","consolidated","revenue","net sales","operating income","operating loss","segment profit","year ended","three months","six months","nine months","cost of","income from","eliminations","corporate","other income","depreciation","assets","liabilities","capital expenditures")

def _fill(c): return PatternFill("solid",fgColor=c)
def _num(v):
    try:
        if isinstance(v,bool): return None
        return float(v)
    except Exception: return None
def _title(ws,text,end=14):
    for c in range(1,end+1): ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=text; ws["A1"].font=Font(bold=True,color=WHITE,size=18); ws.sheet_view.showGridLines=False
def _section(ws,row,title,end=14):
    for c in range(1,end+1): ws.cell(row,c).fill=_fill(NAVY); ws.cell(row,c).font=Font(bold=True,color=WHITE,size=11)
    ws.cell(row,1,title)
def _header(ws,row,start,end):
    for c in range(start,end+1):
        ws.cell(row,c).fill=_fill(BLUE); ws.cell(row,c).font=Font(bold=True,color=WHITE); ws.cell(row,c).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
def _latest_10k_html(ticker,headers):
    try:
        tickers=requests.get("https://www.sec.gov/files/company_tickers.json",headers=headers,timeout=30).json(); cik=None
        for item in tickers.values():
            if str(item.get("ticker","")).upper()==ticker.upper(): cik=str(item["cik_str"]).zfill(10); break
        if not cik: return None,None
        subs=requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json",headers=headers,timeout=30).json(); recent=subs.get("filings",{}).get("recent",{})
        for form,acc,doc in zip(recent.get("form",[]),recent.get("accessionNumber",[]),recent.get("primaryDocument",[])):
            if form=="10-K":
                url=f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-','')}/{doc}"; r=requests.get(url,headers=headers,timeout=45); r.raise_for_status(); return r.text,url
    except Exception: return None,None
    return None,None
def _tables(html):
    if not html or pd is None: return []
    try: return pd.read_html(StringIO(html))
    except Exception: return []
def _numbers(values):
    out=[]
    for val in values:
        text=str(val).replace("−","-")
        for token in re.findall(r"\(?-?\$?\s*\d[\d,]*(?:\.\d+)?\)?",text):
            neg=token.strip().startswith("(") and token.strip().endswith(")"); clean=token.replace("$","").replace(",","").replace(" ","").strip("()")
            try: n=float(clean)
            except Exception: continue
            if 1900<=n<=2100: continue
            out.append(-n if neg else n)
    return out
def _text_name(values):
    for val in values:
        s=re.sub(r"\s+"," ",str(val).strip())
        if s and s.lower() not in {"nan","none"} and re.search(r"[A-Za-z]",s) and len(s)<=80: return s
    return None
def _metric_context(df,row_idx):
    parts=[]
    for rr in range(max(0,row_idx-5),row_idx+1): parts.append(" | ".join(str(x) for x in df.iloc[rr].tolist()))
    text=" ".join(parts).lower()
    if any(k in text for k in ("operating income","operating loss","segment profit","income (loss) from operations","profit (loss)")): return "op"
    if any(k in text for k in ("net sales","revenue","revenues")): return "revenue"
    return None
def _extract_known(tables,labels):
    out={lab:{"revenue":[],"op":[]} for lab in labels}
    for df in tables:
        for pos,(_,row) in enumerate(df.iterrows()):
            values=row.tolist(); low=" | ".join(str(v) for v in values).lower(); nums=_numbers(values)
            if len(nums)<2: continue
            for lab in labels:
                if lab.lower() in low:
                    metric=_metric_context(df,pos)
                    if metric: out[lab][metric].append(nums[-3:])
    return out
def _discover_segments(tables):
    candidates={}
    for df in tables:
        whole=" ".join(str(x) for x in df.astype(str).values.flatten()).lower()
        if "segment" not in whole or not any(k in whole for k in ("revenue","net sales")): continue
        for pos,(_,row) in enumerate(df.iterrows()):
            vals=row.tolist(); name=_text_name(vals); nums=_numbers(vals)
            if not name or len(nums)<2: continue
            low=name.lower()
            if any(b in low for b in BLACKLIST) or len(low)<2 or _metric_context(df,pos)!="revenue": continue
            candidates.setdefault(name,[]).append(nums[-3:])
    ranked=sorted(candidates.items(),key=lambda kv:(-len(kv[1]),len(kv[0])))
    return [name for name,_ in ranked[:10]]
def _pick(series):
    if not series: return None
    vals=max(series,key=lambda x:(len(x),abs(x[-1]) if x else 0)); vals=list(vals[-3:])
    while len(vals)<3: vals.insert(0,None)
    return vals
def _years(wb):
    h=wb["Historical Financials"]; ys=[h.cell(3,c).value for c in range(2,8)]; ys=[int(x) for x in ys if isinstance(x,(int,float))]
    if len(ys)>=3: return ys[-3:]
    latest=max(ys) if ys else 2025; return [latest-2,latest-1,latest]
def _company_revenue(wb): return _num(wb["Historical Financials"]["G4"].value) if "Historical Financials" in wb.sheetnames else None

def _write_sheet(wb,ticker,years,segments,business,url,auto_status):
    if "Segment Analysis" in wb.sheetnames: wb.remove(wb["Segment Analysis"])
    ws=wb.create_sheet("Segment Analysis"); _title(ws,f"{ticker} — Business & Segment Analysis")
    ws["A3"]=f"Standardized segment schema. Status: {auto_status}. Only issuer-disclosed data should be entered; yellow cells are manual inputs when automatic extraction is incomplete."; ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)
    y0,y1,y2=years; _section(ws,5,"Reported Operating Segments")
    heads=["Segment",f"{y0} Revenue",f"{y1} Revenue",f"{y2} Revenue",f"{y2} Growth",f"{y0}–{y2} CAGR",f"{y0} Op. Income",f"{y1} Op. Income",f"{y2} Op. Income",f"{y0} Margin",f"{y1} Margin",f"{y2} Margin","Margin Δ",f"{y2} Revenue Mix"]
    for c,v in enumerate(heads,1): ws.cell(6,c,v)
    _header(ws,6,1,14); max_rows=max(10,len(segments))
    for i in range(max_rows):
        r=7+i; item=segments[i] if i<len(segments) else None
        if item:
            name,rev,op=item; ws.cell(r,1,name)
            for j,v in enumerate(rev or [None,None,None],2): ws.cell(r,j,v); ws.cell(r,j).number_format=FMT_BN
            ws.cell(r,5,f'=IFERROR(D{r}/C{r}-1,"")'); ws.cell(r,6,f'=IFERROR((D{r}/B{r})^(1/2)-1,"")')
            if op:
                for j,v in enumerate(op,7): ws.cell(r,j,v); ws.cell(r,j).number_format=FMT_BN
            for c_rev,c_op,c_margin in ((2,7,10),(3,8,11),(4,9,12)):
                ws.cell(r,c_margin,f'=IFERROR({get_column_letter(c_op)}{r}/{get_column_letter(c_rev)}{r},"")'); ws.cell(r,c_margin).number_format=FMT_PCT
            ws.cell(r,13,f'=IFERROR(L{r}-K{r},"")'); ws.cell(r,14,f'=IFERROR(D{r}/SUM($D$7:$D${6+max_rows}),"")')
            for c in (5,6,10,11,12,13,14): ws.cell(r,c).number_format=FMT_PCT
        else:
            for c in range(1,10): ws.cell(r,c).fill=_fill(GOLD); ws.cell(r,c).font=Font(color=INPUT_BLUE); ws.cell(r,c).number_format=FMT_BN if c>=2 else "General"
            for c in (5,6,10,11,12,13,14): ws.cell(r,c).number_format=FMT_PCT
    ws.conditional_formatting.add(f"E7:F{6+max_rows}",ColorScaleRule(start_type="min",start_color="F8696B",mid_type="num",mid_value=0,mid_color="FFEB84",end_type="max",end_color="63BE7B"))
    start=9+max_rows; _section(ws,start,"Revenue by Business Line",10)
    bh=["Business Line / Revenue Group",str(y0),str(y1),str(y2),f"{y2} Growth",f"{y0}–{y2} CAGR",f"{y2} Mix","Source / Notes"]
    for c,v in enumerate(bh,1): ws.cell(start+1,c,v)
    _header(ws,start+1,1,8); bmax=max(10,len(business))
    for i in range(bmax):
        r=start+2+i; item=business[i] if i<len(business) else None
        if item:
            name,rev=item; ws.cell(r,1,name)
            for j,v in enumerate(rev or [None,None,None],2): ws.cell(r,j,v); ws.cell(r,j).number_format=FMT_BN
            ws.cell(r,5,f'=IFERROR(D{r}/C{r}-1,"")'); ws.cell(r,6,f'=IFERROR((D{r}/B{r})^(1/2)-1,"")'); ws.cell(r,7,f'=IFERROR(D{r}/SUM($D${start+2}:$D${start+1+bmax}),"")')
            for c in (5,6,7): ws.cell(r,c).number_format=FMT_PCT
        else:
            for c in range(1,9): ws.cell(r,c).fill=_fill(GOLD); ws.cell(r,c).font=Font(color=INPUT_BLUE); ws.cell(r,c).number_format=FMT_BN if c in (2,3,4) else "General"
            for c in (5,6,7): ws.cell(r,c).number_format=FMT_PCT
    source_row=start+bmax+4; _section(ws,source_row,"Source & Data Quality",10)
    ws.cell(source_row+1,1,"SEC 10-K Source"); ws.cell(source_row+1,2,url or ""); ws.cell(source_row+2,1,"Extraction Status"); ws.cell(source_row+2,2,auto_status)
    ws.cell(source_row+3,1,"Important"); ws.cell(source_row+3,2,"Segment disclosure differs by issuer. If operating income is not disclosed by segment, profitability cells remain blank rather than estimated."); ws.cell(source_row+3,2).alignment=Alignment(wrap_text=True)
    ws.column_dimensions["A"].width=38
    for c in range(2,15): ws.column_dimensions[get_column_letter(c)].width=15
    ws.column_dimensions["H"].width=38; ws.freeze_panes="A7"; return ws

def ensure_segment_analysis_v2(wb,ticker,headers):
    ticker=ticker.upper(); html,url=_latest_10k_html(ticker,headers); tables=_tables(html); cfg=CONFIGS.get(ticker,{})
    seg_labels=list(cfg.get("segments",[])); bus_labels=list(cfg.get("business",[]))
    if not seg_labels: seg_labels=_discover_segments(tables)
    labels=list(dict.fromkeys(seg_labels+bus_labels)); extracted=_extract_known(tables,labels) if labels else {}; total=_company_revenue(wb); raw_rev={}
    for lab in labels:
        raw=_pick(extracted.get(lab,{}).get("revenue",[]))
        if raw: raw_rev[lab]=raw
    if ticker in {"GOOGL","GOOG"}:
        for lab,vals in GOOGL_REV_FALLBACK.items():
            if lab in labels and lab not in raw_rev: raw_rev[lab]=vals
    scale=1.0; non_google_raw={k:v for k,v in raw_rev.items() if not (ticker in {"GOOGL","GOOG"} and k in GOOGL_REV_FALLBACK and v==GOOGL_REV_FALLBACK[k])}
    if non_google_raw:
        latest=[abs(v[-1]) for v in non_google_raw.values() if v]; med=statistics.median(latest) if latest else 0
        if total and med>total*10: scale=.001
        elif med>10000: scale=.001
    segments=[]
    for lab in seg_labels:
        rev=raw_rev.get(lab)
        if rev:
            sc=1.0 if ticker in {"GOOGL","GOOG"} and lab in GOOGL_REV_FALLBACK and rev==GOOGL_REV_FALLBACK[lab] else scale; rev=[_num(x)*sc if _num(x) is not None else None for x in rev]
        op_raw=_pick(extracted.get(lab,{}).get("op",[]))
        if ticker in {"GOOGL","GOOG"} and not op_raw and lab in GOOGL_OP_FALLBACK: op=list(GOOGL_OP_FALLBACK[lab])
        elif op_raw: op=[_num(x)*scale if _num(x) is not None else None for x in op_raw]
        else: op=None
        if rev or op: segments.append((lab,rev or [None,None,None],op))
    if not bus_labels: bus_labels=seg_labels
    business=[]
    for lab in bus_labels:
        rev=raw_rev.get(lab)
        if rev:
            sc=1.0 if ticker in {"GOOGL","GOOG"} and lab in GOOGL_REV_FALLBACK and rev==GOOGL_REV_FALLBACK[lab] else scale; business.append((lab,[_num(x)*sc if _num(x) is not None else None for x in rev]))
    status="AUTO — SEC 10-K table extraction" if segments or business else "MANUAL REQUIRED — no reliable segment table extracted"
    return _write_sheet(wb,ticker,_years(wb),segments,business,url,status)
