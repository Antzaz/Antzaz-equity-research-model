"""Decision-focused investment summary for generated equity-research workbooks.

The summary uses existing model engines instead of cached Excel formula values wherever
possible. It presents the investment decision, the actual numbers behind it, and a transparent
neutral-50 quantitative score bridge. Missing dimensions remain visible and are explicitly
marked rather than silently disappearing from the score.
"""

from __future__ import annotations

import math
import statistics

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Reuse the exact deterministic valuation engine already used by Advanced Analytics. This
# avoids depending on formula-cache values in Dashboard/DCF when openpyxl builds the workbook.
from advanced_analytics_v2 import _base_value as _analytics_value

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; LIGHT="F5F9FC"; GREY="666666"
PALE_GREEN="E2F0D9"; PALE_RED="FCE4D6"; PALE_YELLOW="FFF2CC"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_MULT='0.0x;[Red](0.0x);-'; FMT_SCORE='0.0'
THIN=Side(style="thin",color="D9E1F2"); MEDIUM=Side(style="medium",color=NAVY)

DIMENSIONS=[
    ("Absolute valuation",30),("Relative valuation",15),("Growth",10),("Profitability",10),
    ("FCF quality",10),("Balance sheet",10),("Stress resilience",10),("Bayesian skew",5),
]


def _fill(c): return PatternFill("solid",fgColor=c)

def _num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        x=float(v); return x if math.isfinite(x) else default
    except Exception: return default

def _clamp(x,lo=0,hi=100): return max(lo,min(hi,x))

def _find_row(ws,label,col=1):
    if ws is None: return None
    needle=str(label).strip().lower()
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,col).value or "").strip().lower()==needle: return r
    return None

def _label_value(wb,sheet,label,value_col=2):
    if sheet not in wb.sheetnames: return None
    ws=wb[sheet]; r=_find_row(ws,label,1)
    return ws.cell(r,value_col).value if r else None

def _any_label_value(wb,sheet,label,search_cols=(1,),value_offset=1):
    if sheet not in wb.sheetnames: return None
    ws=wb[sheet]; needle=str(label).strip().lower()
    for c in search_cols:
        for r in range(1,ws.max_row+1):
            if str(ws.cell(r,c).value or "").strip().lower()==needle:
                return ws.cell(r,c+value_offset).value
    return None

def _style_title(ws,text):
    for c in range(1,11): ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=text; ws["A1"].font=Font(bold=True,color=WHITE,size=19); ws.sheet_view.showGridLines=False

def _section(ws,row,title):
    for c in range(1,11):
        x=ws.cell(row,c); x.fill=_fill(NAVY); x.font=Font(bold=True,color=WHITE,size=11); x.border=Border(bottom=MEDIUM)
    ws.cell(row,1,title)

def _header(ws,row,headers):
    for c,v in enumerate(headers,1):
        x=ws.cell(row,c,v); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE)
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); x.border=Border(bottom=THIN)

def _score_fill(score):
    if score is None: return _fill(LIGHT)
    return _fill(PALE_GREEN if score>=65 else PALE_RED if score<40 else PALE_YELLOW)

def _latest_num(ws,row,start=2,end=7):
    if ws is None: return None
    for c in range(end,start-1,-1):
        v=_num(ws.cell(row,c).value)
        if v is not None: return v
    return None

def _history_cagr(ws,row):
    if ws is None: return None
    pts=[]
    for c in range(2,8):
        y=ws.cell(3,c).value; v=_num(ws.cell(row,c).value)
        if isinstance(y,(int,float)) and v is not None and v>0: pts.append((int(y),v))
    if len(pts)<2: return None
    y0,v0=pts[0]; y1,v1=pts[-1]; years=max(1,y1-y0)
    return (v1/v0)**(1/years)-1 if v0>0 and v1>0 else None

def _scenario_base_cagr(wb):
    if "Three-Case Scenarios" not in wb.sheetnames: return None
    s=wb["Three-Case Scenarios"]; vals=[_num(s.cell(12,c).value) for c in range(14,24)]; vals=[x for x in vals if x is not None]
    if not vals: return None
    p=1.0
    for g in vals: p*=1+g
    return p**(1/len(vals))-1

def _data_quality(wb):
    if "Data Quality" not in wb.sheetnames: return 0,0,0
    p=r=f=0; ws=wb["Data Quality"]
    for row in range(1,ws.max_row+1):
        s=str(ws.cell(row,2).value or "").upper().strip()
        if s=="PASS": p+=1
        elif s=="REVIEW": r+=1
        elif s=="FAIL": f+=1
    return p,r,f

def _advanced_score(wb,label):
    if "Advanced Analytics" not in wb.sheetnames: return None
    ws=wb["Advanced Analytics"]; r=_find_row(ws,label,1); return _num(ws.cell(r,2).value) if r else None

def _bayes_prob(wb,scenario):
    if "Base Rates & Probabilities" not in wb.sheetnames: return None
    ws=wb["Base Rates & Probabilities"]; r=_find_row(ws,scenario,1); return _num(ws.cell(r,4).value) if r else None


def _peer_snapshot(wb):
    """Target metrics and peer medians from authoritative Peer Comps; direct/exact peers first."""
    out={"target":{},"medians":{},"direct_medians":{},"coverage":None}
    if "Peer Comps" not in wb.sheetnames: return out
    ws=wb["Peer Comps"]; headers={}
    for c in range(1,ws.max_column+1):
        h=str(ws.cell(3,c).value or "").strip()
        if h: headers[h]=c
    metrics=["Forward P/E","EV/Revenue","EV/EBITDA","Revenue Growth","Operating Margin","ROE"]
    if not any(x in headers for x in metrics): return out
    target_row=None; peers=[]; direct=[]
    for r in range(4,min(ws.max_row,40)+1):
        symbol=str(ws.cell(r,headers.get("Ticker",2)).value or "").strip()
        if not symbol: continue
        typ=str(ws.cell(r,headers.get("Peer Type",16)).value or "").strip()
        if typ=="Target classification": target_row=r
        else:
            peers.append(r)
            if typ in {"Direct business-model peer","Exact industry"}: direct.append(r)
    if target_row is None and str(ws.cell(4,headers.get("Ticker",2)).value or "").strip(): target_row=4
    if target_row:
        for h in metrics:
            if h in headers: out["target"][h]=_num(ws.cell(target_row,headers[h]).value)
        if "Data Coverage %" in headers: out["coverage"]=_num(ws.cell(target_row,headers["Data Coverage %"]).value)
    def med(rows,h):
        if h not in headers: return None
        vals=[_num(ws.cell(r,headers[h]).value) for r in rows]; vals=[x for x in vals if x is not None]
        return statistics.median(vals) if vals else None
    for h in metrics:
        out["medians"][h]=med(peers,h); d=med(direct,h); out["direct_medians"][h]=d if d is not None else out["medians"][h]
    return out


def _next_year_consensus(wb):
    out={"revenue_consensus":None,"revenue_model":None,"revenue_gap":None,"eps_consensus":None,"eps_model":None,"eps_gap":None}
    if "Expectations & Consensus" not in wb.sheetnames: return out
    ws=wb["Expectations & Consensus"]; rows=[]
    for r in range(7,min(ws.max_row,50)+1):
        metric=str(ws.cell(r,1).value or "").strip(); y=ws.cell(r,2).value
        if metric in {"Revenue","EPS"} and isinstance(y,(int,float)): rows.append((int(y),metric,r))
    if not rows: return out
    y0=min(y for y,_,_ in rows)
    for y,metric,r in rows:
        if y!=y0: continue
        cons=_num(ws.cell(r,3).value); model=_num(ws.cell(r,4).value); gap=model/cons-1 if cons not in (None,0) and model is not None else None
        k=metric.lower(); out[f"{k}_consensus"]=cons; out[f"{k}_model"]=model; out[f"{k}_gap"]=gap
    return out


def _safe_model_value(wb,*args):
    try: return _num(_analytics_value(wb,*args))
    except Exception: return None


def _metric_snapshot(wb):
    d=wb["Company Data"] if "Company Data" in wb.sheetnames else None
    h=wb["Historical Financials"] if "Historical Financials" in wb.sheetnames else None
    s=wb["Three-Case Scenarios"] if "Three-Case Scenarios" in wb.sheetnames else None
    price=_num(d["B8"].value) if d else None; mc=_num(d["B10"].value) if d else None; ev=_num(d["B11"].value) if d else None
    cash=_num(d["B12"].value) if d else None; debt=_num(d["B13"].value) if d else None
    net_debt=(debt-cash) if debt is not None and cash is not None else None; net_debt_mc=net_debt/mc if net_debt is not None and mc not in (None,0) else None
    sector=d["B6"].value if d else None; industry=d["B7"].value if d else None

    # Cached Dashboard formulas are useful when available, but deterministic model values are
    # the fallback so the summary remains populated in freshly generated openpyxl workbooks.
    base=_num(_label_value(wb,"Dashboard","Base DCF Value / Share"))
    if base is None: base=_safe_model_value(wb)
    pw=_num(_label_value(wb,"Dashboard","Probability-Weighted Value / Share"))
    severe=_num(_label_value(wb,"Three-Case Scenarios","Severe Bear Value / Share",2))
    if severe is None: severe=_safe_model_value(wb,-.05,-.05,.05,.02,-.01)
    base_up=base/price-1 if base is not None and price else None; pw_up=pw/price-1 if pw is not None and price else None
    severe_down=severe/price-1 if severe is not None and price else None

    revenue=_latest_num(h,4); op=_latest_num(h,9); ni=_latest_num(h,11); ocf=_latest_num(h,14); capex=_latest_num(h,15)
    fcf=(ocf-abs(capex)) if ocf is not None and capex is not None else None
    opm=op/revenue if revenue not in (None,0) and op is not None else None; fcfm=fcf/revenue if revenue not in (None,0) and fcf is not None else None
    rev_cagr=_history_cagr(h,4); eps_cagr=_history_cagr(h,12)

    peer=_peer_snapshot(wb); target=peer["target"]; med=peer["direct_medians"]
    fpe=target.get("Forward P/E"); evrev=target.get("EV/Revenue"); evebitda=target.get("EV/EBITDA"); latest_growth=target.get("Revenue Growth"); roe=target.get("ROE")
    peer_fpe=med.get("Forward P/E"); peer_evrev=med.get("EV/Revenue"); peer_evebitda=med.get("EV/EBITDA"); peer_growth=med.get("Revenue Growth"); peer_margin=med.get("Operating Margin"); peer_roe=med.get("ROE")
    fpe_gap=fpe/peer_fpe-1 if fpe is not None and peer_fpe not in (None,0) else None
    evrev_gap=evrev/peer_evrev-1 if evrev is not None and peer_evrev not in (None,0) else None
    evebitda_gap=evebitda/peer_evebitda-1 if evebitda is not None and peer_evebitda not in (None,0) else None
    growth_gap=latest_growth-peer_growth if latest_growth is not None and peer_growth is not None else None
    margin_gap=opm-peer_margin if opm is not None and peer_margin is not None else None; roe_gap=roe-peer_roe if roe is not None and peer_roe is not None else None

    wacc=_num(s["C6"].value) if s else None; tgr=_num(s["C7"].value) if s else None; base_10y=_scenario_base_cagr(wb)
    market_base=_num(_label_value(wb,"Market Expectations","10Y Revenue CAGR",2)); market_implied=_num(_label_value(wb,"Market Expectations","10Y Revenue CAGR",3))
    market_gap=market_implied-market_base if market_implied is not None and market_base is not None else None

    mc_median=_num(_any_label_value(wb,"Advanced Analytics","Median Value / Share",(9,1)))
    mc_p10=_num(_any_label_value(wb,"Advanced Analytics","P10 Value / Share",(9,1)))
    mc_prob=_num(_any_label_value(wb,"Advanced Analytics","Probability > Current Price",(9,1)))
    implied_fcf=_num(_any_label_value(wb,"Advanced Analytics","Implied 10Y FCF CAGR",(1,9)))
    consensus=_next_year_consensus(wb)
    return {
        "price":price,"market_cap":mc,"enterprise_value":ev,"cash":cash,"debt":debt,"net_debt":net_debt,"net_debt_mc":net_debt_mc,"sector":sector,"industry":industry,
        "base":base,"pw":pw,"base_up":base_up,"pw_up":pw_up,"severe":severe,"severe_down":severe_down,"revenue":revenue,"net_income":ni,"fcf":fcf,
        "rev_cagr":rev_cagr,"eps_cagr":eps_cagr,"latest_growth":latest_growth,"op_margin":opm,"fcf_margin":fcfm,"roe":roe,
        "fpe":fpe,"peer_fpe":peer_fpe,"fpe_gap":fpe_gap,"evrev":evrev,"peer_evrev":peer_evrev,"evrev_gap":evrev_gap,"evebitda":evebitda,"peer_evebitda":peer_evebitda,"evebitda_gap":evebitda_gap,
        "peer_growth":peer_growth,"growth_gap":growth_gap,"peer_margin":peer_margin,"margin_gap":margin_gap,"peer_roe":peer_roe,"roe_gap":roe_gap,"peer_coverage":peer.get("coverage"),
        "growth_score":_advanced_score(wb,"Growth"),"profit_score":_advanced_score(wb,"Profitability"),"fcf_score":_advanced_score(wb,"FCF Quality"),"balance_score":_advanced_score(wb,"Balance Sheet"),"stress_score":_advanced_score(wb,"Stress Robustness"),
        "bear_prob":_bayes_prob(wb,"Bear"),"base_prob":_bayes_prob(wb,"Base"),"bull_prob":_bayes_prob(wb,"Bull"),"wacc":wacc,"tgr":tgr,"base_10y_growth":base_10y,
        "market_base_growth":market_base,"market_implied_growth":market_implied,"market_expectation_gap":market_gap,"mc_median":mc_median,"mc_p10":mc_p10,"mc_prob_gt_price":mc_prob,"implied_fcf_growth":implied_fcf,
        **consensus,
    }


def _score_from_premium(x): return None if x is None else _clamp(50-50*x)
def _fmt_pct(v): return "N/A" if v is None else f"{v:.1%}"
def _fmt_ppt(v): return "N/A" if v is None else f"{v*100:.1f} ppt"
def _fmt_mult(v): return "N/A" if v is None else f"{v:.1f}x"
def _fmt_price(v): return "N/A" if v is None else f"${v:,.2f}"
def _fmt_bn(v): return "N/A" if v is None else f"{v:,.1f}"


def _dimension_rows(m):
    rows=[]
    ups=[x for x in (m["base_up"],m["pw_up"]) if x is not None]; avg=sum(ups)/len(ups) if ups else None
    rows.append({"name":"Absolute valuation","base_weight":30,"score":_clamp(50+100*avg) if avg is not None else None,
        "actual":f"Base {_fmt_price(m['base'])} ({_fmt_pct(m['base_up'])}); PW {_fmt_price(m['pw'])} ({_fmt_pct(m['pw_up'])})",
        "benchmark":"Current price + 15% margin-of-safety reference","gap":"N/A" if avg is None else f"{avg-.15:+.1%} vs 15%",
        "status":"Complete" if len(ups)==2 else "Partial" if ups else "Missing"})
    gaps=[x for x in (m["fpe_gap"],m["evrev_gap"],m["evebitda_gap"]) if x is not None]
    rows.append({"name":"Relative valuation","base_weight":15,"score":statistics.mean(_score_from_premium(x) for x in gaps) if gaps else None,
        "actual":f"P/E {_fmt_mult(m['fpe'])}; EV/Rev {_fmt_mult(m['evrev'])}; EV/EBITDA {_fmt_mult(m['evebitda'])}",
        "benchmark":f"Peer {_fmt_mult(m['peer_fpe'])}; {_fmt_mult(m['peer_evrev'])}; {_fmt_mult(m['peer_evebitda'])}",
        "gap":f"P/E {_fmt_pct(m['fpe_gap'])}; EV/Rev {_fmt_pct(m['evrev_gap'])}; EV/EBITDA {_fmt_pct(m['evebitda_gap'])}","status":"Complete" if len(gaps)==3 else "Partial" if gaps else "Missing"})
    gs=m["growth_score"] if m["growth_score"] is not None else (_clamp(50+(m["rev_cagr"]-.05)/.15*50) if m["rev_cagr"] is not None else None)
    rows.append({"name":"Growth","base_weight":10,"score":gs,"actual":f"5Y rev CAGR {_fmt_pct(m['rev_cagr'])}; latest {_fmt_pct(m['latest_growth'])}","benchmark":f"Peer latest growth {_fmt_pct(m['peer_growth'])}","gap":_fmt_ppt(m["growth_gap"]),"status":"Complete" if m["rev_cagr"] is not None and m["latest_growth"] is not None else "Partial" if gs is not None else "Missing"})
    ps=m["profit_score"] if m["profit_score"] is not None else (_clamp(50+(m["op_margin"]-.08)/.20*50) if m["op_margin"] is not None else None)
    rows.append({"name":"Profitability","base_weight":10,"score":ps,"actual":f"Op margin {_fmt_pct(m['op_margin'])}; ROE {_fmt_pct(m['roe'])}","benchmark":f"Peer op margin {_fmt_pct(m['peer_margin'])}; ROE {_fmt_pct(m['peer_roe'])}","gap":f"Margin {_fmt_ppt(m['margin_gap'])}; ROE {_fmt_ppt(m['roe_gap'])}","status":"Complete" if m["op_margin"] is not None and m["roe"] is not None else "Partial" if ps is not None else "Missing"})
    fs=m["fcf_score"] if m["fcf_score"] is not None else (_clamp(50+(m["fcf_margin"]-.05)/.15*50) if m["fcf_margin"] is not None else None)
    rows.append({"name":"FCF quality","base_weight":10,"score":fs,"actual":f"FCF {_fmt_bn(m['fcf'])}bn; margin {_fmt_pct(m['fcf_margin'])}","benchmark":"10% FCF-margin reference","gap":"N/A" if m["fcf_margin"] is None else f"{m['fcf_margin']-.10:+.1%} vs 10%","status":"Complete" if m["fcf_margin"] is not None else "Partial" if fs is not None else "Missing"})
    bs=m["balance_score"] if m["balance_score"] is not None else (85 if m["net_debt"] is not None and m["net_debt"]<0 else 55 if m["net_debt"] is not None else None)
    rows.append({"name":"Balance sheet","base_weight":10,"score":bs,"actual":f"Net debt/(cash) {_fmt_bn(m['net_debt'])}bn; {_fmt_pct(m['net_debt_mc'])} of market cap","benchmark":"Net debt = 0 reference","gap":"Net cash" if m["net_debt"] is not None and m["net_debt"]<0 else "Net debt" if m["net_debt"] is not None else "N/A","status":"Complete" if m["net_debt"] is not None else "Partial" if bs is not None else "Missing"})
    ss=_clamp(100*(1+m["severe_down"])) if m["severe_down"] is not None else m["stress_score"]
    rows.append({"name":"Stress resilience","base_weight":10,"score":ss,"actual":f"Severe bear {_fmt_price(m['severe'])} ({_fmt_pct(m['severe_down'])}); MC P10 {_fmt_price(m['mc_p10'])}","benchmark":"-30% severe-downside reference","gap":"N/A" if m["severe_down"] is None else f"{m['severe_down']+.30:+.1%} vs -30%","status":"Complete" if m["severe_down"] is not None else "Partial" if ss is not None else "Missing"})
    bayes=_clamp(50+(m["bull_prob"]-m["bear_prob"])*100) if m["bull_prob"] is not None and m["bear_prob"] is not None else None
    rows.append({"name":"Bayesian skew","base_weight":5,"score":bayes,"actual":f"Bull {_fmt_pct(m['bull_prob'])}; Base {_fmt_pct(m['base_prob'])}; Bear {_fmt_pct(m['bear_prob'])}","benchmark":"Bull - Bear = 0 ppt","gap":_fmt_ppt(m["bull_prob"]-m["bear_prob"] if m["bull_prob"] is not None and m["bear_prob"] is not None else None),"status":"Complete" if bayes is not None else "Missing"})
    return rows


def _investment_score(m):
    rows=_dimension_rows(m); available=sum(x["base_weight"] for x in rows if x["score"] is not None)
    for x in rows:
        if x["score"] is None or not available: x["effective_weight"]=0.0; x["contribution"]=0.0; x["impact"]=0.0
        else:
            w=x["base_weight"]/available; x["effective_weight"]=w; x["contribution"]=x["score"]*w; x["impact"]=(x["score"]-50)*w
    return (sum(x["contribution"] for x in rows) if available else None),rows,(available/100 if available else 0)


def _verdict(score,m,failed,coverage):
    if failed: return "REVIEW — DATA QUALITY"
    if score is None or coverage<.60: return "REVIEW — INSUFFICIENT DATA"
    ups=[x for x in (m["base_up"],m["pw_up"]) if x is not None]; avg=sum(ups)/len(ups) if ups else None
    if avg is not None and avg>=.20 and score>=65: return "ATTRACTIVE"
    if avg is not None and avg>=.10 and score>=55: return "POTENTIALLY ATTRACTIVE"
    if (avg is None or avg>-.10) and score>=45: return "NEUTRAL / WATCHLIST"
    return "UNATTRACTIVE AT CURRENT PRICE"

def _why_line(score,m,coverage):
    p=[]
    if m["pw_up"] is not None: p.append(f"PW value {_fmt_pct(m['pw_up'])} vs price")
    elif m["base_up"] is not None: p.append(f"Base DCF {_fmt_pct(m['base_up'])} vs price")
    if score is not None: p.append(f"quant score {score:.1f}/100 with {coverage:.0%} score-weight coverage")
    if m["severe_down"] is not None: p.append(f"severe-bear {_fmt_pct(m['severe_down'])}")
    if m["mc_prob_gt_price"] is not None: p.append(f"MC P(value > price) {_fmt_pct(m['mc_prob_gt_price'])}")
    return "; ".join(p)+"." if p else "Not enough validated numeric evidence for a concise conclusion."


def _strengths_risks(m,rows):
    strengths=[]; risks=[]
    if m["base_up"] is not None: (strengths if m["base_up"]>=.15 else risks if m["base_up"]<0 else strengths).append(f"Base DCF implies {_fmt_pct(m['base_up'])} upside/downside.")
    if m["pw_up"] is not None: (strengths if m["pw_up"]>=.15 else risks if m["pw_up"]<0 else strengths).append(f"Probability-weighted value implies {_fmt_pct(m['pw_up'])} upside/downside.")
    if m["rev_cagr"] is not None:
        t=f"Revenue CAGR {_fmt_pct(m['rev_cagr'])}"
        if m["latest_growth"] is not None and m["peer_growth"] is not None: t+=f"; latest {_fmt_pct(m['latest_growth'])} vs peer {_fmt_pct(m['peer_growth'])}."
        else: t+="."
        (strengths if m["rev_cagr"]>=.08 else risks if m["rev_cagr"]<.03 else strengths).append(t)
    if m["op_margin"] is not None:
        t=f"Operating margin {_fmt_pct(m['op_margin'])}"+(f" vs peer {_fmt_pct(m['peer_margin'])} ({_fmt_ppt(m['margin_gap'])})." if m["peer_margin"] is not None else ".")
        (strengths if m["op_margin"]>=.15 else risks if m["op_margin"]<.05 else strengths).append(t)
    if m["fcf_margin"] is not None: (strengths if m["fcf_margin"]>=.10 else risks if m["fcf_margin"]<.04 else strengths).append(f"Latest FCF margin {_fmt_pct(m['fcf_margin'])}.")
    if m["fpe_gap"] is not None: (strengths if m["fpe_gap"]<=-.10 else risks if m["fpe_gap"]>=.20 else strengths).append(f"Forward P/E is {abs(m['fpe_gap']):.1%} {'below' if m['fpe_gap']<0 else 'above'} the direct/exact peer median.")
    if m["net_debt"] is not None and m["net_debt"]<0: strengths.append(f"Net cash is {abs(m['net_debt']):,.1f}bn.")
    if m["severe_down"] is not None and m["severe_down"]<=-.50: risks.append(f"Severe-bear downside is {_fmt_pct(m['severe_down'])}.")
    if m["market_expectation_gap"] is not None:
        if m["market_expectation_gap"]>.05: risks.append(f"Market-implied 10Y revenue growth is {_fmt_ppt(m['market_expectation_gap'])} above the model Base case.")
        elif m["market_expectation_gap"]<-.03: strengths.append(f"Market-implied 10Y revenue growth is {_fmt_ppt(abs(m['market_expectation_gap']))} below the model Base case.")
    if m["revenue_gap"] is not None:
        (strengths if m["revenue_gap"]>.05 else risks if m["revenue_gap"]<-.05 else strengths).append(f"Next-year model revenue is {_fmt_pct(m['revenue_gap'])} vs public consensus.")
    if m["mc_prob_gt_price"] is not None: (strengths if m["mc_prob_gt_price"]>=.60 else risks if m["mc_prob_gt_price"]<.40 else strengths).append(f"Monte Carlo P(value > price) is {_fmt_pct(m['mc_prob_gt_price'])}.")
    ranked=[x for x in rows if x["score"] is not None]
    if ranked:
        best=max(ranked,key=lambda x:x["impact"]); worst=min(ranked,key=lambda x:x["impact"])
        if best["impact"]>1: strengths.append(f"Largest score support: {best['name']} adds {best['impact']:+.1f} points vs neutral.")
        if worst["impact"]<-1: risks.append(f"Largest score drag: {worst['name']} subtracts {abs(worst['impact']):.1f} points vs neutral.")
    return strengths[:7],risks[:7]


def _write_pair(ws,lc,vc,label,value,fmt=None,fill=None,bold=False):
    ws[lc]=label; ws[lc].font=Font(bold=True,color=GREY); ws[vc]=value
    if fmt: ws[vc].number_format=fmt
    if fill: ws[vc].fill=_fill(fill)
    if bold: ws[vc].font=Font(bold=True)


def ensure_investment_summary(wb,ticker):
    if "Investment Summary" in wb.sheetnames: wb.remove(wb["Investment Summary"])
    ws=wb.create_sheet("Investment Summary",1); _style_title(ws,f"{ticker} — Investment Summary & Decision View")
    ws["A3"]="Decision-focused synthesis of valuation, peer, operating, stress, expectations and scenario outputs. All score dimensions stay visible; missing inputs are marked rather than silently dropped."
    ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True); ws.merge_cells("A3:J3")
    m=_metric_snapshot(wb); passed,review,failed=_data_quality(wb); score,score_rows,coverage=_investment_score(m); verdict=_verdict(score,m,failed,coverage)

    _section(ws,5,"Investment Decision")
    ws["A6"]="Model View"; ws["B6"]=verdict; ws["B6"].fill=_fill(PALE_GREEN if verdict=="ATTRACTIVE" else PALE_RED if verdict.startswith("UNATTRACTIVE") else PALE_YELLOW); ws["B6"].font=Font(bold=True,size=14); ws.merge_cells("B6:C6")
    color=PALE_GREEN if score is not None and score>=65 else PALE_RED if score is not None and score<40 else PALE_YELLOW if score is not None else LIGHT
    _write_pair(ws,"D6","E6","Quant Score / 100",score,FMT_SCORE,color,True); _write_pair(ws,"F6","G6","Score Coverage",coverage,FMT_PCT,PALE_GREEN if coverage>=.85 else PALE_YELLOW if coverage>=.60 else PALE_RED,True)
    ws["H6"]="Data Quality"; ws["I6"]=f"{passed} PASS / {review} REVIEW / {failed} FAIL"; ws.merge_cells("I6:J6")
    ws["A7"]="Why"; ws["B7"]=_why_line(score,m,coverage); ws["B7"].alignment=Alignment(wrap_text=True); ws.merge_cells("B7:J7")
    _write_pair(ws,"A8","B8","Current Price",m["price"],FMT_PRICE); _write_pair(ws,"C8","D8","Base DCF / Share",m["base"],FMT_PRICE); _write_pair(ws,"E8","F8","Base Upside",m["base_up"],FMT_PCT); _write_pair(ws,"G8","H8","PW Value / Share",m["pw"],FMT_PRICE); _write_pair(ws,"I8","J8","PW Upside",m["pw_up"],FMT_PCT)
    _write_pair(ws,"A9","B9","Severe Bear / Share",m["severe"],FMT_PRICE); _write_pair(ws,"C9","D9","Severe Downside",m["severe_down"],FMT_PCT); _write_pair(ws,"E9","F9","MC Median / Share",m["mc_median"],FMT_PRICE); _write_pair(ws,"G9","H9","MC P(Value > Price)",m["mc_prob_gt_price"],FMT_PCT); _write_pair(ws,"I9","J9","Market Cap (bn)",m["market_cap"],FMT_BN)
    _write_pair(ws,"A10","B10","Net Debt / (Cash) (bn)",m["net_debt"],FMT_BN); _write_pair(ws,"C10","D10","WACC",m["wacc"],FMT_PCT); _write_pair(ws,"E10","F10","Terminal Growth",m["tgr"],FMT_PCT); ws["G10"]="Sector / Industry"; ws["G10"].font=Font(bold=True,color=GREY); ws["H10"]=f"{m['sector'] or 'Unknown'} / {m['industry'] or 'Unknown'}"; ws.merge_cells("H10:J10")

    _section(ws,12,"Key Numbers Behind the View"); _header(ws,13,["Metric","Company / Model","Peer / Reference","Gap / Skew","Signal","Context","","","",""])
    key_rows=[
        ("5Y Revenue CAGR",m["rev_cagr"],None,None,"Historical growth","Compounded across available annual history",FMT_PCT),
        ("Latest Revenue Growth",m["latest_growth"],m["peer_growth"],m["growth_gap"],"vs direct/exact peers","Public latest growth from Peer Comps",FMT_PCT),
        ("Operating Margin",m["op_margin"],m["peer_margin"],m["margin_gap"],"vs direct/exact peers","Latest reported operating margin",FMT_PCT),
        ("ROE",m["roe"],m["peer_roe"],m["roe_gap"],"vs direct/exact peers","Public ROE / statement fallback",FMT_PCT),
        ("FCF Margin",m["fcf_margin"],.10,m["fcf_margin"]-.10 if m["fcf_margin"] is not None else None,"10% reference","OCF less capex / revenue",FMT_PCT),
        ("Forward P/E",m["fpe"],m["peer_fpe"],m["fpe_gap"],"vs direct/exact peers","Relative valuation",FMT_MULT),
        ("EV / Revenue",m["evrev"],m["peer_evrev"],m["evrev_gap"],"vs direct/exact peers","Useful when margin structures are comparable",FMT_MULT),
        ("EV / EBITDA",m["evebitda"],m["peer_evebitda"],m["evebitda_gap"],"vs direct/exact peers","Capital-structure-neutral valuation",FMT_MULT),
        ("Base 10Y Revenue CAGR",m["base_10y_growth"],m["market_implied_growth"],m["base_10y_growth"]-m["market_implied_growth"] if m["base_10y_growth"] is not None and m["market_implied_growth"] is not None else None,"model vs market-implied","Expectations gap",FMT_PCT),
        ("Implied 10Y FCF CAGR",m["implied_fcf_growth"],None,None,"reverse DCF","Growth required by current enterprise value",FMT_PCT),
        ("Next-Year Revenue vs Consensus",m["revenue_model"],m["revenue_consensus"],m["revenue_gap"],"variant perception","Public consensus where available",FMT_BN),
        ("Monte Carlo P10 / Median",m["mc_p10"],m["mc_median"],None,"valuation distribution",f"P(value > current) {_fmt_pct(m['mc_prob_gt_price'])}",FMT_PRICE),
    ]
    for r,(lab,val,bench,gap,signal,ctx,fmt) in enumerate(key_rows,14):
        ws.cell(r,1,lab); ws.cell(r,2,val); ws.cell(r,3,bench); ws.cell(r,4,gap); ws.cell(r,5,signal); ws.cell(r,6,ctx); ws.cell(r,2).number_format=fmt; ws.cell(r,3).number_format=fmt; ws.cell(r,4).number_format=FMT_PCT; ws.cell(r,6).alignment=Alignment(wrap_text=True)

    bridge=28; _section(ws,bridge,"Quantitative Score Bridge — Neutral 50 to Final Score")
    ws.cell(bridge+1,1,"Start at neutral"); ws.cell(bridge+1,2,50.0); ws.cell(bridge+1,2).number_format=FMT_SCORE; ws.cell(bridge+1,3,"Each available dimension adds or subtracts points versus neutral 50. Missing dimensions stay visible with 0% effective weight."); ws.merge_cells(start_row=bridge+1,start_column=3,end_row=bridge+1,end_column=10); ws.cell(bridge+1,3).font=Font(italic=True,color=GREY); ws.cell(bridge+1,3).alignment=Alignment(wrap_text=True)
    hdr=bridge+2; _header(ws,hdr,["Dimension","Score / 100","Base Weight","Effective Weight","Contribution","Impact vs Neutral","Key Actuals","Peer / Reference","Gap / Skew","Data Status"])
    first=hdr+1
    for r,x in enumerate(score_rows,first):
        vals=[x["name"],x["score"],x["base_weight"]/100,x["effective_weight"],x["contribution"],x["impact"],x["actual"],x["benchmark"],x["gap"],x["status"]]
        for c,v in enumerate(vals,1): ws.cell(r,c,v); ws.cell(r,c).border=Border(bottom=THIN)
        ws.cell(r,2).number_format=FMT_SCORE; ws.cell(r,3).number_format=FMT_PCT; ws.cell(r,4).number_format=FMT_PCT; ws.cell(r,5).number_format=FMT_SCORE; ws.cell(r,6).number_format='+0.0;[Red]-0.0;-'; ws.cell(r,2).fill=_score_fill(x["score"]); ws.cell(r,10).fill=_fill(PALE_RED if x["status"]=="Missing" else PALE_YELLOW if x["status"]=="Partial" else PALE_GREEN)
        for c in (7,8,9,10): ws.cell(r,c).alignment=Alignment(wrap_text=True,vertical="top")
    total=first+len(score_rows); impact=sum(x["impact"] for x in score_rows)
    vals=["Final Quantitative Score",score,1.0,sum(x["effective_weight"] for x in score_rows),sum(x["contribution"] for x in score_rows),impact,f"50 neutral + {impact:+.1f} dimension impact","","",f"{coverage:.0%} score-weight coverage"]
    for c,v in enumerate(vals,1): ws.cell(total,c,v); ws.cell(total,c).font=Font(bold=True); ws.cell(total,c).border=Border(top=MEDIUM)
    ws.cell(total,2).number_format=FMT_SCORE; ws.cell(total,3).number_format=FMT_PCT; ws.cell(total,4).number_format=FMT_PCT; ws.cell(total,5).number_format=FMT_SCORE; ws.cell(total,6).number_format='+0.0;[Red]-0.0;-'; ws.cell(total,2).fill=_score_fill(score)
    note=total+1; ws.cell(note,1,"Score interpretation"); ws.cell(note,2,"80–100 very strong | 65–79 favorable | 50–64 mixed-positive | 35–49 weak | below 35 poor. The score is a decision aid; valuation and thesis evidence still require analyst judgment."); ws.merge_cells(start_row=note,start_column=2,end_row=note,end_column=10); ws.cell(note,2).font=Font(italic=True,color=GREY); ws.cell(note,2).alignment=Alignment(wrap_text=True)

    sr=note+2; _section(ws,sr,"What the Numbers Like"); strengths,risks=_strengths_risks(m,score_rows)
    for r,text in enumerate(strengths or ["No major quantitative strength passed the current rule thresholds."],sr+1): ws.cell(r,1,"+"); ws.cell(r,2,text); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=10); ws.cell(r,2).alignment=Alignment(wrap_text=True)
    rr=sr+max(4,len(strengths)+1)+1; _section(ws,rr,"Biggest Quantitative Risks / What Must Go Right")
    for r,text in enumerate(risks or ["No major quantitative risk passed the current rule thresholds; qualitative risks still require analyst review."],rr+1): ws.cell(r,1,"!"); ws.cell(r,2,text); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=10); ws.cell(r,2).alignment=Alignment(wrap_text=True)
    foot=rr+max(4,len(risks)+1)+1; ws.cell(foot,1,"Interpretation"); ws.cell(foot,2,"Re-underwrite when earnings, consensus, guidance, capital allocation, sector conditions or the share price change. Missing data should be filled at the source rather than overridden in this summary."); ws.merge_cells(start_row=foot,start_column=2,end_row=foot,end_column=10); ws.cell(foot,2).font=Font(italic=True,color=GREY); ws.cell(foot,2).alignment=Alignment(wrap_text=True)

    widths={"A":29,"B":17,"C":17,"D":17,"E":18,"F":18,"G":38,"H":34,"I":28,"J":18}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    for r in range(1,ws.max_row+1): ws.row_dimensions[r].height=20
    ws.row_dimensions[3].height=34; ws.row_dimensions[7].height=34
    for r in range(first,total): ws.row_dimensions[r].height=48
    ws.freeze_panes="A12"
    return ws
