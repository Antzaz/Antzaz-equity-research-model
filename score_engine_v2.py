from __future__ import annotations

"""Single-source, auditable investment scoring engine.

The old model used absolute 40% operating-margin and 25% FCF-margin hurdles for every
business. That systematically understated high-quality low-margin business models such as
warehouse retail. This engine separates business quality from stock valuation and uses
business-model-aware/peer-relative economics wherever practical.

All scores are 0-100 and every dimension carries its actual inputs, benchmark, formula,
source cells and limitations. No score is intended to be a factual rating from an external
institution; it is a transparent project calculation.
"""

import math
import statistics
from typing import Any

DIMENSION_WEIGHTS={
    "Absolute Valuation":30.0,
    "Relative Valuation":15.0,
    "Growth":10.0,
    "Profitability":10.0,
    "FCF Quality":10.0,
    "Balance Sheet":10.0,
    "Stress Robustness":10.0,
    "Bayesian Skew":5.0,
}

QUALITY_WEIGHTS={"Growth":.20,"Profitability":.35,"FCF Quality":.30,"Balance Sheet":.15}
VALUATION_WEIGHTS={"Absolute Valuation":.65,"Relative Valuation":.35}


def num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        x=float(v)
        return x if math.isfinite(x) else default
    except Exception: return default


def clamp(v,lo=0.0,hi=100.0):
    return max(lo,min(hi,float(v)))


def weighted_available(parts: list[tuple[float|None,float]]) -> float|None:
    used=[(v,w) for v,w in parts if v is not None and w>0]
    if not used: return None
    total=sum(w for _,w in used)
    return sum(v*w for v,w in used)/total


def _find_row(ws,label,start=1,end=None):
    if ws is None: return None
    needle=str(label).strip().lower(); end=end or ws.max_row
    for r in range(start,min(end,ws.max_row)+1):
        if str(ws.cell(r,1).value or "").strip().lower()==needle: return r
    return None


def _history(wb):
    out={"years":[],"revenue":[],"op":[],"ni":[],"ocf":[],"capex":[]}
    if "Historical Financials" not in wb.sheetnames: return out
    ws=wb["Historical Financials"]
    for c in range(2,8):
        y=ws.cell(3,c).value; rev=num(ws.cell(4,c).value)
        if not isinstance(y,(int,float)) or rev is None: continue
        out["years"].append(int(y)); out["revenue"].append(rev)
        out["op"].append(num(ws.cell(9,c).value)); out["ni"].append(num(ws.cell(11,c).value))
        out["ocf"].append(num(ws.cell(14,c).value)); out["capex"].append(num(ws.cell(15,c).value))
    if len(out["years"])>=2 and out["revenue"][0]>0 and out["revenue"][-1]>0:
        n=max(1,out["years"][-1]-out["years"][0])
        out["revenue_cagr"]=(out["revenue"][-1]/out["revenue"][0])**(1/n)-1
        out["latest_revenue_growth"]=out["revenue"][-1]/out["revenue"][-2]-1 if out["revenue"][-2] else None
    else:
        out["revenue_cagr"]=None; out["latest_revenue_growth"]=None
    rev=out["revenue"][-1] if out["revenue"] else None
    op=out["op"][-1] if out["op"] else None; ni=out["ni"][-1] if out["ni"] else None
    ocf=out["ocf"][-1] if out["ocf"] else None; cap=out["capex"][-1] if out["capex"] else None
    fcf=(ocf-abs(cap)) if ocf is not None and cap is not None else None
    out.update({
        "operating_margin":op/rev if rev not in (None,0) and op is not None else None,
        "fcf":fcf,
        "fcf_margin":fcf/rev if rev not in (None,0) and fcf is not None else None,
        "fcf_to_net_income":fcf/ni if ni not in (None,0) and fcf is not None else None,
        "ocf_to_net_income":ocf/ni if ni not in (None,0) and ocf is not None else None,
    })
    valid_fcf=[]
    for o,c in zip(out["ocf"],out["capex"]):
        if o is not None and c is not None: valid_fcf.append(o-abs(c))
    out["fcf_positive_ratio"]=(sum(v>0 for v in valid_fcf)/len(valid_fcf)) if valid_fcf else None
    return out


def _peer_snapshot(wb):
    out={"target":{},"peer":{},"target_row":None,"peer_rows":[]}
    if "Peer Comps" not in wb.sheetnames: return out
    ws=wb["Peer Comps"]; headers={str(ws.cell(3,c).value or "").strip():c for c in range(1,ws.max_column+1)}
    metrics=("Forward P/E","EV/Revenue","EV/EBITDA","Revenue Growth","Operating Margin","ROE")
    direct=[]; all_peers=[]
    for r in range(4,min(ws.max_row,40)+1):
        ticker=str(ws.cell(r,headers.get("Ticker",2)).value or "").strip()
        if not ticker: continue
        typ=str(ws.cell(r,headers.get("Peer Type",16)).value or "").strip()
        if typ=="Target classification": out["target_row"]=r
        else:
            all_peers.append(r)
            if typ in {"Direct business-model peer","Exact industry"}: direct.append(r)
    if out["target_row"] is None and str(ws.cell(4,headers.get("Ticker",2)).value or "").strip(): out["target_row"]=4
    rows=direct or all_peers; out["peer_rows"]=rows
    if out["target_row"]:
        for m in metrics:
            if m in headers: out["target"][m]=num(ws.cell(out["target_row"],headers[m]).value)
    for m in metrics:
        if m not in headers: continue
        vals=[num(ws.cell(r,headers[m]).value) for r in rows]; vals=[v for v in vals if v is not None]
        out["peer"][m]=statistics.median(vals) if vals else None
    return out


def _section_bounds(ws,title,next_title=None):
    start=_find_row(ws,title)
    if start is None: return None,None
    end=ws.max_row
    if next_title:
        n=_find_row(ws,next_title,start+1)
        if n: end=n-1
    return start,end


def _statement_latest(wb):
    """Return ROIC inputs from Financial Statements when enough reported balance-sheet data exists."""
    out={"roic":None,"tax_rate":None,"invested_capital":None,"source":"Financial Statements"}
    if "Financial Statements" not in wb.sheetnames: return out
    ws=wb["Financial Statements"]
    i0,i1=_section_bounds(ws,"Income Statement","Balance Sheet")
    b0,b1=_section_bounds(ws,"Balance Sheet","Cash Flow Statement")
    if i0 is None or b0 is None: return out
    rev_r=_find_row(ws,"Revenue",i0,i1); op_r=_find_row(ws,"Operating Income",i0,i1)
    pretax_r=_find_row(ws,"Pre-Tax Income",i0,i1); tax_r=_find_row(ws,"Income Taxes",i0,i1)
    income_header=i0+1
    year_cols=[c for c in range(2,min(8,ws.max_column+1)) if isinstance(ws.cell(income_header,c).value,(int,float))]
    if not year_cols or op_r is None: return out
    c=year_cols[-1]; op=num(ws.cell(op_r,c).value); pretax=num(ws.cell(pretax_r,c).value) if pretax_r else None; taxes=num(ws.cell(tax_r,c).value) if tax_r else None
    tax_rate=(taxes/pretax) if pretax not in (None,0) and taxes is not None else None
    if tax_rate is None or not 0<=tax_rate<=.50: tax_rate=.21
    out["tax_rate"]=tax_rate

    header=None
    for r in range(b0,min(b1,b0+4)+1):
        if str(ws.cell(r,1).value or "").strip().lower()=="metric": header=r; break
    if header is None: header=b0+1
    bs_cols=[c for c in range(2,min(8,ws.max_column+1)) if isinstance(ws.cell(header,c).value,(int,float))]
    eq_r=_find_row(ws,"Stockholders' Equity",b0,b1); debt_r=_find_row(ws,"Long-Term Debt",b0,b1); cash_r=_find_row(ws,"Cash & Cash Equivalents",b0,b1)
    ics=[]
    for bc in bs_cols[-2:]:
        eq=num(ws.cell(eq_r,bc).value) if eq_r else None; debt=num(ws.cell(debt_r,bc).value,0) if debt_r else 0; cash=num(ws.cell(cash_r,bc).value,0) if cash_r else 0
        if eq is not None: ics.append(eq+(debt or 0)-(cash or 0))
    avg_ic=statistics.mean(ics) if ics and all(x>0 for x in ics) else None
    nopat=op*(1-tax_rate) if op is not None else None
    if avg_ic not in (None,0) and nopat is not None:
        out["invested_capital"]=avg_ic; out["roic"]=nopat/avg_ic
    return out


def _data_lineage(wb):
    if "Historical Financials" not in wb.sheetnames: return "Workbook historical financials"
    ws=wb["Historical Financials"]
    for r in range(23,min(ws.max_row,30)+1):
        label=str(ws.cell(r,1).value or "").lower()
        if "source" in label:
            v=str(ws.cell(r,2).value or "").strip()
            if v: return v
    return "Historical Financials / source hierarchy in workbook"


def _score_growth(hist,peer):
    cagr=hist.get("revenue_cagr"); latest=hist.get("latest_revenue_growth"); p=peer.get("Revenue Growth")
    hist_score=clamp(50+((cagr-.05)/.10)*50) if cagr is not None else None
    peer_score=clamp(50+((latest-p)/.10)*50) if latest is not None and p is not None else None
    score=weighted_available([(hist_score,.70),(peer_score,.30)])
    return score,hist_score,peer_score


def _score_profitability(hist,peer,stmt,wacc):
    roic=stmt.get("roic"); roe=peer.get("_target_roe"); proe=peer.get("ROE"); opm=hist.get("operating_margin"); popm=peer.get("Operating Margin")
    roic_score=clamp(50+((roic-wacc)/.15)*50) if roic is not None and wacc is not None else None
    roe_score=clamp(50+((roe-proe)/.20)*50) if roe is not None and proe is not None else None
    margin_score=clamp(50+((opm-popm)/max(abs(popm),.05))*50) if opm is not None and popm is not None else None
    score=weighted_available([(roic_score,.55),(roe_score,.25),(margin_score,.20)])
    return score,roic_score,roe_score,margin_score


def _score_fcf(hist):
    conv=hist.get("fcf_to_net_income"); ocf=hist.get("ocf_to_net_income"); positive=hist.get("fcf_positive_ratio")
    conv_score=clamp(50+((conv-.70)/.50)*50) if conv is not None else None
    ocf_score=clamp(50+((ocf-1.00)/.50)*50) if ocf is not None else None
    positive_score=clamp(50+((positive-.70)/.30)*50) if positive is not None else None
    score=weighted_available([(conv_score,.45),(ocf_score,.30),(positive_score,.25)])
    return score,conv_score,ocf_score,positive_score


def _score_balance(wb,hist):
    if "Company Data" not in wb.sheetnames: return None,None
    ws=wb["Company Data"]; cash=num(ws["B12"].value); debt=num(ws["B13"].value); net=(debt-cash) if debt is not None and cash is not None else num(ws["B14"].value)
    fcf=hist.get("fcf")
    if net is None: return None,None
    if net<=0:
        ratio=abs(net)/(abs(fcf) if fcf not in (None,0) else 1)
        return clamp(85+min(15,ratio*10)),net
    leverage=net/(abs(fcf) if fcf not in (None,0) else max(abs(net),1))
    return clamp(85-15*leverage),net


def _score_absolute(price,base):
    if price in (None,0) or base is None: return None,None
    upside=base/price-1
    return clamp(50+100*upside),upside


def _score_relative(target,peer):
    pieces=[]; detail={}
    for key in ("Forward P/E","EV/Revenue","EV/EBITDA"):
        v=target.get(key); p=peer.get(key)
        if v is None or p in (None,0): continue
        premium=v/p-1; s=clamp(50-50*premium); pieces.append(s); detail[key]=(v,p,premium,s)
    return (statistics.mean(pieces) if pieces else None),detail


def _score_stress(price,severe_value,mc_prob):
    if price in (None,0): return None,None,None
    severe=max(0.0,severe_value) if severe_value is not None else None
    downside=severe/price-1 if severe is not None else None
    downside_score=clamp(80+100*downside) if downside is not None else None  # 80=no loss, 50=-30%, 0=-80% or worse
    mc_score=clamp(mc_prob*100) if mc_prob is not None else None
    return weighted_available([(downside_score,.65),(mc_score,.35)]),downside,downside_score


def _bayes(wb):
    if "Base Rates & Probabilities" not in wb.sheetnames: return None,{}
    ws=wb["Base Rates & Probabilities"]; vals={}
    for name in ("Bear","Base","Bull"):
        r=_find_row(ws,name); vals[name]=num(ws.cell(r,4).value) if r else None
    if vals.get("Bull") is None or vals.get("Bear") is None: return None,vals
    return clamp(50+(vals["Bull"]-vals["Bear"])*100),vals


def compute_score_bundle(wb,ticker:str|None=None,base_value=None,severe_value=None,current_price=None,mc_prob=None):
    hist=_history(wb); peers=_peer_snapshot(wb); target=peers["target"]; peer=peers["peer"]
    peer["_target_roe"]=target.get("ROE")
    stmt=_statement_latest(wb)
    d=wb["Company Data"] if "Company Data" in wb.sheetnames else None
    price=num(current_price) if current_price is not None else (num(d["B8"].value) if d else None)
    s=wb["Three-Case Scenarios"] if "Three-Case Scenarios" in wb.sheetnames else None
    wacc=num(s["C6"].value,.09) if s else .09
    if base_value is None and "Dashboard" in wb.sheetnames:
        # Only use a materialized numeric value. Callers that have the deterministic DCF should pass it explicitly.
        for r in range(1,wb["Dashboard"].max_row+1):
            if str(wb["Dashboard"].cell(r,1).value or "").strip()=="Base DCF Value / Share":
                base_value=num(wb["Dashboard"].cell(r,2).value); break
    if severe_value is None and "Three-Case Scenarios" in wb.sheetnames:
        r=_find_row(wb["Three-Case Scenarios"],"Severe Bear Value / Share")
        if r: severe_value=num(wb["Three-Case Scenarios"].cell(r,2).value)
    if mc_prob is None and "Advanced Analytics" in wb.sheetnames:
        ws=wb["Advanced Analytics"]
        for r in range(1,ws.max_row+1):
            if str(ws.cell(r,9).value or "").strip()=="Probability > Current Price": mc_prob=num(ws.cell(r,10).value); break

    growth,gh,gp=_score_growth(hist,peer)
    profitability,pr,proe,pm=_score_profitability(hist,peer,stmt,wacc)
    fcf,fc,fo,fp=_score_fcf(hist)
    balance,net_debt=_score_balance(wb,hist)
    absolute,upside=_score_absolute(price,num(base_value))
    relative,rel_detail=_score_relative(target,peer)
    stress,downside,downside_score=_score_stress(price,num(severe_value),num(mc_prob))
    bayes,bayes_vals=_bayes(wb)
    lineage=_data_lineage(wb)

    dims={
        "Growth":{"score":growth,"category":"Business Quality","actual":f"Revenue CAGR={hist.get('revenue_cagr')}; latest annual growth={hist.get('latest_revenue_growth')}","benchmark":f"5% long-run anchor; direct/exact peer growth={peer.get('Revenue Growth')}","formula":"70% × clamp(50 + (Revenue CAGR − 5%)/10% × 50) + 30% × clamp(50 + (latest annual growth − peer growth)/10% × 50); available components reweighted.","components":f"historical={gh}; peer-relative={gp}","source_cells":"Historical Financials!B4:G4; Peer Comps!F","source":lineage,"status":"Complete" if growth is not None and gh is not None and gp is not None else "Partial" if growth is not None else "Missing"},
        "Profitability":{"score":profitability,"category":"Business Quality","actual":f"ROIC={stmt.get('roic')}; ROE={target.get('ROE')}; operating margin={hist.get('operating_margin')}","benchmark":f"WACC={wacc}; peer ROE={peer.get('ROE')}; peer operating margin={peer.get('Operating Margin')}","formula":"55% × ROIC-vs-WACC score + 25% × ROE-vs-peer score + 20% × operating-margin-vs-peer score. ROIC score=clamp(50+(ROIC−WACC)/15ppt×50); ROE score=clamp(50+(ROE−peer ROE)/20ppt×50); margin score is relative to peer margin. Available components reweighted.","components":f"ROIC={pr}; ROE={proe}; margin={pm}","source_cells":"Financial Statements (Operating Income, taxes, equity, debt, cash); Peer Comps!G:H; Three-Case Scenarios!C6","source":lineage,"status":"Complete" if profitability is not None and pr is not None and proe is not None and pm is not None else "Partial" if profitability is not None else "Missing"},
        "FCF Quality":{"score":fcf,"category":"Business Quality","actual":f"FCF/Net Income={hist.get('fcf_to_net_income')}; OCF/Net Income={hist.get('ocf_to_net_income')}; positive-FCF years={hist.get('fcf_positive_ratio')}","benchmark":"70% FCF/NI neutral anchor; 100% OCF/NI neutral anchor; 70% positive-FCF-years neutral anchor","formula":"45% × cash-conversion score + 30% × OCF/NI score + 25% × positive-FCF-history score. This measures cash conversion and consistency, not an industry-invariant FCF margin hurdle.","components":f"FCF conversion={fc}; OCF conversion={fo}; positive history={fp}","source_cells":"Historical Financials!B11:G16","source":lineage,"status":"Complete" if fcf is not None and fc is not None and fo is not None and fp is not None else "Partial" if fcf is not None else "Missing"},
        "Balance Sheet":{"score":balance,"category":"Business Quality","actual":f"Net debt/(cash)={net_debt}; latest FCF={hist.get('fcf')}","benchmark":"Net cash scores ≥85; net debt is scaled to FCF","formula":"If net cash: clamp(85 + min(15, |net cash|/FCF × 10)); if net debt: clamp(85 − 15 × net debt/FCF).","components":"Net cash/debt and FCF capacity","source_cells":"Company Data!B12:B14; Historical Financials!G14:G15","source":lineage,"status":"Complete" if balance is not None else "Missing"},
        "Absolute Valuation":{"score":absolute,"category":"Valuation / Stock Attractiveness","actual":f"Current price={price}; base intrinsic value={base_value}; upside={upside}","benchmark":"50 = base intrinsic value equals current price","formula":"clamp(50 + 100 × (Base DCF / Current Price − 1)). Thus +20% upside=70, fair value=50, −20% downside=30.","components":f"base upside={upside}","source_cells":"Company Data!B8; deterministic Base DCF from Three-Case Scenarios/Advanced Analytics","source":"Project deterministic DCF engine + market price","status":"Complete" if absolute is not None else "Missing"},
        "Relative Valuation":{"score":relative,"category":"Valuation / Stock Attractiveness","actual":"; ".join(f"{k}={v[0]}" for k,v in rel_detail.items()),"benchmark":"; ".join(f"peer {k}={v[1]}" for k,v in rel_detail.items()),"formula":"For each available lower-is-better multiple: clamp(50 − 50 × (Company/PeerMedian − 1)); score = mean of P/E, EV/Revenue and EV/EBITDA component scores.","components":"; ".join(f"{k}: premium={v[2]}, score={v[3]}" for k,v in rel_detail.items()),"source_cells":"Peer Comps!C:E, direct/exact peer median preferred","source":"Peer Comps provider/fallback notes","status":"Complete" if relative is not None and len(rel_detail)==3 else "Partial" if relative is not None else "Missing"},
        "Stress Robustness":{"score":stress,"category":"Risk / Valuation Stress","actual":f"Severe-bear value={max(0,num(severe_value,0)) if severe_value is not None else None}; severe downside={downside}; MC P(value>price)={mc_prob}","benchmark":"Downside score: 80 at no loss, 50 at −30%, 0 at −80% or worse","formula":"65% × clamp(80 + 100 × severe downside) + 35% × Monte Carlo P(value > price)×100; available components reweighted. Severe equity value is floored at zero. This is valuation downside stress, not operating-business resilience.","components":f"downside component={downside_score}; MC component={num(mc_prob)*100 if num(mc_prob) is not None else None}","source_cells":"Three-Case Scenarios severe bear; Advanced Analytics Monte Carlo; Company Data!B8","source":"Project deterministic stress DCF + regime Monte Carlo","status":"Complete" if stress is not None and downside is not None and mc_prob is not None else "Partial" if stress is not None else "Missing"},
        "Bayesian Skew":{"score":bayes,"category":"Risk / Scenario Skew","actual":f"Bull={bayes_vals.get('Bull')}; Base={bayes_vals.get('Base')}; Bear={bayes_vals.get('Bear')}","benchmark":"Bull probability − Bear probability = 0ppt gives 50","formula":"clamp(50 + 100 × (Bull probability − Bear probability)).","components":"Scenario probability skew","source_cells":"Base Rates & Probabilities scenario probability table","source":"Project Bayesian/base-rate layer","status":"Complete" if bayes is not None else "Missing"},
    }

    quality=weighted_available([(dims[k]["score"],w) for k,w in QUALITY_WEIGHTS.items()])
    valuation=weighted_available([(dims[k]["score"],w) for k,w in VALUATION_WEIGHTS.items()])
    risk=weighted_available([(stress,.70),(bayes,.30)])
    available=sum(DIMENSION_WEIGHTS[k] for k,d in dims.items() if d["score"] is not None)
    overall=(sum(dims[k]["score"]*DIMENSION_WEIGHTS[k] for k in DIMENSION_WEIGHTS if dims[k]["score"] is not None)/available) if available else None
    return {
        "ticker":ticker,
        "dimensions":dims,
        "category_scores":{"Business Quality":quality,"Valuation / Stock Attractiveness":valuation,"Downside / Scenario Risk":risk,"Overall Investment Score":overall},
        "coverage":available/100 if available else 0,
        "weights":DIMENSION_WEIGHTS.copy(),
        "method_version":"Score Engine v2 — sector-aware economics / single source of truth",
    }


def advanced_scorecard(wb,current_price,forward_pe,base_value,severe_value):
    """Drop-in replacement for advanced_analytics_v2._scorecard."""
    bundle=compute_score_bundle(wb,base_value=base_value,severe_value=severe_value,current_price=current_price)
    order=("Growth","Profitability","FCF Quality","Balance Sheet","Absolute Valuation","Relative Valuation","Stress Robustness")
    return [(k,bundle["dimensions"][k]["score"],bundle["dimensions"][k]["formula"]+" See Score Audit Trail for inputs and sources.") for k in order]
