"""Model-driven investment summary for the equity research workbook.

The sheet is deliberately auditable: the verdict is derived from workbook outputs rather
than an opaque recommendation model. Missing inputs are excluded and the remaining weights
are renormalized. This is a research decision aid, not personalized investment advice.
"""

import math
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; LIGHT="F5F9FC"; GOLD="FFF2CC"
PALE_GREEN="E2F0D9"; PALE_RED="FCE4D6"; PALE_YELLOW="FFF2CC"; GREY="666666"; BLACK="000000"
LINK_GREEN="008000"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'; FMT_SCORE='0.0'
THIN=Side(style="thin",color="D9E1F2")


def _fill(color): return PatternFill("solid",fgColor=color)

def _num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        return float(v)
    except Exception:
        return default

def _clamp(x,lo=0,hi=100): return max(lo,min(hi,x))

def _find_row(ws,label,col=1,contains=False):
    if ws is None: return None
    needle=str(label).strip().lower()
    for r in range(1,ws.max_row+1):
        text=str(ws.cell(r,col).value or "").strip().lower()
        if (needle in text) if contains else (text==needle):
            return r
    return None

def _label_value(wb,sheet,label,value_col=2,contains=False):
    if sheet not in wb.sheetnames: return None
    ws=wb[sheet]; r=_find_row(ws,label,1,contains)
    return ws.cell(r,value_col).value if r else None

def _style_title(ws,text):
    for c in range(1,9):
        ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=text; ws["A1"].font=Font(bold=True,color=WHITE,size=19)
    ws.sheet_view.showGridLines=False

def _section(ws,row,title,end=8):
    for c in range(1,end+1):
        cell=ws.cell(row,c); cell.fill=_fill(NAVY); cell.font=Font(bold=True,color=WHITE,size=11)
    ws.cell(row,1,title)

def _header(ws,row,headers):
    for c,v in enumerate(headers,1):
        cell=ws.cell(row,c,v); cell.fill=_fill(BLUE); cell.font=Font(bold=True,color=WHITE)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=Border(bottom=THIN)

def _score_from_premium(premium,lower_better=True):
    if premium is None: return None
    return _clamp(50 + (-50*premium if lower_better else 35*premium))

def _data_quality(wb):
    if "Data Quality" not in wb.sheetnames: return 0,0,0
    ws=wb["Data Quality"]; passed=review=failed=0
    for r in range(1,ws.max_row+1):
        s=str(ws.cell(r,2).value or "").upper().strip()
        if s=="PASS": passed+=1
        elif s=="REVIEW": review+=1
        elif s=="FAIL": failed+=1
    return passed,review,failed

def _advanced_score(wb,label):
    if "Advanced Analytics" not in wb.sheetnames: return None
    ws=wb["Advanced Analytics"]; r=_find_row(ws,label,1)
    return _num(ws.cell(r,2).value) if r else None

def _bayes_prob(wb,scenario):
    if "Base Rates & Probabilities" not in wb.sheetnames: return None
    ws=wb["Base Rates & Probabilities"]; r=_find_row(ws,scenario,1)
    return _num(ws.cell(r,4).value) if r else None

def _comparative_gap(wb,label):
    if "Comparative Analysis" not in wb.sheetnames: return None,None,None
    ws=wb["Comparative Analysis"]; r=_find_row(ws,label,1)
    if not r: return None,None,None
    return _num(ws.cell(r,2).value),_num(ws.cell(r,3).value),_num(ws.cell(r,4).value)

def _metric_snapshot(wb):
    company=wb["Company Data"] if "Company Data" in wb.sheetnames else None
    price=_num(company["B8"].value) if company else None
    sector=company["B6"].value if company else None; industry=company["B7"].value if company else None
    base=_num(_label_value(wb,"Dashboard","Base DCF Value / Share"))
    pw=_num(_label_value(wb,"Dashboard","Probability-Weighted Value / Share"))
    rev_cagr=_num(_label_value(wb,"Dashboard","Revenue CAGR",contains=True))
    op_margin=_num(_label_value(wb,"Dashboard","Operating Margin",contains=True))
    fcf_margin=_num(_label_value(wb,"Dashboard","FCF Margin",contains=True))
    severe=_num(_label_value(wb,"Three-Case Scenarios","Severe Bear Value / Share",value_col=2))
    base_up=(base/price-1) if base is not None and price else None
    pw_up=(pw/price-1) if pw is not None and price else None
    severe_down=(severe/price-1) if severe is not None and price else None
    fpe,peer_fpe,fpe_gap=_comparative_gap(wb,"Forward P/E")
    evrev,peer_evrev,evrev_gap=_comparative_gap(wb,"EV/Revenue")
    evebitda,peer_evebitda,evebitda_gap=_comparative_gap(wb,"EV/EBITDA")
    _,peer_margin,margin_gap=_comparative_gap(wb,"Operating Margin")
    market_base=_num(_label_value(wb,"Market Expectations","10Y Revenue CAGR",2))
    market_implied=_num(_label_value(wb,"Market Expectations","10Y Revenue CAGR",3))
    moat=_num(_label_value(wb,"Moat & Competitive Advantage","Average Moat Score"))
    ai=_num(_label_value(wb,"AI Impact Analysis","AI opportunity intensity",8,contains=True))
    return {
        "price":price,"sector":sector,"industry":industry,"base":base,"pw":pw,"base_up":base_up,"pw_up":pw_up,
        "rev_cagr":rev_cagr,"op_margin":op_margin,"fcf_margin":fcf_margin,"severe":severe,"severe_down":severe_down,
        "fpe":fpe,"peer_fpe":peer_fpe,"fpe_gap":fpe_gap,"evrev":evrev,"peer_evrev":peer_evrev,"evrev_gap":evrev_gap,
        "evebitda":evebitda,"peer_evebitda":peer_evebitda,"evebitda_gap":evebitda_gap,"peer_margin":peer_margin,"margin_gap":margin_gap,
        "growth_score":_advanced_score(wb,"Growth"),"profit_score":_advanced_score(wb,"Profitability"),"fcf_score":_advanced_score(wb,"FCF Quality"),
        "balance_score":_advanced_score(wb,"Balance Sheet"),"stress_score":_advanced_score(wb,"Stress Robustness"),
        "bear_prob":_bayes_prob(wb,"Bear"),"base_prob":_bayes_prob(wb,"Base"),"bull_prob":_bayes_prob(wb,"Bull"),
        "market_base_growth":market_base,"market_implied_growth":market_implied,"moat":moat,"ai":ai,
    }

def _investment_score(m):
    pieces=[]
    def add(name,score,weight,detail):
        if score is not None and math.isfinite(score): pieces.append((name,_clamp(score),weight,detail))
    ups=[x for x in (m["base_up"],m["pw_up"]) if x is not None]
    if ups:
        avg=sum(ups)/len(ups); add("Absolute valuation",50+100*avg,30,f"Average model upside/downside {avg:.1%}")
    valuation_gaps=[x for x in (m["fpe_gap"],m["evrev_gap"],m["evebitda_gap"]) if x is not None]
    if valuation_gaps:
        rel=sum(_score_from_premium(x,True) for x in valuation_gaps)/len(valuation_gaps)
        add("Relative valuation",rel,15,"Same-sector / industry peer multiples")
    if m["growth_score"] is not None: add("Growth",m["growth_score"],10,"Advanced Analytics growth score")
    elif m["rev_cagr"] is not None: add("Growth",50+(m["rev_cagr"]-.05)/.15*50,10,f"Revenue CAGR {m['rev_cagr']:.1%}")
    if m["profit_score"] is not None: add("Profitability",m["profit_score"],10,"Advanced Analytics profitability score")
    elif m["op_margin"] is not None: add("Profitability",50+(m["op_margin"]-.08)/.20*50,10,f"Operating margin {m['op_margin']:.1%}")
    if m["fcf_score"] is not None: add("FCF quality",m["fcf_score"],10,"Advanced Analytics FCF score")
    elif m["fcf_margin"] is not None: add("FCF quality",50+(m["fcf_margin"]-.05)/.15*50,10,f"FCF margin {m['fcf_margin']:.1%}")
    if m["balance_score"] is not None: add("Balance sheet",m["balance_score"],10,"Advanced Analytics balance-sheet score")
    if m["severe_down"] is not None: add("Stress resilience",100*(1+m["severe_down"]),10,f"Severe-bear downside {m['severe_down']:.1%}")
    elif m["stress_score"] is not None: add("Stress resilience",m["stress_score"],10,"Advanced Analytics stress score")
    if m["bull_prob"] is not None and m["bear_prob"] is not None:
        add("Bayesian skew",50+(m["bull_prob"]-m["bear_prob"])*100,5,f"Bull {m['bull_prob']:.0%} vs Bear {m['bear_prob']:.0%}")
    total_w=sum(x[2] for x in pieces)
    score=sum(x[1]*x[2] for x in pieces)/total_w if total_w else None
    return score,pieces

def _verdict(score,m,failed):
    if failed:
        return "REVIEW — DATA QUALITY","Resolve failed data-quality checks before relying on the investment conclusion."
    up=m.get("base_up")
    if score is None:
        return "REVIEW — INSUFFICIENT DATA","Not enough validated model outputs are available for a conclusion."
    if up is not None and up>=.25 and score>=65:
        return "ATTRACTIVE","The model shows a meaningful margin of safety and the broader quantitative score is favorable."
    if up is not None and up>=.10 and score>=55:
        return "POTENTIALLY ATTRACTIVE","Valuation is favorable, but the model still contains material execution or quality risks."
    if (up is None or up>-0.10) and score>=45:
        return "NEUTRAL / WATCHLIST","The evidence is mixed or the margin of safety is not yet large enough."
    return "UNATTRACTIVE AT CURRENT PRICE","The modeled return/risk trade-off is not favorable enough at the current price."

def _strengths_risks(m):
    strengths=[]; risks=[]
    if m["base_up"] is not None:
        (strengths if m["base_up"]>=.15 else risks if m["base_up"]<0 else strengths).append(f"Base DCF implies {m['base_up']:.1%} upside/downside to the current price.")
    if m["rev_cagr"] is not None:
        (strengths if m["rev_cagr"]>=.08 else risks if m["rev_cagr"]<.03 else strengths).append(f"Available-history revenue CAGR is {m['rev_cagr']:.1%}.")
    if m["fcf_margin"] is not None:
        (strengths if m["fcf_margin"]>=.10 else risks if m["fcf_margin"]<.04 else strengths).append(f"Latest FCF margin is {m['fcf_margin']:.1%}.")
    if m["margin_gap"] is not None:
        (strengths if m["margin_gap"]>=0 else risks).append(f"Operating margin is {abs(m['margin_gap']):.1%} {'above' if m['margin_gap']>=0 else 'below'} the same-sector peer median on a relative basis.")
    if m["fpe_gap"] is not None:
        (strengths if m["fpe_gap"]<=-.10 else risks if m["fpe_gap"]>=.20 else strengths).append(f"Forward P/E is {abs(m['fpe_gap']):.1%} {'below' if m['fpe_gap']<0 else 'above'} the peer median.")
    if m["severe_down"] is not None and m["severe_down"]<=-.50:
        risks.append(f"Severe-bear stress implies {m['severe_down']:.1%} downside, indicating substantial tail risk.")
    if m["market_implied_growth"] is not None and m["market_base_growth"] is not None:
        gap=m["market_implied_growth"]-m["market_base_growth"]
        if gap>.05: risks.append(f"Market-implied 10Y growth exceeds the Base case by {gap:.1%}, so expectations are demanding.")
        elif gap<-.03: strengths.append(f"Market-implied growth is below the Base case by {abs(gap):.1%}, leaving room for positive surprise if the model is right.")
    if m["bull_prob"] is not None and m["bear_prob"] is not None:
        if m["bull_prob"]>m["bear_prob"]+.10: strengths.append(f"Bayesian posterior favors Bull ({m['bull_prob']:.0%}) over Bear ({m['bear_prob']:.0%}).")
        elif m["bear_prob"]>m["bull_prob"]+.10: risks.append(f"Bayesian posterior favors Bear ({m['bear_prob']:.0%}) over Bull ({m['bull_prob']:.0%}).")
    return strengths[:6],risks[:6]

def ensure_investment_summary(wb,ticker):
    name="Investment Summary"
    if name in wb.sheetnames: wb.remove(wb[name])
    ws=wb.create_sheet(name,1)
    _style_title(ws,f"{ticker} — Investment Summary & Model View")
    ws["A3"]="Rules-based synthesis of the workbook's valuation, peer, quality, stress and Bayesian outputs. The verdict changes automatically when model inputs change; it is a research aid, not personalized investment advice."
    ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)
    try: ws.merge_cells("A3:H3")
    except Exception: pass

    m=_metric_snapshot(wb); passed,review,failed=_data_quality(wb); score,pieces=_investment_score(m); verdict,why=_verdict(score,m,failed)
    fill=PALE_GREEN if verdict=="ATTRACTIVE" else PALE_YELLOW if "POTENTIALLY" in verdict or "NEUTRAL" in verdict or "REVIEW" in verdict else PALE_RED
    _section(ws,5,"Investment Conclusion")
    ws["A6"]="Model View"; ws["B6"]=verdict; ws["B6"].fill=_fill(fill); ws["B6"].font=Font(bold=True,size=14)
    ws["D6"]="Quant Score / 100"; ws["E6"]=score; ws["E6"].number_format=FMT_SCORE
    ws["A7"]="Why"; ws["B7"]=why; ws["B7"].alignment=Alignment(wrap_text=True)
    try: ws.merge_cells("B7:H7")
    except Exception: pass
    ws["A8"]="Sector / Industry"; ws["B8"]=f"{m['sector'] or 'Unknown'} / {m['industry'] or 'Unknown'}"
    ws["D8"]="Data Quality"; ws["E8"]=f"{passed} PASS / {review} REVIEW / {failed} FAIL"

    _section(ws,10,"Valuation & Risk Snapshot")
    _header(ws,11,["Metric","Company / Model","Benchmark","Signal","Interpretation","","",""])
    rows=[
        ("Current Price",m["price"],None,"Market",FMT_PRICE),
        ("Base DCF Value / Share",m["base"],m["base_up"],"Upside / downside",FMT_PRICE),
        ("Probability-Weighted Value / Share",m["pw"],m["pw_up"],"Upside / downside",FMT_PRICE),
        ("Severe Bear Value / Share",m["severe"],m["severe_down"],"Stress downside",FMT_PRICE),
        ("Forward P/E",m["fpe"],m["peer_fpe"],"Same-sector peer median","0.0x"),
        ("EV/EBITDA",m["evebitda"],m["peer_evebitda"],"Same-sector peer median","0.0x"),
        ("Revenue CAGR",m["rev_cagr"],None,"Available history",FMT_PCT),
        ("Operating Margin",m["op_margin"],m["peer_margin"],"Peer median",FMT_PCT),
        ("FCF Margin",m["fcf_margin"],None,"Latest actual",FMT_PCT),
    ]
    for r,(lab,val,bench,signal,fmt) in enumerate(rows,12):
        ws.cell(r,1,lab); ws.cell(r,2,val); ws.cell(r,2).number_format=fmt; ws.cell(r,3,bench)
        if lab in {"Base DCF Value / Share","Probability-Weighted Value / Share","Severe Bear Value / Share"}: ws.cell(r,3).number_format=FMT_PCT
        elif bench is not None: ws.cell(r,3).number_format=fmt
        ws.cell(r,4,signal)
        if lab=="Base DCF Value / Share" and m["base_up"] is not None: ws.cell(r,5,"Favorable" if m["base_up"]>=.15 else "Mixed" if m["base_up"]>=0 else "Unfavorable")
        elif lab=="Severe Bear Value / Share" and m["severe_down"] is not None: ws.cell(r,5,"High tail risk" if m["severe_down"]<=-.5 else "Moderate stress risk")
        elif lab=="Forward P/E" and m["fpe_gap"] is not None: ws.cell(r,5,"Discount" if m["fpe_gap"]<0 else "Premium")

    _section(ws,23,"Quantitative Score Bridge")
    _header(ws,24,["Dimension","Score / 100","Weight","Weighted Contribution","Evidence","","",""])
    totalw=sum(x[2] for x in pieces)
    for r,(name_i,sc,w,detail) in enumerate(pieces,25):
        ws.cell(r,1,name_i); ws.cell(r,2,sc); ws.cell(r,3,w/totalw if totalw else None); ws.cell(r,4,sc*(w/totalw) if totalw else None); ws.cell(r,5,detail)
        ws.cell(r,2).number_format=FMT_SCORE; ws.cell(r,3).number_format=FMT_PCT; ws.cell(r,4).number_format=FMT_SCORE; ws.cell(r,5).alignment=Alignment(wrap_text=True)

    strength_row=25+max(8,len(pieces))+1
    _section(ws,strength_row,"What the Numbers Like")
    strengths,risks=_strengths_risks(m)
    for i,text in enumerate(strengths or ["No major quantitative strength passed the current rule thresholds."],strength_row+1):
        ws.cell(i,1,"+"); ws.cell(i,2,text); ws.cell(i,2).alignment=Alignment(wrap_text=True)
        try: ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=8)
        except Exception: pass
    risk_row=strength_row+max(3,len(strengths)+1)+2
    _section(ws,risk_row,"Biggest Quantitative Risks / What Must Go Right")
    for i,text in enumerate(risks or ["No major quantitative risk passed the current rule thresholds; qualitative risks still require analyst review."],risk_row+1):
        ws.cell(i,1,"!"); ws.cell(i,2,text); ws.cell(i,2).alignment=Alignment(wrap_text=True)
        try: ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=8)
        except Exception: pass

    foot=risk_row+max(3,len(risks)+1)+2
    ws.cell(foot,1,"Interpretation")
    ws.cell(foot,2,"A favorable model view does not guarantee a favorable return. Re-underwrite the thesis when earnings, consensus, management guidance, capital allocation, sector conditions or the share price change.")
    ws.cell(foot,2).font=Font(italic=True,color=GREY); ws.cell(foot,2).alignment=Alignment(wrap_text=True)
    try: ws.merge_cells(start_row=foot,start_column=2,end_row=foot,end_column=8)
    except Exception: pass

    widths={"A":31,"B":24,"C":18,"D":23,"E":48,"F":3,"G":3,"H":3}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    ws.freeze_panes="A5"
    return ws
