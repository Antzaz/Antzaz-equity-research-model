from __future__ import annotations

from pathlib import Path
from typing import Iterable
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

from .common import MLResult

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREY="666666"; LIGHT="F5F9FC"; GOLD="FFF2CC"; RED="FCE4D6"; GREEN="E2F0D9"
THIN=Side(style="thin",color="D9E1F2")


def _fill(c): return PatternFill("solid",fgColor=c)
def _header(ws,row,headers):
    for c,v in enumerate(headers,1):
        x=ws.cell(row,c,v); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE); x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); x.border=Border(bottom=THIN)
def _section(ws,row,title,end=9):
    for c in range(1,end+1): ws.cell(row,c).fill=_fill(NAVY); ws.cell(row,c).font=Font(bold=True,color=WHITE,size=11)
    ws.cell(row,1,title)
def _wf(res):
    m=res.metrics or {}
    return m.get("walk_forward") or m.get("hgb_walk_forward") or {}
def _pct(v):
    try: return float(v)
    except Exception: return None


def _validation_summary(res: MLResult) -> str:
    """Compact validation copy for the dashboard; full diagnostics remain in the detail sections."""
    wf=_wf(res)
    parts=[]
    n=wf.get("n")
    da=_pct(wf.get("directional_accuracy"))
    r2=_pct(wf.get("r2"))
    mae=_pct(wf.get("mae"))
    if n is not None: parts.append(f"N={int(n):,}")
    if da is not None: parts.append(f"DA {da:.1%}")
    if r2 is not None: parts.append(f"R² {r2:.2f}")
    if mae is not None: parts.append(f"MAE {mae:.1%}")
    if parts: return " | ".join(parts)

    metrics=res.metrics or {}
    if res.name=="Financial Anomaly Detection":
        years=metrics.get("years") or []
        return f"{len(years)} annual observations" if years else "No robust annual validation sample"
    if res.name=="Market Regime Classifier":
        probs=metrics.get("regime_probabilities") or {}
        top=max([_pct(v) or 0 for v in probs.values()] or [0])
        rows=int(metrics.get("monthly_rows") or 0)
        return f"N={rows:,} months | top regime weight {top:.1%}" if rows else f"Top regime weight {top:.1%}"
    return "See detail section"


def _add_ml_charts(ws, result_list):
    """Add comparable, decision-useful ML charts with explicit percentage scales."""
    start=7
    ws.cell(start,11,"Model"); ws.cell(start,12,"Directional Accuracy"); ws.cell(start,13,"Walk-forward R²"); ws.cell(start,14,"Validation N")
    for c in range(11,15): ws.cell(start,c).fill=_fill(BLUE); ws.cell(start,c).font=Font(bold=True,color=WHITE)
    rr=start+1
    for res in result_list:
        wf=_wf(res); da=_pct(wf.get("directional_accuracy")); r2=_pct(wf.get("r2")); n=wf.get("n")
        if da is None and r2 is None: continue
        ws.cell(rr,11,res.name); ws.cell(rr,12,da); ws.cell(rr,13,r2); ws.cell(rr,14,n)
        ws.cell(rr,12).number_format='0.0%'; ws.cell(rr,13).number_format='0.00'; rr+=1
    if rr>start+1:
        ch=BarChart(); ch.type="bar"; ch.style=10
        ch.title="Walk-forward directional accuracy (50% = no edge)"
        ch.y_axis.title="Model"; ch.x_axis.title="Directional accuracy"
        ch.x_axis.scaling.min=0; ch.x_axis.scaling.max=1; ch.x_axis.numFmt="0%"
        ch.height=6.5; ch.width=12.5
        data=Reference(ws,min_col=12,min_row=start,max_row=rr-1); cats=Reference(ws,min_col=11,min_row=start+1,max_row=rr-1)
        ch.add_data(data,titles_from_data=True); ch.set_categories(cats); ch.legend=None
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"K14")

    expected=next((r for r in result_list if r.name=="Expected 12M Excess Return" and r.drivers),None)
    if expected:
        s=29; ws.cell(s,11,"Feature"); ws.cell(s,12,"Relative Importance")
        for c in (11,12): ws.cell(s,c).fill=_fill(BLUE); ws.cell(s,c).font=Font(bold=True,color=WHITE)
        drivers=expected.drivers[:6]
        raw=[abs(_pct(d.get("importance")) or 0) for d in drivers]
        total=sum(raw)
        shares=[v/total if total>0 else 0 for v in raw]
        for i,(d,share) in enumerate(zip(drivers,shares),s+1):
            ws.cell(i,11,d.get("feature")); ws.cell(i,12,share); ws.cell(i,12).number_format="0.0%"
        ch=BarChart(); ch.type="bar"; ch.style=11
        ch.title="Expected-return drivers — relative importance"
        ch.y_axis.title="Feature"; ch.x_axis.title="Share of top-driver importance"; ch.x_axis.numFmt="0%"
        ch.height=6.5; ch.width=12.5
        ch.add_data(Reference(ws,min_col=12,min_row=s,max_row=s+len(drivers)),titles_from_data=True)
        ch.set_categories(Reference(ws,min_col=11,min_row=s+1,max_row=s+len(drivers))); ch.legend=None
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"K36")

    regime=next((r for r in result_list if r.name=="Market Regime Classifier"),None)
    probs=(regime.metrics or {}).get("regime_probabilities") if regime else None
    if probs:
        s=51; ws.cell(s,11,"Regime"); ws.cell(s,12,"Weight")
        for c in (11,12): ws.cell(s,c).fill=_fill(BLUE); ws.cell(s,c).font=Font(bold=True,color=WHITE)
        items=sorted(probs.items(),key=lambda x:x[1],reverse=True)
        for i,(name,val) in enumerate(items,s+1): ws.cell(i,11,name); ws.cell(i,12,val); ws.cell(i,12).number_format='0.0%'
        ch=BarChart(); ch.type="bar"; ch.style=12
        ch.title="Current market-regime weights"
        ch.y_axis.title="Regime"; ch.x_axis.title="Distance-based weight"
        ch.x_axis.scaling.min=0; ch.x_axis.scaling.max=1; ch.x_axis.numFmt="0%"
        ch.height=6.5; ch.width=12.5
        ch.add_data(Reference(ws,min_col=12,min_row=s,max_row=s+len(items)),titles_from_data=True)
        ch.set_categories(Reference(ws,min_col=11,min_row=s+1,max_row=s+len(items))); ch.legend=None
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"K58")


def write_ml_sheet(workbook_path: Path, ticker: str, results: Iterable[MLResult]) -> Path:
    wb=load_workbook(workbook_path)
    if "ML & Quantitative Research" in wb.sheetnames: wb.remove(wb["ML & Quantitative Research"])
    ws=wb.create_sheet("ML & Quantitative Research")
    ws.sheet_view.showGridLines=False; ws.freeze_panes="A6"
    for c in range(1,10): ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=f"{ticker} — Machine Learning & Quantitative Research"; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws.merge_cells("A3:I3"); ws["A3"]="ML is a second-opinion layer. Computational success is not enough: validation quality gates downgrade weak out-of-sample evidence. ML never overwrites DCF assumptions, reported financials or thesis decisions."; ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)
    _section(ws,5,"Six-Model Decision Dashboard",9)
    _header(ws,6,["Model","Evidence Status","Prediction / State","Confidence","Validation / Data","Top Drivers","Use in Research","Caveat","Action"])
    row=7
    result_list=list(results)
    for res in result_list:
        drivers=", ".join(str(d.get("feature")) for d in (res.drivers or [])[:4])
        use={
            "Expected 12M Excess Return":"Cross-check valuation only when walk-forward skill is credible.",
            "Consensus / Earnings Surprise":"Challenge near-term consensus only after enough realized earnings observations.",
            "Financial Anomaly Detection":"Flag unusual accounting/operating patterns for diligence; never proof of wrongdoing.",
            "Market Regime Classifier":"Contextualize portfolio/factor exposure; not a timing signal.",
            "AI Impact ML":"Test whether accumulating AI KPIs predict monetization/economic outcomes.",
            "Portfolio ML / Position Sizing":"Translate validated expected-return evidence into constrained risk-aware sizing.",
        }.get(res.name,"Research cross-check")
        usable=res.status=="PASS"
        values=[
            res.name,res.status,res.prediction,res.confidence,_validation_summary(res),drivers,use,
            "Historical relationships can fail out of sample. Evidence-status gate determines whether output is usable.",
            "USE AS SECOND OPINION" if usable else "DO NOT USE IN SCORE / REVIEW",
        ]
        for c,v in enumerate(values,1): ws.cell(row,c,v); ws.cell(row,c).alignment=Alignment(wrap_text=True,vertical="top")
        if isinstance(res.prediction,float): ws.cell(row,3).number_format='0.0%;[Red](0.0%);-'
        ws.cell(row,2).fill=_fill(GREEN if res.status=="PASS" else GOLD if res.status in {"REVIEW","WEAK_SIGNAL"} else RED); ws.cell(row,2).font=Font(bold=True)
        ws.row_dimensions[row].height=58; row+=1

    for res in result_list:
        row+=1; _section(ws,row,res.name,9); row+=1
        ws.cell(row,1,"Summary"); ws.cell(row,2,res.summary); ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=9); ws.cell(row,2).alignment=Alignment(wrap_text=True); row+=1
        notes=(res.details or {}).get("quality_gate_notes") or []
        if notes:
            ws.cell(row,1,"Evidence gate"); ws.cell(row,2," ".join(notes)); ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=9); ws.cell(row,2).alignment=Alignment(wrap_text=True); row+=1
        if res.drivers:
            _header(ws,row,["Driver","Importance / Coefficient","Notes"]); row+=1
            for d in res.drivers[:8]:
                key=d.get("feature"); val=d.get("importance",d.get("coefficient",d.get("standardized_deviation")))
                ws.cell(row,1,key); ws.cell(row,2,val); ws.cell(row,3,"Model-relative explanatory signal; not causal proof."); row+=1
        if res.name=="Portfolio ML / Position Sizing" and (res.details or {}).get("weights"):
            _header(ws,row,["Ticker","Suggested Weight","Current Weight","Change","Expected Return Input"]); row+=1
            for rec in res.details["weights"]:
                for c,key in enumerate(["ticker","suggested_weight","current_weight","weight_change","expected_return_input"],1): ws.cell(row,c,rec.get(key))
                for c in range(2,6): ws.cell(row,c).number_format='0.0%;[Red](0.0%);-'
                row+=1

    _add_ml_charts(ws,result_list)
    widths={"A":32,"B":24,"C":22,"D":15,"E":28,"F":32,"G":42,"H":42,"I":26,"J":3,"K":34,"L":20,"M":18,"N":14}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    wb.save(workbook_path)
    return workbook_path
