from __future__ import annotations

from pathlib import Path
from typing import Iterable
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .common import MLResult

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREY="666666"; LIGHT="F5F9FC"; GOLD="FFF2CC"
THIN=Side(style="thin",color="D9E1F2")


def _fill(c): return PatternFill("solid",fgColor=c)
def _header(ws,row,headers):
    for c,v in enumerate(headers,1):
        x=ws.cell(row,c,v); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE); x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); x.border=Border(bottom=THIN)
def _section(ws,row,title,end=9):
    for c in range(1,end+1): ws.cell(row,c).fill=_fill(NAVY); ws.cell(row,c).font=Font(bold=True,color=WHITE,size=11)
    ws.cell(row,1,title)


def write_ml_sheet(workbook_path: Path, ticker: str, results: Iterable[MLResult]) -> Path:
    wb=load_workbook(workbook_path)
    if "ML & Quantitative Research" in wb.sheetnames: wb.remove(wb["ML & Quantitative Research"])
    ws=wb.create_sheet("ML & Quantitative Research")
    ws.sheet_view.showGridLines=False; ws.freeze_panes="A6"
    for c in range(1,10): ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=f"{ticker} — Machine Learning & Quantitative Research"; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws.merge_cells("A3:I3"); ws["A3"]="ML is a second-opinion layer. It does not overwrite DCF assumptions, consensus inputs, thesis decisions or portfolio trades. Walk-forward validation and data-readiness gates are preferred to apparent completeness."; ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)
    _section(ws,5,"Six-Model Decision Dashboard",9)
    _header(ws,6,["Model","Status","Prediction / State","Confidence","Validation / Data","Top Drivers","Use in Research","Caveat","Action"])
    row=7
    result_list=list(results)
    for res in result_list:
        metrics=res.metrics or {}; validation=metrics.get("walk_forward") or metrics.get("hgb_walk_forward") or metrics
        drivers=", ".join(str(d.get("feature")) for d in (res.drivers or [])[:4])
        use={
            "Expected 12M Excess Return":"Cross-check valuation and variant perception.",
            "Consensus / Earnings Surprise":"Challenge near-term consensus and revision risk.",
            "Financial Anomaly Detection":"Flag unusual accounting/operating patterns for diligence.",
            "Market Regime Classifier":"Contextualize portfolio and factor exposure; not a timing signal.",
            "AI Impact ML":"Test whether accumulating AI KPIs predict monetization/economic outcomes.",
            "Portfolio ML / Position Sizing":"Translate expected-return evidence into constrained risk-aware sizing.",
        }.get(res.name,"Research cross-check")
        values=[res.name,res.status,res.prediction,res.confidence,str(validation)[:240],drivers,use,"Model output is conditional on historical data and may fail out of sample.","REVIEW" if res.status!="PASS" else "USE AS SECOND OPINION"]
        for c,v in enumerate(values,1): ws.cell(row,c,v); ws.cell(row,c).alignment=Alignment(wrap_text=True,vertical="top")
        if isinstance(res.prediction,float): ws.cell(row,3).number_format='0.0%;[Red](0.0%);-'
        ws.cell(row,2).fill=_fill(LIGHT if res.status=="PASS" else GOLD); ws.cell(row,2).font=Font(bold=True)
        ws.row_dimensions[row].height=52; row+=1

    for res in result_list:
        row+=1; _section(ws,row,res.name,9); row+=1
        ws.cell(row,1,"Summary"); ws.cell(row,2,res.summary); ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=9); ws.cell(row,2).alignment=Alignment(wrap_text=True); row+=1
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

    widths={"A":32,"B":24,"C":22,"D":15,"E":42,"F":36,"G":45,"H":45,"I":24}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    wb.save(workbook_path)
    return workbook_path
