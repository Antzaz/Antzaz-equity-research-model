from __future__ import annotations

"""Workbook-wide numeric coverage audit.

Classifies important blank cells so extraction failures are distinguishable from legitimate
non-disclosure, immature history, non-applicability and analyst-input fields.
"""

from openpyxl.styles import Alignment, Font, PatternFill

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GOLD="FFF2CC"
IMPORTANT_SHEETS=(
    "Company Data","Historical Financials","Financial Statements","Expectations & Consensus",
    "Segment Analysis","Supply Chain","Capital Allocation","Valuation History",
    "Advanced Analytics","Peer Comps","Cost of Capital","SOTP Framework",
)
ANALYST_INPUT_HINTS=("analyst","input","assumption","multiple","manual")
WAIT_HINTS=("forecast accountability","earnings & revisions","thesis timeline")
NON_NUMERIC_HINTS=("source","notes","status","company","metric","segment","supplier","relationship","evidence","year","date","period")

def _fill(c): return PatternFill("solid",fgColor=c)

def _header_text(ws,col):
    for r in range(1,min(ws.max_row,12)+1):
        v=ws.cell(r,col).value
        if isinstance(v,str) and v.strip(): return v.strip()
    return ""

def _classify(sheet,label,header):
    text=f"{sheet} {label} {header}".lower()
    if any(x in text for x in ANALYST_INPUT_HINTS): return "ANALYST INPUT REQUIRED"
    if any(x in sheet.lower() for x in WAIT_HINTS): return "WAITING FOR HISTORY"
    if sheet=="Supply Chain": return "NOT PUBLICLY DISCLOSED / EXTRACTION REVIEW"
    if sheet=="Segment Analysis": return "NOT PUBLICLY DISCLOSED / EXTRACTION REVIEW"
    if sheet=="SOTP Framework": return "ANALYST INPUT REQUIRED"
    return "AVAILABLE DATA / EXTRACTION REVIEW"

def ensure_numeric_coverage_audit(wb,ticker):
    if "Numeric Coverage Audit" in wb.sheetnames: wb.remove(wb["Numeric Coverage Audit"])
    ws=wb.create_sheet("Numeric Coverage Audit"); ws.sheet_view.showGridLines=False
    for c in range(1,8): ws.cell(1,c).fill=_fill(NAVY)
    ws["A1"]=f"{ticker} — Numeric Coverage Audit"; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws["A3"]="Blank numeric fields are classified rather than silently accepted. REVIEW means the pipeline should check available evidence; it does not mean a value should be estimated."
    ws.merge_cells("A3:G3"); ws["A3"].alignment=Alignment(wrap_text=True)
    heads=["Sheet","Cell","Row / Metric","Column / Field","Classification","Action","Value State"]
    for c,h in enumerate(heads,1): ws.cell(5,c,h); ws.cell(5,c).fill=_fill(BLUE); ws.cell(5,c).font=Font(bold=True,color=WHITE)
    out=[]; r=6
    for sheet in IMPORTANT_SHEETS:
        if sheet not in wb.sheetnames: continue
        src=wb[sheet]
        for rr in range(1,src.max_row+1):
            label=str(src.cell(rr,1).value or "").strip()
            if not label: continue
            for cc in range(2,src.max_column+1):
                cell=src.cell(rr,cc)
                if cell.value not in (None,""): continue
                header=_header_text(src,cc)
                descriptor=f"{label} {header}".lower()
                if any(x in descriptor for x in NON_NUMERIC_HINTS): continue
                # Only audit blanks that are styled/formatted like model numeric fields.
                fmt=str(cell.number_format or "").lower()
                numeric_style=any(x in fmt for x in ("0","%","#","$","€")) or cell.fill.fill_type=="solid"
                if not numeric_style: continue
                cls=_classify(sheet,label,header)
                action=("Research / extraction fallback required" if "REVIEW" in cls else
                        "Populate only with explicit analyst assumption" if "ANALYST" in cls else
                        "Allow history to mature")
                row=[sheet,cell.coordinate,label,header,cls,action,"BLANK"]
                for c,v in enumerate(row,1): ws.cell(r,c,v)
                if "REVIEW" in cls: ws.cell(r,5).fill=_fill(GOLD)
                out.append({"sheet":sheet,"cell":cell.coordinate,"metric":label,"field":header,"classification":cls})
                r+=1
    if not out:
        ws["A6"]="PASS"; ws["B6"]="No material styled numeric blanks detected in audited sheets."
    widths={"A":28,"B":12,"C":38,"D":30,"E":38,"F":40,"G":14}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    ws.freeze_panes="A6"
    return {"blank_numeric_fields":len(out),"review_fields":sum("REVIEW" in x["classification"] for x in out),"items":out}
