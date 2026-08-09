"""Verified GE Vernova fallback repairs.

Runs only for GEV. It restores directly disclosed FY2023-FY2025 segment revenue and
Segment EBITDA when generic SEC HTML-table extraction fails, adds the consolidated
Equipment/Services revenue mix, and supplies D&A history when Company Facts tags are
incomplete. No undisclosed product economics are estimated.
"""

from openpyxl.styles import Font, Alignment
from openpyxl.comments import Comment

SEC_SEG="https://www.sec.gov/Archives/edgar/data/1996810/000199681026000015/R132.htm"
SEC_10K="https://www.sec.gov/Archives/edgar/data/1996810/000199681026000015/gev-20251231.htm"
BLUE="0000FF"; GREEN="008000"; BLACK="000000"; GREY="666666"
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PCT='0.0%;[Red](0.0%);-'

SEGMENTS=[
    ("Power",[17.436,18.127,19.767],[1.722,2.268,2.902]),
    ("Wind",[9.826,9.701,9.110],[-1.033,-0.588,-0.598]),
    ("Electrification",[6.378,7.550,9.642],[0.234,0.679,1.433]),
]
REVENUE_GROUPS=[
    ("Equipment",[18.258,18.952,20.934]),
    ("Services",[14.981,15.983,17.134]),
]
D_AND_A={2022:1.797,2023:.964,2024:1.172,2025:.853}

def _input(cell,value,source=SEC_SEG,fmt=None):
    cell.value=value; cell.font=Font(color=BLUE); cell.comment=Comment(f"Issuer-reported / filing-derived input. Source: {source}","Model repair")
    if fmt: cell.number_format=fmt

def _formula(cell,formula,fmt=None):
    cell.value=formula; cell.font=Font(color=BLACK)
    if fmt: cell.number_format=fmt

def _segment_has_data(ws):
    for r in range(7,min(ws.max_row,16)+1):
        if str(ws.cell(r,1).value or "") in {"Power","Wind","Electrification"} and isinstance(ws.cell(r,4).value,(int,float)): return True
    return False

def _repair_history_d_and_a(wb):
    if "Historical Financials" not in wb.sheetnames: return
    h=wb["Historical Financials"]
    for c in range(2,8):
        y=h.cell(3,c).value
        if isinstance(y,(int,float)) and int(y) in D_AND_A and not isinstance(h.cell(18,c).value,(int,float)):
            _input(h.cell(18,c),D_AND_A[int(y)],SEC_10K,FMT_BN)
    if "Financial Statements" in wb.sheetnames:
        ws=wb["Financial Statements"]
        # Locate Cash Flow Statement D&A row and match visible year columns dynamically.
        da_row=None; header_row=None
        for r in range(1,ws.max_row+1):
            if str(ws.cell(r,1).value or "").strip()=="Depreciation & Amortization": da_row=r
            if str(ws.cell(r,1).value or "").strip()=="Metric" and r>40: header_row=r
        if da_row and header_row:
            for c in range(2,8):
                y=ws.cell(header_row,c).value
                if isinstance(y,(int,float)) and int(y) in D_AND_A and not isinstance(ws.cell(da_row,c).value,(int,float)):
                    _input(ws.cell(da_row,c),D_AND_A[int(y)],SEC_10K,FMT_BN)

def repair_gev_model(wb,ticker):
    if str(ticker).upper()!="GEV": return False
    _repair_history_d_and_a(wb)
    if "Segment Analysis" not in wb.sheetnames: return False
    ws=wb["Segment Analysis"]
    if not _segment_has_data(ws):
        ws["A3"]="Issuer-disclosed segment schema. Status: AUTO/FALLBACK — verified GE Vernova FY2025 10-K. GE Vernova principally evaluates segments using Segment EBITDA, so profitability uses Segment EBITDA rather than GAAP operating income."
        ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)
        headers=["Segment","2023 Revenue","2024 Revenue","2025 Revenue","2025 Growth","2023–2025 CAGR","2023 Segment EBITDA","2024 Segment EBITDA","2025 Segment EBITDA","2023 EBITDA Margin","2024 EBITDA Margin","2025 EBITDA Margin","Margin Δ","2025 Revenue Mix"]
        for c,v in enumerate(headers,1): ws.cell(6,c,v)
        for r,(name,rev,ebitda) in enumerate(SEGMENTS,7):
            ws.cell(r,1,name)
            for c,v in enumerate(rev,2): _input(ws.cell(r,c),v,SEC_SEG,FMT_BN)
            for c,v in enumerate(ebitda,7): _input(ws.cell(r,c),v,SEC_SEG,FMT_BN)
            _formula(ws.cell(r,5),f'=IFERROR(D{r}/C{r}-1,"")',FMT_PCT); _formula(ws.cell(r,6),f'=IFERROR((D{r}/B{r})^(1/2)-1,"")',FMT_PCT)
            _formula(ws.cell(r,10),f'=IFERROR(G{r}/B{r},"")',FMT_PCT); _formula(ws.cell(r,11),f'=IFERROR(H{r}/C{r},"")',FMT_PCT); _formula(ws.cell(r,12),f'=IFERROR(I{r}/D{r},"")',FMT_PCT); _formula(ws.cell(r,13),f'=IFERROR(L{r}-K{r},"")',FMT_PCT); _formula(ws.cell(r,14),f'=IFERROR(D{r}/SUM($D$7:$D$9),"")',FMT_PCT)
        for r in range(10,17):
            for c in range(1,15): ws.cell(r,c).value=None
        # Locate business-line section in the standardized sheet.
        start=None
        for r in range(1,ws.max_row+1):
            if str(ws.cell(r,1).value or "").strip()=="Revenue by Business Line": start=r; break
        if start:
            for i,(name,vals) in enumerate(REVENUE_GROUPS,start+2):
                ws.cell(i,1,name)
                for c,v in enumerate(vals,2): _input(ws.cell(i,c),v,SEC_SEG,FMT_BN)
                _formula(ws.cell(i,5),f'=IFERROR(D{i}/C{i}-1,"")',FMT_PCT); _formula(ws.cell(i,6),f'=IFERROR((D{i}/B{i})^(1/2)-1,"")',FMT_PCT); _formula(ws.cell(i,7),f'=IFERROR(D{i}/SUM($D${start+2}:$D${start+3}),"")',FMT_PCT); ws.cell(i,8,SEC_SEG); ws.cell(i,8).font=Font(color=GREEN)
            for r in range(start+4,min(ws.max_row,start+12)+1):
                if str(ws.cell(r,1).value or "").strip() not in {"Source & Data Quality","SEC 10-K Source","Extraction Status","Important"}:
                    for c in range(1,9): ws.cell(r,c).value=None
        # Source block normally occupies final rows in standardized schema.
        for r in range(1,ws.max_row+1):
            label=str(ws.cell(r,1).value or "").strip()
            if label=="SEC 10-K Source": ws.cell(r,2,SEC_SEG); ws.cell(r,2).font=Font(color=GREEN)
            elif label=="Extraction Status": ws.cell(r,2,"AUTO/FALLBACK — verified FY2025 10-K: 3 operating segments; Equipment/Services revenue mix")
            elif label=="Important": ws.cell(r,2,"Segment revenues include intersegment activity; Segment EBITDA is GE Vernova's segment performance measure. Equipment/Services are consolidated revenue categories."); ws.cell(r,2).alignment=Alignment(wrap_text=True)
    return True
