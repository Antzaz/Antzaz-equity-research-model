"""Repair company-agnostic Business Mix and Segment Profitability charts.

The original chart layer assumed Alphabet's disclosure layout. This overlay runs after
Analysis Charts is built, discovers the actual Segment Analysis schema, and recreates
both charts with dynamic year labels and generic interpretation.
"""

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

GREY="666666"; FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PCT='0.0%;[Red](0.0%);-'

def _find(ws,label):
    target=label.strip().lower()
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or "").strip().lower()==target: return r
    return None
def _chart_row(ch):
    try: return ch.anchor._from.row
    except Exception: return None
def _remove_target_charts(ws):
    ws._charts=[ch for ch in ws._charts if _chart_row(ch)!=52]
def _labels(ch):
    try:
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True
    except Exception: pass

def repair_segment_charts(wb,ticker):
    if not {"Analysis Charts","Segment Analysis"}.issubset(wb.sheetnames): return
    ws=wb["Analysis Charts"]; seg=wb["Segment Analysis"]; _remove_target_charts(ws)
    for row in ws.iter_rows(min_row=2,max_row=20,min_col=39,max_col=45):
        for cell in row: cell.value=None

    section=_find(seg,"Revenue by Business Line")
    if section is None: section=_find(seg,"Revenue by Business Line / Disclosed Revenue Group")
    business=[]; latest_year="Latest"
    if section:
        header=section+1; latest_col=4; h=str(seg.cell(header,latest_col).value or ""); latest_year=h if h else "Latest"
        for r in range(header+1,min(seg.max_row,header+20)+1):
            name=seg.cell(r,1).value; val=seg.cell(r,latest_col).value
            if name in (None,""):
                if business: break
                continue
            if val in (None,""): continue
            business.append((r,str(name)))
    ws["AM2"]="Business Line / Revenue Group"; ws["AN2"]=f"{latest_year} Revenue ($bn)"
    for out_r,(src_r,name) in enumerate(business[:10],3):
        ws.cell(out_r,39,name); ws.cell(out_r,40,f"='Segment Analysis'!D{src_r}"); ws.cell(out_r,40).number_format=FMT_BN
    if business:
        end=2+min(10,len(business)); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title=f"{latest_year} Revenue Mix"; ch.height=8.5; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=40,min_row=2,max_row=end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=39,min_row=3,max_row=end)); ch.x_axis.numFmt="$0"; ch.x_axis.title=f"{latest_year} revenue ($bn)"
        ch.visible_cells_only=False; ch.display_blanks="gap"; _labels(ch); ws.add_chart(ch,"A53")
        ws["A70"]=f"Units: {latest_year} revenue ($bn)"; ws["D70"]="Issuer-disclosed segments / revenue groups"; ws["F70"]="Source: Segment Analysis / 10-K"; ws["A71"]="Business mix uses only disclosed categories; no standalone product revenue is invented."
    else:
        ws["A55"]="No reliable disclosed business-line revenue was extracted. Complete the yellow Segment Analysis inputs to enable this chart."; ws["A55"].font=Font(italic=True,color=GREY)

    seg_section=_find(seg,"Reported Operating Segments"); margins=[]; margin_col=None; margin_header="Latest Margin"
    if seg_section:
        header=seg_section+1
        for c in range(1,min(18,seg.max_column)+1):
            text=str(seg.cell(header,c).value or "")
            if "Margin" in text and "Δ" not in text: margin_col=c; margin_header=text
        if margin_col:
            for r in range(header+1,min(seg.max_row,header+15)+1):
                name=seg.cell(r,1).value; val=seg.cell(r,margin_col).value
                if name in (None,""):
                    if margins: break
                    continue
                if val in (None,""): continue
                margins.append((r,str(name)))
    ws["AP2"]="Segment"; ws["AQ2"]=margin_header
    for out_r,(src_r,name) in enumerate(margins[:10],3):
        ws.cell(out_r,42,name); ws.cell(out_r,43,f"='Segment Analysis'!{get_column_letter(margin_col)}{src_r}"); ws.cell(out_r,43).number_format=FMT_PCT
    if margins:
        end=2+min(10,len(margins)); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title=f"{margin_header} by Segment"; ch.height=8.5; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=43,min_row=2,max_row=end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=42,min_row=3,max_row=end)); ch.x_axis.numFmt="0%"; ch.x_axis.title="Operating margin"
        ch.visible_cells_only=False; ch.display_blanks="gap"; _labels(ch); ws.add_chart(ch,"I53")
        ws["I70"]="Margin = segment operating income ÷ segment revenue"; ws["M70"]="Source: Segment Analysis / 10-K"; ws["I71"]="Profitability is shown only when the issuer discloses segment operating income; missing margins are not estimated."
    else:
        ws["I55"]="Segment operating income is not disclosed or not reliably extracted. Profitability chart intentionally remains blank rather than estimating it."; ws["I55"].font=Font(italic=True,color=GREY)
    ws["A112"]=f"External segment source: {ticker} annual filing / Segment Analysis sheet. Monte Carlo, scenario, stress and DCF values are model outputs based on workbook assumptions."; ws["A112"].font=Font(italic=True,color=GREY,size=9)
