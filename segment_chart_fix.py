"""Final cross-company reliability, peer, segment-chart and AI refresh stage.

This stage deliberately runs after the first-pass workbook build. It repairs shorter
historical records, applies verified issuer fallbacks, recalibrates cash conversion,
rebuilds analytics that depended on stale history, enforces dynamic same-sector peers,
and then recreates segment/business charts from the final disclosure schema.
"""

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from ai_effect_analysis import ensure_ai_impact_analysis
from amzn_model_repair import repair_amzn_model
from cvs_model_repair import repair_cvs_model
from gev_model_repair import repair_gev_model
from cross_company_cleanup import refresh_cross_company_tabs
from dynamic_peer_engine import ensure_dynamic_peer_comps
from model_reliability import prepare_model_reliability, finalize_model_reliability
from model_quality_v3 import calibrate_scenario_cash_flow, ensure_model_quality
from advanced_analytics_v2 import ensure_advanced_analytics
from institutional_layers import ensure_institutional_layers
from sector_peer_bayes import ensure_bayesian_base_rates, repair_cross_sheet_context

GREY="666666"; PALE_GREEN="E2F0D9"; GOLD="FFF2CC"
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PCT='0.0%;[Red](0.0%);-'

def _fill(c): return PatternFill("solid",fgColor=c)
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

def _shorten_final_titles(wb):
    if "Three-Case Scenarios" in wb.sheetnames: wb["Three-Case Scenarios"]["A1"]="Three-Case Scenarios — 10-Year DCF"
    if "AI Impact Analysis" in wb.sheetnames:
        ai=wb["AI Impact Analysis"]; ai["A3"]="Institutional AI lens: reported monetization, segment exposure, capital intensity, disruption risk, and AI upside/downside versus the existing Base DCF."; ai["A23"]="AI Segment Exposure & Scoring"; ai["A33"]="AI Surprise Scenarios vs Base Case"

def _peer_quality_check(wb,ticker,peers):
    if "Data Quality" not in wb.sheetnames: return
    target_sector=str(wb["Company Data"]["B6"].value or "") if "Company Data" in wb.sheetnames else ""
    peer_ws=wb["Peer Comps"] if "Peer Comps" in wb.sheetnames else None
    sectors=[]
    if peer_ws:
        for r in range(5,10):
            if peer_ws.cell(r,2).value: sectors.append(str(peer_ws.cell(r,9).value or ""))
    ok=bool(sectors) and all(s==target_sector for s in sectors)
    status="PASS" if ok and len(sectors)>=3 else ("REVIEW" if ok else "FAIL")
    ws=wb["Data Quality"]; row=None
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or "").strip()=="Peer comps sector alignment": row=r; break
    row=row or ws.max_row+1
    ws.cell(row,1,"Peer comps sector alignment"); ws.cell(row,2,status); ws.cell(row,3,f"{len(sectors)} peer(s); target sector={target_sector}; peer sectors={sectors}"); ws.cell(row,4,"Every comparison company must match the target company's detected sector; exact industry is preferred.")
    ws.cell(row,2).fill=_fill(PALE_GREEN if status=="PASS" else GOLD); ws.cell(row,2).font=Font(bold=True)
    for c in range(1,5): ws.cell(row,c).alignment=Alignment(wrap_text=True,vertical="top")

def _apply_final_controls(wb,ticker):
    # 1) Establish the historical-layout invariant before rebuilding anything downstream.
    prepare_model_reliability(wb,ticker)

    # 2) Verified issuer disclosures fill extraction gaps only where directly supported.
    repair_amzn_model(wb,ticker); repair_cvs_model(wb,ticker); repair_gev_model(wb,ticker)
    prepare_model_reliability(wb,ticker)  # picks up issuer-fallback D&A and re-syncs scenario D&A intensity

    # 3) Recalibrate forward cash conversion after the history / D&A repair.
    try: calibrate_scenario_cash_flow(wb)
    except Exception as exc: print(f"Warning: final cash-flow calibration failed: {exc}")

    # 4) Automatic peer classification and discovery: industry first, same-sector fallback only.
    peers=[]
    try: peers=ensure_dynamic_peer_comps(wb,ticker)
    except Exception as exc: print(f"Warning: dynamic peer selection failed: {exc}")

    # 5) Rebuild layers that may have consumed stale/left-aligned history earlier in the run.
    try: ensure_advanced_analytics(wb,ticker)
    except Exception as exc: print(f"Warning: final Advanced Analytics refresh failed: {exc}")
    try: ensure_model_quality(wb,ticker)
    except Exception as exc: print(f"Warning: final Model Quality refresh failed: {exc}")
    try: ensure_institutional_layers(wb,ticker)
    except Exception as exc: print(f"Warning: final Institutional Layers refresh failed: {exc}")

    # 6) Rebuild operating maps after issuer segment fallbacks.
    try: refresh_cross_company_tabs(wb,ticker)
    except Exception as exc: print(f"Warning: final cross-company refresh failed: {exc}")

    # Institutional-layer regeneration must never restore stale template peers.
    try: peers=ensure_dynamic_peer_comps(wb,ticker)
    except Exception as exc: print(f"Warning: final dynamic peer refresh failed: {exc}")

    # 7) Proper three-way Bayesian update and final cross-sheet / history controls.
    try: ensure_bayesian_base_rates(wb,ticker)
    except Exception as exc: print(f"Warning: Bayesian base-rate refresh failed: {exc}")
    try: repair_cross_sheet_context(wb)
    except Exception as exc: print(f"Warning: cross-sheet context repair failed: {exc}")
    try: finalize_model_reliability(wb,ticker)
    except Exception as exc: print(f"Warning: final model reliability controls failed: {exc}")
    _peer_quality_check(wb,ticker,peers)

def repair_segment_charts(wb,ticker):
    _apply_final_controls(wb,ticker)

    if not {"Analysis Charts","Segment Analysis"}.issubset(wb.sheetnames):
        ensure_ai_impact_analysis(wb,ticker); _shorten_final_titles(wb); return
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
    for out_r,(src_r,name) in enumerate(business[:10],3): ws.cell(out_r,39,name); ws.cell(out_r,40,f"='Segment Analysis'!D{src_r}"); ws.cell(out_r,40).number_format=FMT_BN
    if business:
        end=2+min(10,len(business)); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title=f"{latest_year} Revenue Mix"; ch.height=8.5; ch.width=13.5; ch.legend=None; ch.add_data(Reference(ws,min_col=40,min_row=2,max_row=end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=39,min_row=3,max_row=end)); ch.x_axis.numFmt="$0"; ch.x_axis.title=f"{latest_year} revenue ($bn)"; ch.visible_cells_only=False; ch.display_blanks="gap"; _labels(ch); ws.add_chart(ch,"A53")
        ws["A70"]=f"Units: {latest_year} revenue ($bn)"; ws["D70"]="Issuer-disclosed segments / revenue groups"; ws["F70"]="Source: Segment Analysis / annual filing"; ws["A71"]="Business mix uses only disclosed categories; no standalone product revenue is invented."
    else:
        ws["A55"]="No reliable disclosed business-line revenue was extracted. Complete the yellow Segment Analysis inputs to enable this chart."; ws["A55"].font=Font(italic=True,color=GREY)

    seg_section=_find(seg,"Reported Operating Segments"); margins=[]; margin_col=None; profit_col=None; margin_header="Latest Margin"; profit_header="Segment profitability"
    if seg_section:
        header=seg_section+1
        for c in range(1,min(18,seg.max_column)+1):
            text=str(seg.cell(header,c).value or "")
            if "Margin" in text and "Δ" not in text: margin_col=c; margin_header=text
            if any(k in text for k in ("Op. Income","Operating Income","EBITDA","Segment Profit")): profit_col=c; profit_header=text
        if margin_col and profit_col:
            for r in range(header+1,min(seg.max_row,header+15)+1):
                name=seg.cell(r,1).value; profit=seg.cell(r,profit_col).value
                if name in (None,""):
                    if margins: break
                    continue
                if profit in (None,""): continue
                margins.append((r,str(name)))
    ws["AP2"]="Segment"; ws["AQ2"]=margin_header
    for out_r,(src_r,name) in enumerate(margins[:10],3): ws.cell(out_r,42,name); ws.cell(out_r,43,f"='Segment Analysis'!{get_column_letter(margin_col)}{src_r}"); ws.cell(out_r,43).number_format=FMT_PCT
    if margins:
        end=2+min(10,len(margins)); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title=f"{margin_header} by Segment"; ch.height=8.5; ch.width=13.5; ch.legend=None; ch.add_data(Reference(ws,min_col=43,min_row=2,max_row=end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=42,min_row=3,max_row=end)); ch.x_axis.numFmt="0%"; ch.x_axis.title=margin_header; ch.visible_cells_only=False; ch.display_blanks="gap"; _labels(ch); ws.add_chart(ch,"I53")
        ws["I70"]=f"Margin = {profit_header} ÷ segment revenue"; ws["M70"]="Source: Segment Analysis / annual filing"; ws["I71"]="Profitability uses the issuer's disclosed segment performance measure; missing economics are not estimated."
    else:
        ws["I55"]="Segment profitability is not disclosed or not reliably extracted. Chart remains blank rather than estimating it."; ws["I55"].font=Font(italic=True,color=GREY)
    ws["A112"]=f"External segment source: {ticker} annual filing / Segment Analysis sheet. Monte Carlo, scenario, stress and DCF values are model outputs based on workbook assumptions."; ws["A112"].font=Font(italic=True,color=GREY,size=9)

    ensure_ai_impact_analysis(wb,ticker)
    finalize_model_reliability(wb,ticker)
    _shorten_final_titles(wb)
