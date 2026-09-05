"""Final chart-readability pass for the deterministic research workbook.

The workbook is assembled by several layers.  A late idempotent pass is therefore safer than
assuming every earlier chart writer ran exactly once.  This module removes overlapping charts at
known analysis anchors, rebuilds the three text-heavy charts with short labels, and reduces label
clutter on Monte Carlo / stress outputs without removing underlying research detail from tables.
"""

from __future__ import annotations

import re

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

GREY="666666"
FMT_BN='#,##0.0;[Red](#,##0.0);-'
FMT_PCT='0.0%;[Red](0.0%);-'


def _num(v):
    try:
        if isinstance(v,bool) or v in (None,""): return None
        return float(v)
    except Exception:
        return None


def _anchor(chart):
    try:
        return int(chart.anchor._from.row),int(chart.anchor._from.col)
    except Exception:
        return None,None


def _chart_title(chart):
    """Best-effort plain title extraction is intentionally avoided; anchors are deterministic."""
    return getattr(chart,"title",None)


def _remove_at(ws,anchors):
    anchors=set(anchors)
    ws._charts=[ch for ch in getattr(ws,"_charts",[]) if _anchor(ch) not in anchors]


def _set_text_categories(chart,ws,col,first,last,titles):
    sheet_name=ws.title.replace("'","''")
    letter=get_column_letter(col)
    formula=f"'{sheet_name}'!${letter}${first}:${letter}${last}"
    for i,series in enumerate(chart.series):
        series.cat=AxDataSource(strRef=StrRef(f=formula))
        if i<len(titles): series.tx=SeriesLabel(v=str(titles[i]))


def _short_label(label,max_len=27):
    text=str(label or "").strip()
    aliases={
        "Google Search & other":"Search",
        "YouTube ads":"YouTube",
        "Google Network":"Network",
        "Google subscriptions, platforms, and devices":"Subs / Platforms / Devices",
        "Google Cloud":"Cloud",
        "Google Services":"Services",
        "Other Bets":"Other Bets",
    }
    if text in aliases: return aliases[text]
    # Parenthetical detail belongs in the table/notes, not on chart axes.
    text=re.sub(r"\s*\([^)]{8,}\)\s*$","",text).strip()
    replacements={
        "Revenue Growth":"Growth",
        "Operating Margin":"Margin",
        "EBIT Margin":"Margin",
        "Terminal Growth":"TGR",
        "Combined Severe Bear":"Severe Bear",
        "Probability Weighted":"Prob.-Weighted",
    }
    for old,new in replacements.items(): text=text.replace(old,new)
    text=re.sub(r"\s+"," ",text)
    return text if len(text)<=max_len else text[:max_len-1].rstrip()+"…"


def _find_section(ws,*labels):
    wanted={str(x).strip().lower() for x in labels}
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or "").strip().lower() in wanted:
            return r
    return None


def _business_rows(seg):
    section=_find_section(seg,"Revenue by Business Line","Revenue by Business Line / Product Group","Revenue by Business Line / Disclosed Revenue Group")
    if not section: return []
    header=section+1; latest_col=4; out=[]
    for r in range(header+1,min(seg.max_row,header+20)+1):
        name=seg.cell(r,1).value; value=_num(seg.cell(r,latest_col).value)
        if name in (None,""):
            if out: break
            continue
        if value is None: continue
        out.append((r,_short_label(name)))
    return out[:8]


def _margin_rows(seg):
    section=_find_section(seg,"Reported Operating / Reportable Segments","Reported Operating Segments","Reported Segments")
    if not section: return [],None,None
    header=section+1; margin_col=None; profit_label="Segment profit"
    for c in range(1,min(18,seg.max_column)+1):
        txt=str(seg.cell(header,c).value or "")
        if "Margin" in txt and "Δ" not in txt: margin_col=c
        if any(k in txt for k in ("Op. Income","Operating Income","EBITDA","Segment Profit","Adjusted Earnings")): profit_label=txt or profit_label
    if margin_col is None: return [],None,profit_label
    rows=[]; excluded=[]
    for r in range(header+1,min(seg.max_row,header+15)+1):
        name=seg.cell(r,1).value; margin=_num(seg.cell(r,margin_col).value)
        if name in (None,""):
            if rows or excluded: break
            continue
        if margin is None: continue
        # A -489% segment margin can be economically real, but plotting it beside 20–40% margins
        # destroys readability. Keep the fact in the note and omit it only from the common scale.
        if abs(margin)>1.0:
            excluded.append((_short_label(name),margin))
        else:
            rows.append((r,_short_label(name),margin))
    return rows[:8],excluded,profit_label


def _polish_existing_chart_titles(ws):
    for ch in getattr(ws,"_charts",[]):
        row,col=_anchor(ch)
        if (row,col)==(6,0):
            try:
                ch.title="Monte Carlo Value Distribution"
                ch.x_axis.tickLblSkip=2
            except Exception: pass
        elif (row,col)==(28,0):
            try:
                ch.title="DCF Stress Tests"
                ch.height=9.5
            except Exception: pass


def _shorten_stress_helpers(ws):
    # analysis_charts.py stores stress categories in AD3:AD13.
    for r in range(3,14):
        v=ws.cell(r,30).value
        if isinstance(v,str) and not v.startswith("="):
            ws.cell(r,30,_short_label(v,22))


def _rebuild_history_chart(wb,ws):
    if "Historical Financials" not in wb.sheetnames: return
    h=wb["Historical Financials"]
    ws["AX2"]="Year"; ws["AY2"]="Revenue"; ws["AZ2"]="FCF"
    for out_r,src_col in enumerate(range(2,8),3):
        letter=get_column_letter(src_col)
        ws.cell(out_r,50,f"='Historical Financials'!{letter}3")
        ws.cell(out_r,51,f"='Historical Financials'!{letter}4")
        ws.cell(out_r,52,f"='Historical Financials'!{letter}16")
        ws.cell(out_r,51).number_format=FMT_BN; ws.cell(out_r,52).number_format=FMT_BN
    ch=LineChart(); ch.style=10; ch.title="Revenue & Free Cash Flow"; ch.height=9; ch.width=13.5
    ch.y_axis.title="$bn"; ch.x_axis.title="Fiscal year"; ch.legend.position="b"
    ch.add_data(Reference(ws,min_col=51,max_col=52,min_row=2,max_row=8),titles_from_data=True)
    ch.set_categories(Reference(ws,min_col=50,min_row=3,max_row=8)); ch.visible_cells_only=False; ch.display_blanks="gap"
    ws.add_chart(ch,"I29")
    ws["I48"]="Revenue = reported sales | FCF = OCF − Capex | Source: Historical Financials"
    ws["I48"].font=Font(size=9,color=GREY)


def _rebuild_business_chart(wb,ws):
    if "Segment Analysis" not in wb.sheetnames: return
    seg=wb["Segment Analysis"]; rows=_business_rows(seg)
    for r in range(2,14):
        ws.cell(r,39).value=None; ws.cell(r,40).value=None
    ws["AM2"]="Business"; ws["AN2"]="Revenue ($bn)"
    for out_r,(src_r,name) in enumerate(rows,3):
        ws.cell(out_r,39,name); ws.cell(out_r,40,f"='Segment Analysis'!D{src_r}"); ws.cell(out_r,40).number_format=FMT_BN
    if not rows: return
    end=2+len(rows); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title="Revenue Mix"; ch.height=8.5; ch.width=13.5; ch.legend=None
    ch.add_data(Reference(ws,min_col=40,min_row=3,max_row=end),titles_from_data=False)
    _set_text_categories(ch,ws,39,3,end,["Revenue"]); ch.x_axis.title="$bn"; ch.x_axis.numFmt="0"; ch.visible_cells_only=False; ch.display_blanks="gap"
    ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0"
    ws.add_chart(ch,"A53")


def _rebuild_margin_chart(wb,ws):
    if "Segment Analysis" not in wb.sheetnames: return
    seg=wb["Segment Analysis"]; rows,excluded,profit_label=_margin_rows(seg)
    for r in range(2,14):
        ws.cell(r,42).value=None; ws.cell(r,43).value=None
    ws["AP2"]="Segment"; ws["AQ2"]="Margin"
    for out_r,(src_r,name,margin) in enumerate(rows,3):
        ws.cell(out_r,42,name); ws.cell(out_r,43,margin); ws.cell(out_r,43).number_format=FMT_PCT
    if rows:
        end=2+len(rows); ch=BarChart(); ch.type="bar"; ch.style=10; ch.title="Latest Segment Margin"; ch.height=8.5; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=43,min_row=3,max_row=end),titles_from_data=False)
        _set_text_categories(ch,ws,42,3,end,["Margin"]); ch.x_axis.title="Operating margin"; ch.x_axis.numFmt="0%"; ch.visible_cells_only=False; ch.display_blanks="gap"
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"I53")
    note=f"Margin = {profit_label} ÷ segment revenue."
    if excluded:
        details=", ".join(f"{name} {margin:.1%}" for name,margin in excluded)
        note+=f" Outlier omitted from common chart scale but retained in Segment Analysis: {details}."
    ws["I70"]=note; ws["I70"].font=Font(size=9,color=GREY)


def polish_analysis_charts(wb,ticker=None):
    if "Analysis Charts" not in wb.sheetnames: return {"rebuilt":0}
    ws=wb["Analysis Charts"]
    _shorten_stress_helpers(ws)
    # Known final anchors: historical performance (I29), business mix (A53), segment margin (I53).
    # Remove every prior object at those anchors before rebuilding so repeated pipeline passes are idempotent.
    _remove_at(ws,{(28,8),(52,0),(52,8)})
    _rebuild_history_chart(wb,ws); _rebuild_business_chart(wb,ws); _rebuild_margin_chart(wb,ws)
    _polish_existing_chart_titles(ws)
    return {"rebuilt":3,"chart_count":len(getattr(ws,"_charts",[]) or [])}
